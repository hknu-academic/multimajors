"""
============================================================
🎯 다전공 비교 분석 모듈 (simulation.py)
============================================================
버전: 2.0 (2026.11. 정확도 점검 반영)
설명: 희망 전공을 여러 제도(복수/부전공/융합 등)로 이수할 때 학점 비교

[기준 적용 원칙]
1. 기준 학점은 엑셀 표 값을 그대로 사용합니다.
   - data/primary_requirements.xlsx    : 원전공(본전공) 기준
   - data/graduation_requirements.xlsx : 다전공 기준
2. 전공명 · 제도유형 · 입학연도 · 입학구분이 모두 일치하는 행만 사용합니다.
3. 일치하는 행이 없으면 '자료 없음'으로 안내합니다.
   (비슷한 이름의 전공 · 가까운 연도 · 기본값으로 대신 계산하지 않음)
4. 졸업학점 = 원전공 졸업학점과 다전공 졸업학점 중 큰 값
   (다전공 표의 졸업학점이 비어 있으면 원전공 졸업학점)
5. 다전공 표의 비고가 '신입생 기준'인 행은 편입생에게 적용하지 않습니다.
============================================================
"""

import html
import math
import os
import re
from dataclasses import dataclass, field
from enum import Enum
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st

# ============================================================
# 상수 정의
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PR_PATH = os.path.join(BASE_DIR, "data", "primary_requirements.xlsx")
GR_PATH = os.path.join(BASE_DIR, "data", "graduation_requirements.xlsx")
MAJORS_PATH = os.path.join(BASE_DIR, "data", "majors_info.xlsx")

MAX_CREDITS_PER_SEMESTER = 18    # 학기당 최대 수강학점 (학기별 계획 배분용)
REQUIRED_CAP_PER_SEMESTER = 6    # 학기당 전공필수 배정 상한 (마지막 학기는 상한 없음)
SAFE_MARGIN_PER_SEMESTER = 6     # 학기당 이만큼 여유가 있으면 '가능', 부족하면 '위험'

FRESHMAN_TOTAL_SEMESTERS = 8     # 신입학 총 학기
TRANSFER_TOTAL_SEMESTERS = 4     # 3학년 편입학 총 학기
FIVE_YEAR_KEYWORD = "5년제"       # 전공명에 이 글자가 있으면 5년제 (예: 건축학전공(5년제))
FIVE_YEAR_FRESHMAN_SEMESTERS = 10
FIVE_YEAR_TRANSFER_SEMESTERS = 6

# 화면에 보이는 입학구분 (엑셀은 '3학년편입학(동일계)'처럼 붙여 써도 됨 — 공백은 무시하고 비교)
ADMISSION_TYPES = ["신입학", "3학년 편입학(동일계)", "3학년 편입학(비동일계)"]

# '현재 상태(다전공 미참여)' 계산에 쓰는 원전공 표의 행
BASELINE_PROGRAM = "복수전공"

# 제도 우선순위 (숫자가 낮을수록 우선)
PROGRAM_PRIORITY = {
    "복수전공": 1,
    "융합전공": 2,
    "부전공": 3,
    "융합부전공": 4,
    "연계전공": 5,
}

# 다전공 전공필수 변경 안내 (화면 안내 문구로만 표시, 계산에는 반영하지 않음)
MAJOR_NOTICES = {
    "경영학전공": (
        "📌 **경영학전공** 다전공 전공필수 학점은 2025학년도 2학기 승인자부터 "
        "9학점 → 6학점으로 변경되었습니다. 승인 시기에 따라 적용 기준이 다를 수 있으니 "
        "소속 전공 사무실 또는 학사공지에서 확인해 주세요."
    ),
    "법학전공": (
        "📌 **법학전공** 다전공 전공필수 학점은 2025학년도 2학기 승인자부터 "
        "18학점 → 15학점으로 변경되었습니다. 승인 시기에 따라 적용 기준이 다를 수 있으니 "
        "소속 전공 사무실 또는 학사공지에서 확인해 주세요."
    ),
}

MISSING_GUIDE = "해당 입학연도 기준 자료가 없어 계산하지 않았습니다. 소속 전공 사무실 또는 학사공지에서 확인해 주세요."
RESULT_DISCLAIMER = "※ 이 결과는 엑셀 기준표를 바탕으로 한 참고용입니다. 최종 기준은 학칙·학사운영규정·학교 공지사항입니다."

# ============================================================
# 열거형 정의
# ============================================================

class AdmissionType(Enum):
    FRESHMAN = "신입학"
    TRANSFER_SAME = "3학년 편입학(동일계)"
    TRANSFER_DIFF = "3학년 편입학(비동일계)"

class StudentType(Enum):
    NEW_APPLICANT = "신규 신청자"
    CURRENT_PARTICIPANT = "기존 참여자"

class GraduationStatus(Enum):
    POSSIBLE = "가능"
    RISKY = "위험"
    IMPOSSIBLE = "어려움"
    NO_DATA = "자료 없음"


# ============================================================
# 데이터 클래스
# ============================================================

@dataclass
class StudentInput:
    """학생 입력 정보"""
    student_type: str                    # 신규 신청자 / 기존 참여자
    admission_year: int                  # 입학연도
    primary_major: str                   # 본전공명
    admission_type: str                  # 입학구분
    completed_semesters: int             # 현재까지 이수한 학기 수
    transfer_credits: int = 0            # 편입학 인정학점

    # 교양 학점 (신입학만)
    credits_basic_literacy: int = 0
    credits_basic_science: int = 0
    credits_core_liberal: int = 0

    # 본전공 학점
    credits_major_required: int = 0
    credits_major_elective: int = 0

    # 잔여 학점
    credits_free: int = 0

    # 다전공 관련 (신규 신청자)
    desired_multi_major: Optional[str] = None

    # 다전공 관련 (기존 참여자)
    current_program: Optional[str] = None
    current_multi_major: Optional[str] = None
    credits_multi_required: int = 0
    credits_multi_elective: int = 0


@dataclass
class CreditAnalysis:
    """학점 분석 결과"""
    # 기준 학점
    req_major_required: int = 0
    req_major_elective: int = 0
    req_multi_required: int = 0
    req_multi_elective: int = 0
    req_graduation_credits: int = 0
    primary_graduation_credits: int = 0           # 원전공 표의 졸업학점
    multi_graduation_credits: Optional[int] = None  # 다전공 표의 졸업학점 (없으면 None)

    # 교양 기준 학점
    req_basic_literacy: int = 0
    req_basic_science: int = 0
    req_core_liberal: int = 0

    # 이수 학점
    completed_major_required: int = 0
    completed_major_elective: int = 0
    completed_multi_required: int = 0
    completed_multi_elective: int = 0
    completed_total: int = 0

    # 교양 이수 학점
    completed_basic_literacy: int = 0
    completed_basic_science: int = 0
    completed_core_liberal: int = 0

    # 부족 학점
    deficit_major_required: int = 0
    deficit_major_elective: int = 0
    deficit_multi_required: int = 0
    deficit_multi_elective: int = 0
    deficit_graduation: int = 0

    # 교양 부족 학점
    deficit_basic_literacy: int = 0
    deficit_basic_science: int = 0
    deficit_core_liberal: int = 0

    # 합계
    required_deficit: int = 0     # 영역별 부족 학점 합계 (교양 + 본전공 + 다전공)
    needed_credits: int = 0       # 졸업까지 추가로 필요한 학점 = max(영역별 부족 합계, 졸업학점 부족)

    # 학기 정보
    total_semesters: int = 0
    remaining_semesters: int = 0
    max_additional_credits: int = 0

    # 본전공 변화 학점 (다전공 참여 시)
    req_major_required_changed: int = 0
    req_major_elective_changed: int = 0


@dataclass
class SimulationResult:
    """제도별 비교 분석 결과"""
    program_type: str
    multi_major_name: str

    can_graduate: bool = False
    graduation_status: str = "어려움"

    credit_analysis: CreditAnalysis = field(default_factory=CreditAnalysis)

    recommendation_rank: int = 0
    recommendation_reason: str = ""
    is_supplementary: bool = False

    # 기준 자료 없음
    data_missing: bool = False
    missing_reason: str = ""

    semester_plan: List[Dict] = field(default_factory=list)


@dataclass
class AnalysisOutput:
    """전체 분석 결과"""
    student_input: StudentInput = None

    current_analysis: Optional[CreditAnalysis] = None
    current_can_graduate: bool = False
    current_status: str = ""
    primary_missing_reason: str = ""

    simulation_results: List[SimulationResult] = field(default_factory=list)
    recommended_programs: List[SimulationResult] = field(default_factory=list)
    supplementary_programs: List[SimulationResult] = field(default_factory=list)


# ============================================================
# 작은 도우미 함수
# ============================================================

def _norm(value) -> str:
    """비교용 문자열: 빈 값은 '', 모든 공백 제거"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return re.sub(r"\s+", "", str(value))


def safe_int(value, default=0):
    """안전하게 정수로 변환 ('39', '39.0', 39 → 39 / 빈 값 → default)"""
    try:
        if value is None or pd.isna(value):
            return default
        return int(float(str(value).strip()))
    except (ValueError, TypeError):
        return default


def _is_divider(value) -> bool:
    return bool(value) and str(value).startswith("━━━━━")


def _year_col(df: pd.DataFrame) -> Optional[str]:
    for col in ("입학연도", "기준학번"):
        if col in df.columns:
            return col
    return None


def _first_value(row: pd.Series, patterns: List[str]) -> Optional[int]:
    """여러 열 이름 중 먼저 찾은 값 (빈칸이면 None)"""
    for col in patterns:
        if col in row.index and pd.notna(row[col]):
            return safe_int(row[col], None)
    return None


# ============================================================
# 데이터 로드 함수 (엑셀을 바꾸면 자동으로 다시 읽음)
# ============================================================

@st.cache_data(show_spinner=False)
def _read_excel_cached(path: str, mtime: float) -> pd.DataFrame:
    return pd.read_excel(path)


def _read_excel(path: str) -> pd.DataFrame:
    try:
        return _read_excel_cached(path, os.path.getmtime(path))
    except Exception as e:
        print(f"[다전공 비교 분석] 엑셀을 읽지 못했습니다: {path} → {e}")
        return pd.DataFrame()


def load_primary_requirements() -> pd.DataFrame:
    """본전공 기준 데이터 로드"""
    return _read_excel(PR_PATH)


def load_graduation_requirements() -> pd.DataFrame:
    """다전공 기준 데이터 로드"""
    return _read_excel(GR_PATH)


def load_majors_info() -> pd.DataFrame:
    """전공 정보 로드"""
    return _read_excel(MAJORS_PATH)


def load_majors_list() -> List[str]:
    """전공 목록 로드"""
    df = load_majors_info()
    if df.empty or "전공명" not in df.columns:
        return []
    return sorted(df["전공명"].dropna().unique().tolist())


def load_multi_majors_by_program(program_type: str) -> List[str]:
    """제도별 다전공 목록 로드"""
    gr_df = load_graduation_requirements()
    if gr_df.empty or "제도유형" not in gr_df.columns:
        return []
    return sorted(gr_df[gr_df["제도유형"] == program_type]["전공명"].dropna().unique().tolist())


def available_admission_years() -> List[int]:
    """입학연도 선택지: 원전공 기준표에 있는 연도 (최신순)"""
    pr_df = load_primary_requirements()
    col = _year_col(pr_df)
    if col:
        years = {safe_int(v, None) for v in pr_df[col]}
        years.discard(None)
        if years:
            return sorted(years, reverse=True)
    return list(range(2026, 2019, -1))


# ============================================================
# 기준 조회 (정확히 일치하는 행만 사용, 없으면 '자료 없음')
# ============================================================

def _match_rows(df: pd.DataFrame, major: str, program_type: str, year: int) -> pd.DataFrame:
    """전공명 · 제도유형 · 입학연도가 모두 일치하는 행"""
    yc = _year_col(df)
    if df.empty or yc is None or not {"전공명", "제도유형"} <= set(df.columns):
        return df.iloc[0:0]
    mask = (
        (df["전공명"].map(_norm) == _norm(major)) &
        (df["제도유형"].map(_norm) == _norm(program_type)) &
        (df[yc].map(lambda v: safe_int(v, None)) == int(year))
    )
    return df[mask]


def _note_applies(note, admission_type: str) -> bool:
    """다전공 표 비고가 이 학생에게 적용되는지 ('신입생 기준'은 편입생에게 적용 안 함)"""
    n = _norm(note)
    is_transfer = _norm(admission_type) != _norm("신입학")
    if is_transfer and "신입생기준" in n:
        return False
    if not is_transfer and n.startswith("편입생"):
        return False
    return True


def get_primary_requirement(
    primary_major: str,
    program_type: str,
    admission_year: int,
    admission_type: str,
    pr_df: pd.DataFrame,
    show_program: bool = True
) -> Tuple[Optional[Dict], str]:
    """원전공 기준 조회 → (기준, '') 또는 (None, 자료 없음 사유)"""
    program_part = f" · {program_type}" if show_program else ""
    label = f"원전공 기준({primary_major}{program_part} · {admission_year}학번 · {admission_type})"
    if pr_df.empty:
        return None, "원전공 기준표(primary_requirements.xlsx)를 읽지 못했습니다."
    if not primary_major:
        return None, "원전공이 선택되지 않았습니다."

    rows = _match_rows(pr_df, primary_major, program_type, admission_year)
    if "입학구분" in rows.columns:
        rows = rows[rows["입학구분"].map(_norm) == _norm(admission_type)]
    if rows.empty:
        return None, f"{label} 자료 없음"
    if len(rows) > 1:
        print(f"[다전공 비교 분석] 중복 행: {label} → 첫 번째 행 사용")

    row = rows.iloc[0]
    req = {
        "major_name": row["전공명"],
        "program_type": row["제도유형"],
        "admission_year": admission_year,
        "req_major_required": _first_value(row, ["본전공_전공필수", "본전공 전공필수"]),
        "req_major_elective": _first_value(row, ["본전공_전공선택", "본전공 전공선택"]),
        "req_total": _first_value(row, ["본전공_계", "본전공 계"]),
        "req_major_required_changed": _first_value(row, ["본전공변화_전공필수", "본전공변화 전공필수"]),
        "req_major_elective_changed": _first_value(row, ["본전공변화_전공선택", "본전공변화 전공선택"]),
        "req_basic_literacy": _first_value(row, ["기초교양_기초문해", "기초교양(기초문해)", "기초문해"]),
        "req_basic_science": _first_value(row, ["기초교양_기초과학", "기초교양(기초과학)", "기초과학"]),
        "req_core_liberal": _first_value(row, ["핵심교양", "핵심 교양"]),
        "req_graduation_credits": _first_value(row, ["졸업학점", "졸업 학점"]),
    }

    must = ["req_major_required", "req_major_elective",
            "req_major_required_changed", "req_major_elective_changed",
            "req_graduation_credits"]
    if any(req[k] is None for k in must):
        return None, f"{label} 표에 빈칸이 있어 자료 없음"

    # 교양 빈칸은 '해당 없음(0)'으로 처리
    for k in ("req_basic_literacy", "req_basic_science", "req_core_liberal"):
        if req[k] is None:
            req[k] = 0
    return req, ""


def get_graduation_requirement(
    multi_major: str,
    program_type: str,
    admission_year: int,
    admission_type: str,
    gr_df: pd.DataFrame
) -> Tuple[Optional[Dict], str]:
    """다전공 기준 조회 → (기준, '') 또는 (None, 자료 없음 사유)"""
    label = f"다전공 기준({multi_major} · {program_type} · {admission_year}학번)"
    if gr_df.empty:
        return None, "다전공 기준표(graduation_requirements.xlsx)를 읽지 못했습니다."
    if not multi_major:
        return None, "다전공이 선택되지 않았습니다."

    rows = _match_rows(gr_df, multi_major, program_type, admission_year)
    if "입학구분" in rows.columns:          # 나중에 입학구분 열을 추가하는 경우 대비
        rows = rows[rows["입학구분"].map(_norm) == _norm(admission_type)]
    if rows.empty:
        return None, f"{label} 자료 없음"

    if "비고" in rows.columns:
        applicable = rows[rows["비고"].map(lambda v: _note_applies(v, admission_type))]
        if applicable.empty:
            return None, f"{label}은 '신입생 기준'만 있어 편입생 기준 자료 없음"
        rows = applicable

    row = rows.iloc[0]
    req_r = _first_value(row, ["다전공_전공필수", "다전공 전공필수"])
    req_e = _first_value(row, ["다전공_전공선택", "다전공 전공선택"])
    if req_r is None or req_e is None:
        return None, f"{label} 표에 빈칸이 있어 자료 없음"

    total = _first_value(row, ["다전공_계", "다전공 계"])
    return {
        "major_name": row["전공명"],
        "program_type": row["제도유형"],
        "admission_year": admission_year,
        "req_multi_required": req_r,
        "req_multi_elective": req_e,
        "req_total": total if total is not None else req_r + req_e,
        "req_graduation_credits": _first_value(row, ["졸업학점", "졸업 학점"]),  # 없으면 None
    }, ""


# ============================================================
# 계산 함수
# ============================================================

def _is_five_year(primary_major: str) -> bool:
    return FIVE_YEAR_KEYWORD in str(primary_major or "")


def get_total_semesters(admission_type: str, primary_major: str = "") -> int:
    """총 학기 수 (5년제 전공 반영)"""
    five = _is_five_year(primary_major)
    if admission_type == "신입학":
        return FIVE_YEAR_FRESHMAN_SEMESTERS if five else FRESHMAN_TOTAL_SEMESTERS
    return FIVE_YEAR_TRANSFER_SEMESTERS if five else TRANSFER_TOTAL_SEMESTERS


def calculate_remaining_semesters(admission_type: str, completed_semesters: int,
                                  primary_major: str = "") -> int:
    """남은 학기 계산"""
    return max(0, get_total_semesters(admission_type, primary_major) - completed_semesters)


def calculate_max_additional_credits(remaining_semesters: int) -> int:
    """최대 추가 이수 가능 학점"""
    return remaining_semesters * MAX_CREDITS_PER_SEMESTER


def apply_excess_to_elective(completed_required: int, req_required: int,
                             completed_elective: int) -> Tuple[int, int]:
    """전공필수 초과분 → 전공선택 이월"""
    if completed_required > req_required:
        excess = completed_required - req_required
        return req_required, completed_elective + excess
    return completed_required, completed_elective


def calculate_deficit(completed: int, required: int) -> int:
    """부족 학점 계산"""
    return max(0, required - completed)


def calculate_graduation_credits(primary_grad_credits: int,
                                 multi_grad_credits: Optional[int] = None) -> int:
    """졸업학점 = 원전공 졸업학점과 다전공 졸업학점 중 큰 값 (다전공 값이 없으면 원전공)"""
    if multi_grad_credits is None:
        return primary_grad_credits
    return max(primary_grad_credits, multi_grad_credits)


def calculate_completed_total(student: StudentInput, include_multi: bool) -> int:
    """총 이수 학점 (편입생은 편입 인정학점 포함)"""
    total = (student.credits_major_required + student.credits_major_elective +
             student.credits_free)
    if include_multi:
        total += student.credits_multi_required + student.credits_multi_elective
    if student.admission_type == "신입학":
        total += (student.credits_basic_literacy + student.credits_basic_science +
                  student.credits_core_liberal)
    else:
        total += student.transfer_credits
    return total


def _set_liberal(a: CreditAnalysis, pr_req: Dict, student: StudentInput):
    """교양 기준 · 이수 · 부족"""
    a.req_basic_literacy = pr_req["req_basic_literacy"]
    a.req_basic_science = pr_req["req_basic_science"]
    a.req_core_liberal = pr_req["req_core_liberal"]
    a.completed_basic_literacy = student.credits_basic_literacy
    a.completed_basic_science = student.credits_basic_science
    a.completed_core_liberal = student.credits_core_liberal
    a.deficit_basic_literacy = calculate_deficit(student.credits_basic_literacy, a.req_basic_literacy)
    a.deficit_basic_science = calculate_deficit(student.credits_basic_science, a.req_basic_science)
    a.deficit_core_liberal = calculate_deficit(student.credits_core_liberal, a.req_core_liberal)


def _set_semesters(a: CreditAnalysis, student: StudentInput):
    a.total_semesters = get_total_semesters(student.admission_type, student.primary_major)
    a.remaining_semesters = calculate_remaining_semesters(
        student.admission_type, student.completed_semesters, student.primary_major)
    a.max_additional_credits = calculate_max_additional_credits(a.remaining_semesters)


def _finish_analysis(a: CreditAnalysis):
    """영역별 부족 합계 · 졸업학점 부족 · 졸업까지 필요한 학점"""
    a.required_deficit = (
        a.deficit_basic_literacy + a.deficit_basic_science + a.deficit_core_liberal +
        a.deficit_major_required + a.deficit_major_elective +
        a.deficit_multi_required + a.deficit_multi_elective
    )
    a.deficit_graduation = calculate_deficit(a.completed_total, a.req_graduation_credits)
    a.needed_credits = max(a.required_deficit, a.deficit_graduation)


def determine_graduation_status(needed_credits: int, max_additional: int,
                                remaining_semesters: int) -> Tuple[str, bool]:
    """이수 가능 상태 판단 (영역별 부족 · 졸업학점 부족 모두 반영한 needed_credits 기준)"""
    if needed_credits <= 0:
        return "가능", True
    if remaining_semesters <= 0 or needed_credits > max_additional:
        return "어려움", False
    margin = max_additional - needed_credits
    if margin >= remaining_semesters * SAFE_MARGIN_PER_SEMESTER:
        return "가능", True
    return "위험", True


# ============================================================
# 분석 함수
# ============================================================

def analyze_current_status(student: StudentInput,
                           pr_df: pd.DataFrame) -> Tuple[Optional[CreditAnalysis], str]:
    """현재 상태 분석 (본전공 기준, 다전공 미참여 시)"""
    pr_req, reason = get_primary_requirement(
        student.primary_major, BASELINE_PROGRAM, student.admission_year,
        student.admission_type, pr_df, show_program=False)
    
    if pr_req is None:
        return None, reason

    a = CreditAnalysis()
    _set_semesters(a, student)

    a.req_major_required = pr_req["req_major_required"]
    a.req_major_elective = pr_req["req_major_elective"]
    a.primary_graduation_credits = pr_req["req_graduation_credits"]
    a.req_graduation_credits = pr_req["req_graduation_credits"]
    _set_liberal(a, pr_req, student)

    a.completed_major_required = student.credits_major_required
    a.completed_major_elective = student.credits_major_elective
    adj_req, adj_elec = apply_excess_to_elective(
        a.completed_major_required, a.req_major_required, a.completed_major_elective)
    a.deficit_major_required = calculate_deficit(adj_req, a.req_major_required)
    a.deficit_major_elective = calculate_deficit(adj_elec, a.req_major_elective)

    a.completed_total = calculate_completed_total(student, include_multi=False)
    _finish_analysis(a)
    return a, ""


def simulate_program(
    student: StudentInput,
    program_type: str,
    multi_major: str,
    pr_df: pd.DataFrame,
    gr_df: pd.DataFrame
) -> SimulationResult:
    """단일 제도 분석 (기준 자료가 없으면 data_missing=True)"""
    result = SimulationResult(program_type=program_type, multi_major_name=multi_major)

    pr_req, r1 = get_primary_requirement(
        student.primary_major, program_type, student.admission_year,
        student.admission_type, pr_df)
    gr_req, r2 = get_graduation_requirement(
        multi_major, program_type, student.admission_year,
        student.admission_type, gr_df)

    if pr_req is None or gr_req is None:
        result.data_missing = True
        result.missing_reason = " / ".join(r for r in (r1, r2) if r)
        result.graduation_status = "자료 없음"
        result.can_graduate = False
        return result

    a = CreditAnalysis()
    _set_semesters(a, student)

    # 본전공 기준 (다전공 참여 시 변화된 기준)
    a.req_major_required = pr_req["req_major_required_changed"]
    a.req_major_elective = pr_req["req_major_elective_changed"]
    a.req_major_required_changed = pr_req["req_major_required_changed"]
    a.req_major_elective_changed = pr_req["req_major_elective_changed"]
    _set_liberal(a, pr_req, student)

    # 다전공 기준
    a.req_multi_required = gr_req["req_multi_required"]
    a.req_multi_elective = gr_req["req_multi_elective"]

    # 졸업학점: 원전공과 다전공 중 큰 값
    a.primary_graduation_credits = pr_req["req_graduation_credits"]
    a.multi_graduation_credits = gr_req["req_graduation_credits"]
    a.req_graduation_credits = calculate_graduation_credits(
        a.primary_graduation_credits, a.multi_graduation_credits)

    # 본전공 이수 · 부족
    a.completed_major_required = student.credits_major_required
    a.completed_major_elective = student.credits_major_elective
    adj_req, adj_elec = apply_excess_to_elective(
        a.completed_major_required, a.req_major_required, a.completed_major_elective)
    a.deficit_major_required = calculate_deficit(adj_req, a.req_major_required)
    a.deficit_major_elective = calculate_deficit(adj_elec, a.req_major_elective)

    # 다전공 이수 · 부족 (신규 신청자는 0)
    is_participant = student.student_type == "기존 참여자"
    a.completed_multi_required = student.credits_multi_required if is_participant else 0
    a.completed_multi_elective = student.credits_multi_elective if is_participant else 0
    m_req, m_elec = apply_excess_to_elective(
        a.completed_multi_required, a.req_multi_required, a.completed_multi_elective)
    a.deficit_multi_required = calculate_deficit(m_req, a.req_multi_required)
    a.deficit_multi_elective = calculate_deficit(m_elec, a.req_multi_elective)

    # 총 이수 · 졸업까지 필요한 학점
    a.completed_total = calculate_completed_total(student, include_multi=is_participant)
    _finish_analysis(a)

    status, can_grad = determine_graduation_status(
        a.needed_credits, a.max_additional_credits, a.remaining_semesters)
    result.graduation_status = status
    result.can_graduate = can_grad
    result.credit_analysis = a
    return result


def generate_semester_plan(analysis: CreditAnalysis, student: StudentInput) -> List[Dict]:
    """학기별 이수 계획 생성 - 학년/학기 표기 및 우선순위 반영"""
    plan = []
    if analysis.remaining_semesters <= 0:
        return plan

    required_deficit = analysis.required_deficit
    graduation_deficit = analysis.deficit_graduation
    total_remaining = max(required_deficit, graduation_deficit)
    free_deficit = max(0, graduation_deficit - required_deficit)

    remaining = {
        "basic_literacy": analysis.deficit_basic_literacy,
        "basic_science": analysis.deficit_basic_science,
        "core_liberal": analysis.deficit_core_liberal,
        "major_required": analysis.deficit_major_required,
        "major_elective": analysis.deficit_major_elective,
        "multi_required": analysis.deficit_multi_required,
        "multi_elective": analysis.deficit_multi_elective,
        "free": free_deficit,
    }
    liberal_first = ["basic_literacy", "basic_science", "core_liberal",
                     "major_required", "multi_required",
                     "major_elective", "multi_elective", "free"]
    major_first = ["major_required", "multi_required",
                   "major_elective", "multi_elective",
                   "basic_literacy", "basic_science", "core_liberal", "free"]

    current_semester = student.completed_semesters

    for sem_idx in range(1, analysis.remaining_semesters + 1):
        absolute_semester = current_semester + sem_idx
        if student.admission_type == "신입학":
            grade = (absolute_semester + 1) // 2
        else:
            grade = 3 + (absolute_semester - 1) // 2      # 3학년 편입
        semester = 1 if absolute_semester % 2 == 1 else 2

        sem_plan = {
            "semester": f"{grade}학년 {semester}학기",
            "basic_literacy": 0,
            "basic_science": 0,
            "core_liberal": 0,
            "major_required": 0,
            "major_elective": 0,
            "multi_required": 0,
            "multi_elective": 0,
            "free": 0,
            "total": 0,
        }

        semesters_left = analysis.remaining_semesters - sem_idx + 1
        is_last = semesters_left == 1
        credits_this_semester = min(
            MAX_CREDITS_PER_SEMESTER,
            math.ceil(total_remaining / semesters_left) if total_remaining > 0 else 0,
            total_remaining,
        )
        left = credits_this_semester

        # 1학년 · 4학년 이상: 교양 우선 / 2~3학년: 전공 우선
        order = liberal_first if (grade == 1 or grade >= 4) else major_first
        for key in order:
            if left <= 0:
                break
            if remaining[key] <= 0:
                continue
            cap = left
            if key in ("major_required", "multi_required") and not is_last:
                cap = min(cap, REQUIRED_CAP_PER_SEMESTER)
            take = min(remaining[key], cap)
            sem_plan[key] = take
            remaining[key] -= take
            left -= take

        sem_plan["total"] = sum(sem_plan[k] for k in remaining)
        total_remaining -= sem_plan["total"]

        if sem_plan["total"] > 0:
            plan.append(sem_plan)

    return plan


def rank_recommendations(results: List[SimulationResult]) -> Tuple[List[SimulationResult], List[SimulationResult]]:
    """추천 순위 정렬 (자료 없음 제도는 순위에서 제외하고 뒤에 표시)"""

    def get_score(r: SimulationResult) -> Tuple:
        grad_score = {"가능": 0, "위험": 1, "어려움": 2}.get(r.graduation_status, 2)
        deficit_score = r.credit_analysis.needed_credits // 3   # ±3학점은 같은 수준으로
        priority_score = PROGRAM_PRIORITY.get(r.program_type, 5)
        return (grad_score, deficit_score, priority_score)

    ranked = [r for r in results if not r.data_missing]
    missing = [r for r in results if r.data_missing]
    ranked.sort(key=get_score)

    for idx, r in enumerate(ranked):
        r.recommendation_rank = idx + 1
        r.recommendation_reason = generate_recommendation_reason(r, idx + 1)
    for r in missing:
        r.recommendation_rank = 0
        r.is_supplementary = True
        r.recommendation_reason = r.missing_reason

    return ranked, missing


def generate_recommendation_reason(result: SimulationResult, rank: int) -> str:
    """추천 사유 생성"""
    a = result.credit_analysis
    reasons = []

    if result.graduation_status == "가능":
        reasons.append(f"남은 {a.remaining_semesters}학기 내 이수 가능")
    elif result.graduation_status == "위험":
        reasons.append("학기당 집중 이수 시 이수 가능")
    else:
        reasons.append("남은 학기 안에 이수하기 어려움")

    reasons.append(f"졸업까지 {a.needed_credits}학점 추가 이수 필요")

    if result.program_type == "복수전공":
        reasons.append("학위 2개 취득 가능")
    elif result.program_type == "부전공":
        reasons.append("비교적 적은 학점으로 이수 가능")
    elif result.program_type == "융합전공":
        reasons.append("융합적 역량 강화")
    elif result.program_type == "융합부전공":
        reasons.append("적은 학점으로 융합 역량 확보")
    elif result.program_type == "연계전공":
        reasons.append("다양한 학문 간 연계 학습")

    return " / ".join(reasons)


def _programs_for(desired_multi_major: str) -> List[str]:
    """희망 전공이 융합전공이면 융합 계열 제도, 아니면 일반 제도"""
    info = load_majors_info()
    is_convergence = False
    if not info.empty and {"전공명", "제도유형"} <= set(info.columns):
        sel = info[info["전공명"] == desired_multi_major]
        if not sel.empty and "융합전공" in str(sel.iloc[0]["제도유형"]):
            is_convergence = True
    if is_convergence:
        return ["융합전공", "융합부전공", "연계전공"]
    return ["복수전공", "부전공", "연계전공"]


# ============================================================
# 메인 비교 분석 함수
# ============================================================

def run_simulation(student: StudentInput) -> AnalysisOutput:
    """통합 비교 분석 실행"""
    output = AnalysisOutput()
    output.student_input = student

    pr_df = load_primary_requirements()
    gr_df = load_graduation_requirements()

    # 현재 상태 분석
    current, reason = analyze_current_status(student, pr_df)
    output.current_analysis = current
    output.primary_missing_reason = reason
    if current is not None:
        output.current_status, output.current_can_graduate = determine_graduation_status(
            current.needed_credits, current.max_additional_credits, current.remaining_semesters)

    if student.student_type == "신규 신청자" and student.desired_multi_major:
        for program in _programs_for(student.desired_multi_major):
            result = simulate_program(student, program, student.desired_multi_major, pr_df, gr_df)
            if not result.data_missing:
                result.semester_plan = generate_semester_plan(result.credit_analysis, student)
            output.simulation_results.append(result)

        output.recommended_programs, output.supplementary_programs = rank_recommendations(
            list(output.simulation_results))

    elif student.student_type == "기존 참여자" and student.current_program and student.current_multi_major:
        result = simulate_program(student, student.current_program,
                                  student.current_multi_major, pr_df, gr_df)
        if not result.data_missing:
            result.semester_plan = generate_semester_plan(result.credit_analysis, student)
        output.simulation_results.append(result)

    return output


# ============================================================
# Streamlit UI 도우미
# ============================================================

def _grouped_options(df: pd.DataFrame) -> List[str]:
    """계열별 구분선 + 전공명 목록"""
    if df is None or df.empty or "전공명" not in df.columns:
        return []
    if "계열" not in df.columns:
        return sorted(df["전공명"].dropna().unique().tolist())
    df = df.assign(계열=df["계열"].fillna("기타"))
    options = []
    for category in sorted(df["계열"].unique()):
        options.append(f"━━━━━ 📚 {category} ━━━━━")
        options.extend(sorted(df[df["계열"] == category]["전공명"].dropna().unique().tolist()))
    return options


def _render_major_notices(student: StudentInput):
    """경영학 · 법학 다전공 전공필수 변경 안내"""
    targets = {student.desired_multi_major, student.current_multi_major}
    for major, text in MAJOR_NOTICES.items():
        if major in targets:
            st.info(text)


def _graduation_line(a: CreditAnalysis, program_type: str) -> str:
    if a.multi_graduation_credits is None:
        return f"졸업학점 <b>{a.req_graduation_credits}학점</b> (원전공 기준)"
    return (f"졸업학점 <b>{a.req_graduation_credits}학점</b> "
            f"(원전공 {a.primary_graduation_credits} · {program_type} {a.multi_graduation_credits} 중 큰 값)")


def _missing_box(title: str, reason: str):
    st.markdown(f"""
<div style="background: #fff8f0; border-radius: 12px; padding: 20px; margin-bottom: 15px;
border-left: 4px solid #f0ad4e; box-shadow: 0 2px 10px rgba(0,0,0,0.05);">
<span style="font-size: 1.3rem; font-weight: bold; color: #333;">{html.escape(title)}</span>
<span style="background: #f0ad4e; color: white; padding: 3px 10px; border-radius: 15px;
font-size: 0.8rem; margin-left: 10px;">자료 없음</span>
<p style="color: #666; margin: 12px 0 0 0;">{html.escape(reason)}<br>{MISSING_GUIDE}</p>
</div>
""", unsafe_allow_html=True)


# ============================================================
# Streamlit UI 함수
# ============================================================

def render_simulation_page():
    """다전공 비교 분석 페이지"""

    st.markdown("""
    <p style="text-align: center; color: #667eea; margin-bottom: 10px; font-size: 2rem; font-weight: 600;">
        🎯 다전공 비교 분석
    </p>
    <p style="text-align: center; color: #666; margin-bottom: 30px;">
        희망 전공을 여러 제도로 이수할 때 필요한 학점을 비교해보세요!
    </p>
    """, unsafe_allow_html=True)

    if 'sim_step' not in st.session_state:
        st.session_state.sim_step = 1

    col1, col2, col3, col4 = st.columns(4)
    steps = [
        ("1️⃣", "유형 선택"),
        ("2️⃣", "기본 정보"),
        ("3️⃣", "학점 입력"),
        ("4️⃣", "결과 확인")
    ]

    for idx, (col, (emoji, label)) in enumerate(zip([col1, col2, col3, col4], steps)):
        with col:
            if idx + 1 == st.session_state.sim_step:
                st.markdown(f"""
                <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                            color: white; padding: 10px; border-radius: 10px; text-align: center;">
                    <div style="font-size: 1.5rem;">{emoji}</div>
                    <div style="font-size: 0.85rem; font-weight: bold;">{label}</div>
                </div>
                """, unsafe_allow_html=True)
            elif idx + 1 < st.session_state.sim_step:
                st.markdown(f"""
                <div style="background: #28a745; color: white; padding: 10px; 
                            border-radius: 10px; text-align: center;">
                    <div style="font-size: 1.5rem;">✅</div>
                    <div style="font-size: 0.85rem;">{label}</div>
                </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown(f"""
                <div style="background: #e9ecef; color: #666; padding: 10px; 
                            border-radius: 10px; text-align: center;">
                    <div style="font-size: 1.5rem;">{emoji}</div>
                    <div style="font-size: 0.85rem;">{label}</div>
                </div>
                """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if st.session_state.sim_step == 1:
        render_step1_student_type()
    elif st.session_state.sim_step == 2:
        render_step2_basic_info()
    elif st.session_state.sim_step == 3:
        render_step3_credits()
    elif st.session_state.sim_step == 4:
        render_step4_results()


def render_step1_student_type():
    """STEP 1: 학생 유형 선택"""

    st.markdown("""
    <div style="background: white; border-radius: 15px; padding: 30px; 
                box-shadow: 0 4px 15px rgba(0,0,0,0.1); margin-bottom: 20px;">
        <p style="color: #333; margin-bottom: 20px; font-size: 1.3rem; font-weight: 600;">📋 어떤 상황인가요?</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)

    with col1:
        if st.button("🆕 다전공을 새로 신청하려고 해요", use_container_width=True, type="primary"):
            st.session_state.sim_student_type = "신규 신청자"
            st.session_state.sim_step = 2
            st.rerun()
        st.markdown("""
        <p style="color: #666; font-size: 0.9rem; text-align: center; margin-top: 10px;">
            아직 다전공에 참여하지 않았고,<br>어떤 제도가 좋을지 알고 싶어요
        </p>
        """, unsafe_allow_html=True)

    with col2:
        if st.button("📚 이미 다전공을 하고 있어요", use_container_width=True):
            st.session_state.sim_student_type = "기존 참여자"
            st.session_state.sim_step = 2
            st.rerun()
        st.markdown("""
        <p style="color: #666; font-size: 0.9rem; text-align: center; margin-top: 10px;">
            이미 다전공에 참여 중이고,<br>남은 학점을 확인하고 싶어요
        </p>
        """, unsafe_allow_html=True)


def render_step2_basic_info():
    """STEP 2: 기본 정보 입력"""

    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%); 
                border-radius: 15px; padding: 20px; margin-bottom: 20px;">
        <p style="color: #667eea; margin: 0; font-size: 1.3rem; font-weight: 600;">
            {'🆕 신규 신청자' if st.session_state.sim_student_type == '신규 신청자' else '📚 기존 참여자'} - 기본 정보
        </p>
    </div>
    """, unsafe_allow_html=True)

    majors_info_df = load_majors_info()
    gr_df = load_graduation_requirements()

    # 본전공 목록 (융합전공 제외, 계열별)
    if not majors_info_df.empty and "제도유형" in majors_info_df.columns:
        primary_df = majors_info_df[
            ~majors_info_df["제도유형"].astype(str).str.contains("융합전공", na=False)]
    else:
        primary_df = majors_info_df
    primary_majors_options = _grouped_options(primary_df)

    if not primary_majors_options:
        st.error("전공 목록(data/majors_info.xlsx)을 불러오지 못했습니다. 관리자에게 문의해 주세요.")
        if st.button("⬅️ 이전", use_container_width=True):
            st.session_state.sim_step = 1
            st.rerun()
        return

    col1, col2 = st.columns(2)

    with col1:
        admission_year = st.selectbox(
            "📅 입학연도",
            options=available_admission_years(),
            help="학번 기준 연도를 선택하세요"
        )
        primary_major = st.selectbox(
            "🎓 본전공",
            options=primary_majors_options,
            help="현재 소속된 전공을 선택하세요"
        )
        if _is_divider(primary_major):
            primary_major = None

    with col2:
        admission_type = st.selectbox("📝 입학구분", options=ADMISSION_TYPES)
        max_sem = get_total_semesters(admission_type, primary_major or "")
        completed_semesters = st.selectbox(
            "📆 현재까지 이수한 학기 수",
            options=list(range(1, max_sem + 1)),
            help="휴학 학기 제외"
        )

    # 편입생 인정학점
    transfer_credits = 0
    if admission_type != "신입학":
        st.markdown("---")
        transfer_credits = st.number_input(
            "🔢 편입학 인정학점",
            min_value=0, max_value=100, value=65,
            help="편입 시 인정받은 학점을 입력하세요"
        )

    # 다전공 정보 (신규 신청자)
    desired_multi_major = None
    if st.session_state.sim_student_type == "신규 신청자":
        st.markdown("---")
        st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">🎯 희망하는 다전공</p>', unsafe_allow_html=True)

        multi_majors_options = []
        if not majors_info_df.empty and not gr_df.empty and "제도유형" in majors_info_df.columns:
            double_majors = gr_df[gr_df["제도유형"] == "복수전공"]["전공명"].unique()
            available = majors_info_df[
                (majors_info_df["전공명"].isin(double_majors)) |
                (majors_info_df["제도유형"].astype(str).str.contains("융합전공", na=False))
            ]
            multi_majors_options = _grouped_options(available)
        if not multi_majors_options:
            multi_majors_options = load_multi_majors_by_program("복수전공")

        desired_multi_major = st.selectbox(
            "다전공으로 이수하고 싶은 전공",
            options=multi_majors_options,
            help="희망하는 다전공을 선택하세요"
        )
        if _is_divider(desired_multi_major):
            desired_multi_major = None

    # 기존 참여자 정보
    current_program = None
    current_multi_major = None
    if st.session_state.sim_student_type == "기존 참여자":
        st.markdown("---")
        st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📚 현재 참여 중인 다전공</p>', unsafe_allow_html=True)

        col1, col2 = st.columns(2)
        with col1:
            current_program = st.selectbox(
                "참여 중인 제도",
                options=["복수전공", "부전공", "융합전공", "융합부전공", "연계전공"]
            )

        with col2:
            options = []
            if not majors_info_df.empty and not gr_df.empty and "제도유형" in majors_info_df.columns:
                program_majors = gr_df[gr_df["제도유형"] == current_program]["전공명"].unique()
                available = majors_info_df[
                    (majors_info_df["전공명"].isin(program_majors)) |
                    (majors_info_df["제도유형"].astype(str).str.contains(current_program, na=False))
                ]
                options = _grouped_options(available)
            if not options:
                options = load_multi_majors_by_program(current_program)

            current_multi_major = st.selectbox("참여 중인 다전공명", options=options)
            if _is_divider(current_multi_major):
                current_multi_major = None

    # 세션에 저장
    st.session_state.sim_admission_year = admission_year
    st.session_state.sim_primary_major = primary_major
    st.session_state.sim_admission_type = admission_type
    st.session_state.sim_completed_semesters = completed_semesters
    st.session_state.sim_transfer_credits = transfer_credits
    st.session_state.sim_desired_multi_major = desired_multi_major
    st.session_state.sim_current_program = current_program
    st.session_state.sim_current_multi_major = current_multi_major

    # 네비게이션 버튼
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])

    with col1:
        if st.button("⬅️ 이전", use_container_width=True):
            st.session_state.sim_step = 1
            st.rerun()

    with col3:
        if st.button("다음 ➡️", use_container_width=True, type="primary"):
            if not primary_major:
                st.warning("본전공을 선택해 주세요.")
            elif st.session_state.sim_student_type == "신규 신청자" and not desired_multi_major:
                st.warning("희망하는 다전공을 선택해 주세요.")
            elif st.session_state.sim_student_type == "기존 참여자" and not current_multi_major:
                st.warning("참여 중인 다전공명을 선택해 주세요.")
            else:
                st.session_state.sim_step = 3
                st.rerun()


def render_step3_credits():
    """STEP 3: 학점 입력"""

    st.markdown("""
    <div style="background: linear-gradient(135deg, #667eea15 0%, #764ba215 100%); 
                border-radius: 15px; padding: 20px; margin-bottom: 20px;">
        <p style="color: #667eea; margin: 0; font-size: 1.3rem; font-weight: 600;">📊 현재 이수 학점 입력</p>
        <p style="color: #666; margin: 10px 0 0 0; font-size: 0.9rem;">
            정확한 분석을 위해 현재까지 이수한 학점을 입력해주세요. 이수한 학점은 <a href="https://info.hknu.ac.kr" target="_blank" style="color: #667eea; text-decoration: underline;">학사시스템</a>의 '통합학적부 조회-성적이력'에서 확인 가능합니다.
        </p>
    </div>
    """, unsafe_allow_html=True)

    credits_basic_literacy = 0
    credits_basic_science = 0
    credits_core_liberal = 0

    if st.session_state.sim_admission_type == "신입학":
        st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📚 교양 이수 학점</p>', unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            credits_basic_literacy = st.number_input(
                "기초교양(기초문해)", min_value=0, max_value=60, value=0,
                help="기초문해 영역 이수 학점")
        with col2:
            credits_basic_science = st.number_input(
                "기초교양(기초과학)", min_value=0, max_value=60, value=0,
                help="기초과학 영역 이수 학점 (해당 시)")
        with col3:
            credits_core_liberal = st.number_input(
                "핵심교양", min_value=0, max_value=60, value=0,
                help="핵심교양 영역 이수 학점")

    st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">🎓 본전공 이수 학점</p>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        credits_major_required = st.number_input(
            "전공필수 학점", min_value=0, max_value=200, value=0,
            help="본전공 전공필수 이수 학점")
    with col2:
        credits_major_elective = st.number_input(
            "전공선택 학점", min_value=0, max_value=200, value=0,
            help="본전공 전공선택 이수 학점")

    credits_multi_required = 0
    credits_multi_elective = 0

    if st.session_state.sim_student_type == "기존 참여자":
        st.markdown(f'<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📘 다전공 이수 학점 ({st.session_state.sim_current_program})</p>', unsafe_allow_html=True)
        col1, col2 = st.columns(2)
        with col1:
            credits_multi_required = st.number_input(
                "다전공 전공필수 학점", min_value=0, max_value=100, value=0,
                help="다전공 전공필수 이수 학점")
        with col2:
            credits_multi_elective = st.number_input(
                "다전공 전공선택 학점", min_value=0, max_value=100, value=0,
                help="다전공 전공선택 이수 학점")

    st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📋 기타 이수 학점</p>', unsafe_allow_html=True)
    credits_free = st.number_input(
        "잔여(자유) 학점", min_value=0, max_value=150, value=0,
        help="소양교양, 자유선택 등 기타 이수 학점")

    if st.session_state.sim_admission_type == "신입학":
        total = (credits_basic_literacy + credits_basic_science + credits_core_liberal +
                 credits_major_required + credits_major_elective +
                 credits_multi_required + credits_multi_elective + credits_free)
    else:
        total = (st.session_state.sim_transfer_credits +
                 credits_major_required + credits_major_elective +
                 credits_multi_required + credits_multi_elective + credits_free)

    st.markdown(f"""
    <div style="background: #e3f2fd; border-radius: 10px; padding: 15px; margin-top: 20px;">
        <p style="color: #1565c0; margin: 0; font-size: 1.1rem; font-weight: 600;">📊 총 이수 학점: <span style="font-size: 1.5rem;">{total}</span>학점</p>
    </div>
    """, unsafe_allow_html=True)

    st.session_state.sim_credits_basic_literacy = credits_basic_literacy
    st.session_state.sim_credits_basic_science = credits_basic_science
    st.session_state.sim_credits_core_liberal = credits_core_liberal
    st.session_state.sim_credits_major_required = credits_major_required
    st.session_state.sim_credits_major_elective = credits_major_elective
    st.session_state.sim_credits_multi_required = credits_multi_required
    st.session_state.sim_credits_multi_elective = credits_multi_elective
    st.session_state.sim_credits_free = credits_free

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        if st.button("⬅️ 이전", use_container_width=True):
            st.session_state.sim_step = 2
            st.rerun()
    with col3:
        if st.button("🔍 분석하기", use_container_width=True, type="primary"):
            st.session_state.sim_step = 4
            st.rerun()


def _render_current_status(student: StudentInput, output: AnalysisOutput):
    """신규 신청자: 현재 상태 (본전공 기준)"""
    st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📈 현재 상태 (본전공 기준)</p>', unsafe_allow_html=True)

    analysis = output.current_analysis
    if analysis is None:
        _missing_box("현재 상태", output.primary_missing_reason)
        return

    col1, col2 = st.columns(2)

    with col1:
        liberal_html = ""
        if student.admission_type == "신입학":
            liberal_html = f"""<tr>
<td style="padding: 8px 0; color: #666;">기초교양(기초문해)</td>
<td style="text-align: right; font-weight: bold;">{analysis.completed_basic_literacy} / {analysis.req_basic_literacy} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">기초교양(기초과학)</td>
<td style="text-align: right; font-weight: bold;">{analysis.completed_basic_science} / {analysis.req_basic_science} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">핵심교양</td>
<td style="text-align: right; font-weight: bold;">{analysis.completed_core_liberal} / {analysis.req_core_liberal} 학점</td>
</tr>"""
        transfer_html = ""
        if student.admission_type != "신입학":
            transfer_html = f"""<tr>
<td style="padding: 8px 0; color: #666;">편입 인정학점</td>
<td style="text-align: right; font-weight: bold;">{student.transfer_credits} 학점</td>
</tr>"""

        st.markdown(f"""
<div style="background: white; border-radius: 12px; padding: 20px; 
box-shadow: 0 2px 10px rgba(0,0,0,0.08);">
<p style="color: #333; margin-bottom: 15px; font-size: 1.1rem; font-weight: 600;">📚 학점 현황</p>
<table style="width: 100%;">
{liberal_html}
{transfer_html}
<tr>
<td style="padding: 8px 0; color: #666;">전공필수</td>
<td style="text-align: right; font-weight: bold;">{analysis.completed_major_required} / {analysis.req_major_required} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">전공선택</td>
<td style="text-align: right; font-weight: bold;">{analysis.completed_major_elective} / {analysis.req_major_elective} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">자유학점</td>
<td style="text-align: right; font-weight: bold;">{student.credits_free} 학점</td>
</tr>
<tr style="border-top: 1px solid #eee;">
<td style="padding: 12px 0; color: #333; font-weight: bold;">총 이수</td>
<td style="text-align: right; font-weight: bold; color: #667eea; font-size: 1.1rem;">
{analysis.completed_total} / {analysis.req_graduation_credits} 학점
</td>
</tr>
</table>
</div>
""", unsafe_allow_html=True)

    with col2:
        status_view = {
            "가능": ("이수 가능", "#28a745"),
            "위험": ("집중 이수 필요", "#f0ad4e"),
        }
        grad_text, grad_color = status_view.get(output.current_status, ("학점 부족", "#dc3545"))

        def _red(v):
            return '#dc3545' if v > 0 else '#28a745'

        liberal_deficit_html = ""
        if student.admission_type == "신입학":
            liberal_deficit_html = f"""<tr>
<td style="padding: 8px 0; color: #666;">기초교양(기초문해) 부족</td>
<td style="text-align: right; font-weight: bold; color: {_red(analysis.deficit_basic_literacy)};">{analysis.deficit_basic_literacy} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">기초교양(기초과학) 부족</td>
<td style="text-align: right; font-weight: bold; color: {_red(analysis.deficit_basic_science)};">{analysis.deficit_basic_science} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">핵심교양 부족</td>
<td style="text-align: right; font-weight: bold; color: {_red(analysis.deficit_core_liberal)};">{analysis.deficit_core_liberal} 학점</td>
</tr>"""

        st.markdown(f"""
<div style="background: white; border-radius: 12px; padding: 20px; 
box-shadow: 0 2px 10px rgba(0,0,0,0.08);">
<p style="color: #333; margin-bottom: 15px; font-size: 1.1rem; font-weight: 600;">⚠️ 부족 현황</p>
<table style="width: 100%;">
{liberal_deficit_html}
<tr>
<td style="padding: 8px 0; color: #666;">전공필수 부족</td>
<td style="text-align: right; font-weight: bold; color: {_red(analysis.deficit_major_required)};">{analysis.deficit_major_required} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">전공선택 부족</td>
<td style="text-align: right; font-weight: bold; color: {_red(analysis.deficit_major_elective)};">{analysis.deficit_major_elective} 학점</td>
</tr>
<tr>
<td style="padding: 8px 0; color: #666;">졸업학점 부족</td>
<td style="text-align: right; font-weight: bold; color: {_red(analysis.deficit_graduation)};">{analysis.deficit_graduation} 학점</td>
</tr>
<tr style="border-top: 1px solid #eee;">
<td style="padding: 12px 0; color: #333; font-weight: bold;">이수 가능 여부</td>
<td style="text-align: right; font-weight: bold; color: {grad_color}; font-size: 1.1rem;">{grad_text}</td>
</tr>
</table>
</div>
""", unsafe_allow_html=True)


def render_step4_results():
    """STEP 4: 결과 확인"""

    student = StudentInput(
        student_type=st.session_state.sim_student_type,
        admission_year=st.session_state.sim_admission_year,
        primary_major=st.session_state.sim_primary_major,
        admission_type=st.session_state.sim_admission_type,
        completed_semesters=st.session_state.sim_completed_semesters,
        transfer_credits=st.session_state.get('sim_transfer_credits', 0),
        credits_basic_literacy=st.session_state.get('sim_credits_basic_literacy', 0),
        credits_basic_science=st.session_state.get('sim_credits_basic_science', 0),
        credits_core_liberal=st.session_state.get('sim_credits_core_liberal', 0),
        credits_major_required=st.session_state.get('sim_credits_major_required', 0),
        credits_major_elective=st.session_state.get('sim_credits_major_elective', 0),
        credits_free=st.session_state.get('sim_credits_free', 0),
        desired_multi_major=st.session_state.get('sim_desired_multi_major'),
        current_program=st.session_state.get('sim_current_program'),
        current_multi_major=st.session_state.get('sim_current_multi_major'),
        credits_multi_required=st.session_state.get('sim_credits_multi_required', 0),
        credits_multi_elective=st.session_state.get('sim_credits_multi_elective', 0),
    )

    output = run_simulation(student)

    st.markdown("""
    <div style="background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); 
                border-radius: 15px; padding: 25px; margin-bottom: 20px; color: white;">
        <p style="margin: 0; color: white; font-size: 1.5rem; font-weight: 600;">📊 분석 결과</p>
        <p style="margin: 10px 0 0 0; opacity: 0.9;">
            입력하신 정보를 바탕으로 분석한 결과입니다
        </p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("🎓 본전공", student.primary_major)
    with col2:
        st.metric("📅 입학연도", f"{student.admission_year}학번")
    with col3:
        remaining = calculate_remaining_semesters(
            student.admission_type, student.completed_semesters, student.primary_major)
        st.metric("⏳ 남은 학기", f"{remaining}학기")

    _render_major_notices(student)

    st.markdown("---")

    if student.student_type == "신규 신청자":
        _render_current_status(student, output)

    if student.student_type == "신규 신청자" and output.simulation_results:
        st.markdown("---")
        st.markdown(f'<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">🎯 다전공 제도별 비교 ({html.escape(str(student.desired_multi_major))})</p>', unsafe_allow_html=True)

        if output.recommended_programs and all(
                r.graduation_status == "어려움" for r in output.recommended_programs):
            st.warning("입력한 학점 기준으로는 모든 제도가 남은 학기 안에 이수하기 어렵습니다. "
                       "순서는 부족 학점이 적은 순이며, 소속 전공 사무실과 상담해 보세요.")
        for idx, result in enumerate(output.recommended_programs):
            render_simulation_result_card(result, idx == 0)

        for result in output.supplementary_programs:
            _missing_box(result.program_type, result.missing_reason)

    elif student.student_type == "기존 참여자" and output.simulation_results:
        st.markdown(f'<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📚 현재 참여 중인 다전공 분석 ({student.current_program})</p>', unsafe_allow_html=True)
        result = output.simulation_results[0]
        if result.data_missing:
            _missing_box(f"{result.program_type} · {result.multi_major_name}", result.missing_reason)
        else:
            render_current_participant_analysis(result, student)

    st.caption(RESULT_DISCLAIMER)

    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col1:
        if st.button("⬅️ 이전", use_container_width=True):
            st.session_state.sim_step = 3
            st.rerun()
    with col3:
        if st.button("🔄 처음부터", use_container_width=True):
            keys_to_delete = [k for k in st.session_state.keys() if k.startswith('sim_')]
            for k in keys_to_delete:
                del st.session_state[k]
            st.session_state.sim_step = 1
            st.rerun()


def render_simulation_result_card(result: SimulationResult, is_top: bool):
    """비교 분석 결과 카드 렌더링"""

    analysis = result.credit_analysis
    # '어려움'이면 1위여도 추천 배지 · 강조 테두리를 붙이지 않음
    is_top = is_top and result.graduation_status in ("가능", "위험")
    border_style = "3px solid #667eea" if is_top else "1px solid #e9ecef"

    major_deficit = (
        analysis.deficit_major_required +
        analysis.deficit_major_elective +
        analysis.deficit_multi_required +
        analysis.deficit_multi_elective
    )

    def _cell(title, done, req, deficit, bg):
        mark = f"부족 {deficit}" if deficit > 0 else "✓"
        color = "#dc3545" if deficit > 0 else "#28a745"
        return f"""<div style="text-align: center; padding: 10px; background: {bg}; border-radius: 8px;">
<div style="font-size: 0.85rem; color: #666;">{title}</div>
<div style="font-size: 0.95rem; font-weight: bold; color: #333;">{done}/{req}</div>
<div style="font-size: 0.8rem; color: {color}; margin-top: 3px;">{mark}</div>
</div>"""

    cells = "".join([
        _cell("본전공 필수", analysis.completed_major_required, analysis.req_major_required, analysis.deficit_major_required, "#f8f9fa"),
        _cell("본전공 선택", analysis.completed_major_elective, analysis.req_major_elective, analysis.deficit_major_elective, "#f8f9fa"),
        _cell("다전공 필수", analysis.completed_multi_required, analysis.req_multi_required, analysis.deficit_multi_required, "#e3f2fd"),
        _cell("다전공 선택", analysis.completed_multi_elective, analysis.req_multi_elective, analysis.deficit_multi_elective, "#e3f2fd"),
    ])

    top_badge = ('<span style="background: #667eea; color: white; padding: 3px 10px; border-radius: 15px; '
                 'font-size: 0.8rem; margin-left: 10px;">👑 추천 1위</span>') if is_top else ""

    st.markdown(f"""
<div style="background: white; border-radius: 12px; padding: 20px; 
margin-bottom: 15px; border-left: {border_style};
box-shadow: 0 2px 10px rgba(0,0,0,0.05);">
<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px;">
<div>
<span style="font-size: 1.3rem; font-weight: bold; color: #333;">{result.program_type}</span>
{top_badge}
</div>
</div>
<div style="display: grid; grid-template-columns: repeat(5, 1fr); gap: 15px; margin-bottom: 15px;">
{cells}
<div style="text-align: center; padding: 10px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8px;">
<div style="font-size: 0.85rem; color: white; font-weight: bold;">이수해야 하는<br>총 전공 학점수</div>
<div style="font-size: 1.2rem; font-weight: bold; color: white; margin-top: 5px;">{major_deficit}</div>
</div>
</div>
<div style="color: #555; font-size: 0.9rem; margin-bottom: 10px;">
🎓 {_graduation_line(analysis, result.program_type)} · 현재 총 이수 {analysis.completed_total}학점 · 졸업학점 부족 {analysis.deficit_graduation}학점
</div>
<div style="background: #f8f9fa; border-radius: 8px; padding: 12px;">
<span style="color: #666;">💡 </span>
<span style="color: #333;">{result.recommendation_reason}</span>
</div>
</div>
""", unsafe_allow_html=True)

    if result.semester_plan:
        with st.expander(f"📅 {result.program_type} 학기별 이수 계획"):
            render_semester_plan_table(result.semester_plan)


def render_current_participant_analysis(result: SimulationResult, student: StudentInput):
    """기존 참여자 분석 결과 렌더링"""

    analysis = result.credit_analysis

    def _row(title, done, req, deficit, first=False):
        w = ' width: 50%;' if first else ''
        w2 = ' width: 30%;' if first else ''
        w3 = ' width: 20%;' if first else ''
        color = '#dc3545' if deficit > 0 else '#28a745'
        mark = f"부족 {deficit}학점" if deficit > 0 else "✓"
        return f"""<tr>
<td style="padding: 8px 0; color: #666;{w}">{title}</td>
<td style="text-align: right;{w2}">{done} / {req} 학점</td>
<td style="text-align: right;{w3} color: {color}; font-weight: bold;">{mark}</td>
</tr>"""

    primary_rows = ""
    if student.admission_type == "신입학":
        primary_rows += _row("기초교양(기초문해)", student.credits_basic_literacy, analysis.req_basic_literacy, analysis.deficit_basic_literacy, first=True)
        primary_rows += _row("기초교양(기초과학)", student.credits_basic_science, analysis.req_basic_science, analysis.deficit_basic_science)
        primary_rows += _row("핵심교양", student.credits_core_liberal, analysis.req_core_liberal, analysis.deficit_core_liberal)
    else:
        primary_rows += f"""<tr>
<td style="padding: 8px 0; color: #666; width: 50%;">편입 인정학점</td>
<td style="text-align: right; width: 30%;">{student.transfer_credits} 학점</td>
<td style="width: 20%;"></td>
</tr>"""
    primary_rows += _row("전공필수", student.credits_major_required, analysis.req_major_required, analysis.deficit_major_required)
    primary_rows += _row("전공선택", student.credits_major_elective, analysis.req_major_elective, analysis.deficit_major_elective)
    primary_rows += f"""<tr>
<td style="padding: 8px 0; color: #666;">자유학점</td>
<td style="text-align: right;">{student.credits_free} 학점</td>
<td></td>
</tr>"""

    col1, col2 = st.columns(2)

    with col1:
        st.markdown(f"""
<div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.08);">
<p style="color: #667eea; margin-bottom: 15px; font-size: 1.1rem; font-weight: 600;">🎓 본전공 ({html.escape(str(student.primary_major))})</p>
<table style="width: 100%;">
{primary_rows}
</table>
</div>
""", unsafe_allow_html=True)

    with col2:
        multi_rows = (
            _row("전공필수", student.credits_multi_required, analysis.req_multi_required, analysis.deficit_multi_required, first=True) +
            _row("전공선택", student.credits_multi_elective, analysis.req_multi_elective, analysis.deficit_multi_elective)
        )
        st.markdown(f"""
<div style="background: white; border-radius: 12px; padding: 20px; box-shadow: 0 2px 10px rgba(0,0,0,0.08);">
<p style="color: #764ba2; margin-bottom: 15px; font-size: 1.1rem; font-weight: 600;">📘 다전공 ({html.escape(str(student.current_multi_major))})</p>
<table style="width: 100%;">
{multi_rows}
</table>
</div>
""", unsafe_allow_html=True)

    # 총 이수학점 대비 부족학점 (편입 인정학점 포함)
    col_total, _ = st.columns(2)
    with col_total:
        deficit = analysis.deficit_graduation
        st.markdown(f"""
<div style="margin-top: 20px; padding: 8px 0;">
<table style="width: 100%;">
<tr>
<td style="padding: 8px 0; color: #333; font-weight: bold; width: 50%;">📊 총 이수학점 대비 부족학점</td>
<td style="text-align: right; font-weight: bold; color: #333; width: 30%;">{analysis.completed_total} / {analysis.req_graduation_credits} 학점</td>
<td style="text-align: right; font-weight: bold; color: {'#dc3545' if deficit > 0 else '#28a745'}; width: 20%;">{'부족 ' + str(deficit) + '학점' if deficit > 0 else '✓'}</td>
</tr>
</table>
<p style="color: #666; font-size: 0.85rem; margin: 4px 0 0 0;">🎓 {_graduation_line(analysis, result.program_type)}</p>
</div>
""", unsafe_allow_html=True)

    status_view = {
        "가능": ("이수 가능", "#28a745"),
        "위험": ("집중 이수 필요", "#f0ad4e"),
    }
    status_text, status_color = status_view.get(result.graduation_status, ("학점 부족", "#dc3545"))

    per_semester = math.ceil(analysis.needed_credits / max(1, analysis.remaining_semesters))

    st.markdown(f"""
<div style="background: linear-gradient(135deg, {status_color}15 0%, {status_color}05 100%); 
border-left: 4px solid {status_color}; border-radius: 12px; padding: 20px; margin-top: 20px;">
<p style="color: {status_color}; margin: 0 0 10px 0; font-size: 1.1rem; font-weight: 600;">
{'✅' if result.graduation_status == '가능' else '⚠️'} {status_text}
</p>
<p style="color: #666; margin: 0;">
남은 학기: <strong>{analysis.remaining_semesters}학기</strong> / 
졸업까지 추가로 필요한 학점: <strong>{analysis.needed_credits}학점</strong> /
학기당 평균: <strong>{per_semester}학점</strong>
</p>
</div>
""", unsafe_allow_html=True)

    if result.semester_plan:
        st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📅 학기별 이수 계획</p>', unsafe_allow_html=True)
        render_semester_plan_table(result.semester_plan)


def render_semester_plan_table(plan: List[Dict]):
    """학기별 이수 계획 테이블"""

    if not plan:
        st.info("이수할 학점이 없습니다.")
        return

    df = pd.DataFrame(plan)
    has_basic_science = any(row.get('basic_science', 0) > 0 for row in plan)

    if has_basic_science:
        df.columns = ['학년/학기', '기초교양(기초문해)', '기초교양(기초과학)', '핵심교양',
                      '본전공 필수', '본전공 선택', '다전공 필수', '다전공 선택', '자유학점', '합계']
    else:
        df = df.drop(columns=['basic_science'])
        df.columns = ['학년/학기', '기초교양(기초문해)', '핵심교양',
                      '본전공 필수', '본전공 선택', '다전공 필수', '다전공 선택', '자유학점', '합계']

    column_config = {"학년/학기": st.column_config.TextColumn("학년/학기", width="medium")}
    for col in df.columns[1:]:
        column_config[col] = st.column_config.NumberColumn(col, format="%d학점")

    st.dataframe(df, hide_index=True, use_container_width=True, column_config=column_config)


# ============================================================
# 메인 함수 (테스트용)
# ============================================================

if __name__ == "__main__":
    st.set_page_config(page_title="다전공 비교 분석", page_icon="🎯", layout="wide")
    render_simulation_page()