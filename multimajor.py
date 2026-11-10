"""
============================================================
🎓 한경국립대 다전공 안내 사이트 (multimajor.py)
============================================================
기능
1. 다전공 제도 안내 (제도 한눈에 비교, 상세 정보, 전공별 이수체계도·교과목·연락처, 학생설계전공 샘플)
2. 다전공 비교 분석 (simulation.py)

※ AI 챗봇 기능은 별도 학사안내 AI 챗봇으로 이전하여 이 파일에서 제거함
============================================================
"""

import streamlit as st
import pandas as pd
from streamlit_option_menu import option_menu
from datetime import datetime, timezone, timedelta
import os
import yaml

# ============================================================
# 📌 설정 파일 로드
# ============================================================

def load_yaml_config(filename):
    """YAML 설정 파일 로드"""
    config_path = os.path.join('config', filename)
    if os.path.exists(config_path):
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    return {}

MESSAGES = load_yaml_config('messages.yaml')
MAPPINGS = load_yaml_config('mappings.yaml')
SETTINGS = load_yaml_config('settings.yaml')

# ============================================================
# 📌 상수 정의
# ============================================================

DEFAULT_CONTACT_MESSAGE = "📞 문의: 전공 사무실 또는 교육학사혁신팀 031-670-5035로 연락주시면 보다 상세한 정보를 안내 받을 수 있습니다."
CONTACT_MESSAGE = MESSAGES.get('contact', {}).get('default', DEFAULT_CONTACT_MESSAGE)

LINKS = MESSAGES.get('links', {})
ACADEMIC_NOTICE_URL = LINKS.get('academic_notice', "https://www.hknu.ac.kr/kor/562/subview.do")

PATHS = SETTINGS.get('paths', {})
CURRICULUM_IMAGES_PATH = PATHS.get('curriculum_images', "images/curriculum")

DIFFICULTY_STARS = MAPPINGS.get('difficulty_stars', {})


def convert_difficulty_to_stars(value):
    if pd.isna(value) or value == '':
        return DIFFICULTY_STARS.get('default', '⭐⭐⭐')
    if isinstance(value, str) and '⭐' in value:
        return value
    try:
        num = int(float(value))
        return DIFFICULTY_STARS.get(num, DIFFICULTY_STARS.get('default', '⭐⭐⭐'))
    except:
        return DIFFICULTY_STARS.get('default', '⭐⭐⭐')


# 페이지 설정
st.set_page_config(
    page_title="다전공 안내",
    page_icon="🎓",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items={
        'Get help': 'https://www.hknu.ac.kr',
        'Report a bug': 'https://www.hknu.ac.kr',
        'About': "# 한경국립대 다전공 안내"
    }
)

# CSS 스타일
hide_streamlit_style = """
<style>
    [data-testid="stDecoration"] { display: none !important; }
    footer { display: none !important; }
    .main .block-container { padding-top: 2rem !important; }
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)


def initialize_session_state():
    """세션 상태 초기화"""
    if 'page' not in st.session_state:
        st.session_state.page = "다전공 제도 안내"


# ============================================================
# 📂 데이터 로드
# ============================================================

@st.cache_data
def load_excel_data(file_path, sheet_name=0):
    try:
        if os.path.exists(file_path):
            result = pd.read_excel(file_path, sheet_name=sheet_name)
            if isinstance(result, dict):
                return list(result.values())[0] if result else pd.DataFrame()
            return result
        return pd.DataFrame()
    except Exception as e:
        return pd.DataFrame()


@st.cache_data
def load_program_info():
    df = load_excel_data('data/programs.xlsx')
    if not isinstance(df, pd.DataFrame) or df.empty:
        return {}
    programs = {}
    for _, row in df.iterrows():
        name = row.get('제도명', '')
        if name and pd.notna(name):
            def safe_get(key, default=''):
                val = row.get(key, default)
                return default if pd.isna(val) else val
            
            programs[name] = {
                'description': safe_get('설명', ''),
                'qualification': safe_get('신청자격', ''),
                'credits_general': safe_get('이수학점(교양)', ''),
                'credits_primary': safe_get('원전공 이수학점', ''),
                'credits_multi': safe_get('다전공 이수학점', ''),
                'degree': safe_get('학위기 표기', '-'),
                'features': str(safe_get('특징', '')).split('\n') if safe_get('특징', '') else [],
                'notes': safe_get('기타', ''),
                'difficulty': convert_difficulty_to_stars(safe_get('난이도', '3')),
                'graduation_certification': safe_get('졸업인증', '-'),
                'graduation_exam': safe_get('졸업시험', '-'),
                'summary': str(safe_get('카드요약', '')).strip(),
                'multi_display': str(safe_get('다전공학점표시', '')).strip(),
                'primary_display': str(safe_get('원전공학점표시', '')).strip(),
                'group': str(safe_get('묶음', '')).strip(),
                'difficulty_basis': str(safe_get('난이도기준', '')).strip(),

            }
    return programs


@st.cache_data
def load_curriculum_mapping():
    try:
        if os.path.exists('data/curriculum_mapping.xlsx'):
            return pd.read_excel('data/curriculum_mapping.xlsx')
        return pd.DataFrame(columns=['전공명', '제도유형', '파일명'])
    except:
        return pd.DataFrame(columns=['전공명', '제도유형', '파일명'])


@st.cache_data
def load_courses_data():
    try:
        if os.path.exists('data/courses.xlsx'):
            return pd.read_excel('data/courses.xlsx')
        return pd.DataFrame(columns=['전공명', '제도유형', '학년', '학기', '이수구분', '과목명', '학점'])
    except:
        return pd.DataFrame(columns=['전공명', '제도유형', '학년', '학기', '이수구분', '과목명', '학점'])


@st.cache_data
def load_majors_info():
    """전공 정보 로드 (복수전공, 부전공 등 일반 전공)"""
    return load_excel_data('data/majors_info.xlsx')


@st.cache_data
def load_microdegree_info():
    """마이크로디그리(소단위전공과정) 정보 로드 - 과정 중심"""
    return load_excel_data('data/microdegree_info.xlsx')


@st.cache_data
def load_graduation_requirements():
    return load_excel_data('data/graduation_requirements.xlsx')


@st.cache_data
def load_primary_requirements():
    return load_excel_data('data/primary_requirements.xlsx')


# 데이터 로드
PROGRAM_INFO = load_program_info()
CURRICULUM_MAPPING = load_curriculum_mapping()
COURSES_DATA = load_courses_data()
MAJORS_INFO = load_majors_info()
MICRODEGREE_INFO = load_microdegree_info()
GRADUATION_REQ = load_graduation_requirements()
PRIMARY_REQ = load_primary_requirements()

ALL_DATA = {
    'programs': PROGRAM_INFO,
    'curriculum': CURRICULUM_MAPPING,
    'courses': COURSES_DATA,
    'majors': MAJORS_INFO,
    'microdegree': MICRODEGREE_INFO,
    'grad_req': GRADUATION_REQ,
    'primary_req': PRIMARY_REQ,
}


# ============================================================
# 🏫 계열별 전공 그룹화 헬퍼 함수
# ============================================================

def get_majors_by_category(program_type=None, data_source="majors"):
    """계열별로 전공을 그룹화하여 반환"""
    special_programs = ["융합전공", "융합부전공", "소단위전공과정", "마이크로디그리"]
    
    # 🔥 마이크로디그리/소단위전공과정인 경우 microdegree_info.xlsx 사용
    if program_type in ["소단위전공과정", "마이크로디그리"]:
        if not MICRODEGREE_INFO.empty and '과정명' in MICRODEGREE_INFO.columns:
            field_column = '계열' if '계열' in MICRODEGREE_INFO.columns else None
            
            if field_column:
                field_courses = {}
                for _, row in MICRODEGREE_INFO.iterrows():
                    field = row.get(field_column, '기타')
                    if pd.isna(field) or str(field).strip() == '':
                        field = '기타'
                    field = str(field).strip()
                    
                    course_name = row.get('과정명', '')
                    if course_name:
                        if field not in field_courses:
                            field_courses[field] = []
                        if course_name not in field_courses[field]:
                            field_courses[field].append(course_name)
                
                for field in field_courses:
                    field_courses[field] = sorted(field_courses[field])
                
                return field_courses
            else:
                return {"전체": sorted(MICRODEGREE_INFO['과정명'].unique().tolist())}
        return {}
    
    # 융합전공, 융합부전공
    if program_type in ["융합전공", "융합부전공"]:
        majors_list = []
        
        if not MAJORS_INFO.empty and '제도유형' in MAJORS_INFO.columns:
            if program_type == "융합전공":
                mask = MAJORS_INFO['제도유형'].str.contains('융합전공', na=False) & ~MAJORS_INFO['제도유형'].str.contains('융합부전공', na=False)
            else:
                mask = MAJORS_INFO['제도유형'].str.contains(program_type, na=False)
            majors_list = MAJORS_INFO[mask]['전공명'].unique().tolist()
        
        if data_source == "courses" and not COURSES_DATA.empty and '제도유형' in COURSES_DATA.columns:
            if program_type == "융합전공":
                mask = COURSES_DATA['제도유형'].str.contains('융합전공', na=False) & ~COURSES_DATA['제도유형'].str.contains('융합부전공', na=False)
            else:
                mask = COURSES_DATA['제도유형'].str.contains(program_type, na=False)
            for m in COURSES_DATA[mask]['전공명'].unique():
                if m not in majors_list:
                    majors_list.append(m)
        
        return {"전체": sorted(majors_list)} if majors_list else {}
    
    # 일반 전공 (복수전공, 부전공, 연계전공)
    category_majors = {}
    
    if not MAJORS_INFO.empty:
        has_category = '계열' in MAJORS_INFO.columns
        
        if program_type:
            if program_type == "부전공":
                mask = MAJORS_INFO['제도유형'].str.contains('부전공', na=False) & ~MAJORS_INFO['제도유형'].str.contains('융합부전공', na=False)
            else:
                mask = MAJORS_INFO['제도유형'].str.contains(program_type, na=False)
            filtered_df = MAJORS_INFO[mask]
        else:
            filtered_df = MAJORS_INFO
        
        if has_category:
            for _, row in filtered_df.iterrows():
                category = row.get('계열', '기타')
                if pd.isna(category) or str(category).strip() == '':
                    category = '기타'
                category = str(category).strip()
                major_name = row['전공명']
                
                if category not in category_majors:
                    category_majors[category] = []
                if major_name not in category_majors[category]:
                    category_majors[category].append(major_name)
        else:
            category_majors["전체"] = filtered_df['전공명'].unique().tolist()
    
    for cat in category_majors:
        category_majors[cat] = sorted(category_majors[cat])
    
    return category_majors


# ============================================================
# 📊 이수체계도 및 과목 표시 함수
# ============================================================

def display_curriculum_image(major, program_type):
    """이수체계도/과정 안내 이미지 표시"""
    if not major or major == "선택 안 함":
        return
    
    is_fusion = program_type == "융합전공"
    is_micro = "소단위" in program_type or "마이크로" in program_type
    
    if not is_fusion and not is_micro:
        return
    
    if CURRICULUM_MAPPING.empty:
        return
    
    def match_program_type_for_image(type_value):
        type_str = str(type_value).strip().lower()
        if is_fusion:
            return "융합전공" in type_str and "융합부전공" not in type_str
        if is_micro:
            return any(kw in type_str for kw in ['소단위', '마이크로', 'md'])
        return False
    
    clean_major = major
    if major.endswith(')') and '(' in major:
        last_open_paren = major.rfind('(')
        if last_open_paren > 0:
            clean_major = major[:last_open_paren].strip()
    
    search_keyword = clean_major.replace('전공', '').replace('과정', '').replace('전문가', '').replace('MD', '').replace('(', '').replace(')', '').replace(' ', '').strip()
    
    type_matched = CURRICULUM_MAPPING[CURRICULUM_MAPPING['제도유형'].apply(match_program_type_for_image)]
    
    if type_matched.empty:
        return
    
    filtered = type_matched[type_matched['전공명'] == clean_major]
    
    if filtered.empty:
        filtered = type_matched[type_matched['전공명'] == major]
    
    if filtered.empty:
        clean_major_no_space = clean_major.replace(' ', '')
        for _, row in type_matched.iterrows():
            cm_major = str(row['전공명'])
            cm_major_no_space = cm_major.replace(' ', '')
            if clean_major_no_space == cm_major_no_space:
                filtered = type_matched[type_matched['전공명'] == cm_major]
                break
    
    if filtered.empty and len(search_keyword) >= 2:
        for _, row in type_matched.iterrows():
            cm_major = str(row['전공명'])
            cm_keyword = cm_major.replace('전공', '').replace('과정', '').replace('전문가', '').replace('MD', '').replace('(', '').replace(')', '').replace(' ', '').strip()
            if len(cm_keyword) >= 2 and len(search_keyword) >= 2:
                if search_keyword in cm_keyword or cm_keyword in search_keyword:
                    filtered = type_matched[type_matched['전공명'] == cm_major]
                    break
    
    if not filtered.empty:
        images_shown = 0
        missing_files = []
        total_images = len(filtered)
        
        for idx, row in filtered.iterrows():
            filename = row['파일명']
            
            if pd.notna(filename) and str(filename).strip():
                filename_str = str(filename).strip()
                
                if ',' in filename_str:
                    file_list = [f.strip() for f in filename_str.split(',')]
                    for file in file_list:
                        image_path = f"{CURRICULUM_IMAGES_PATH}/{file}"
                        if os.path.exists(image_path):
                            if is_fusion:
                                caption = f"{clean_major} 이수체계도"
                            else:
                                caption = f"{clean_major} 과정 안내 ({images_shown + 1})"
                            st.image(image_path, caption=caption)
                            images_shown += 1
                        else:
                            missing_files.append(file)
                else:
                    image_path = f"{CURRICULUM_IMAGES_PATH}/{filename_str}"
                    
                    if os.path.exists(image_path):
                        if is_fusion:
                            caption = f"{clean_major} 이수체계도"
                        else:
                            if total_images > 1:
                                caption = f"{clean_major} 과정 안내 ({images_shown + 1}/{total_images})"
                            else:
                                caption = f"{clean_major} 과정 안내"
                        st.image(image_path, caption=caption)
                        images_shown += 1
                    else:
                        missing_files.append(filename_str)
        
        if missing_files:
            st.warning(f"⚠️ 다음 이미지 파일을 찾을 수 없습니다:")
            for missing_file in missing_files:
                st.caption(f"   • `{CURRICULUM_IMAGES_PATH}/{missing_file}`")
        
        if images_shown == 0 and not missing_files:
            st.caption("📷 이미지 파일 준비 중입니다.")
    else:
        st.info(f"💡 '{major}' 또는 '{clean_major}'에 해당하는 이미지 정보를 찾을 수 없습니다.")


def render_course_list(df, is_micro):
    for idx, row in df.iterrows():
        course_name = row.get('과목명', '')
        credit = f"{int(row.get('학점', 0))}학점" if pd.notna(row.get('학점')) else ""
        desc = row.get('교과목개요')

        title = f"📘 {course_name} ({credit})"

        with st.expander(title):
            if desc and pd.notna(desc) and str(desc).strip():
                st.write(desc)
            else:
                st.info("교과목 개요 정보가 없습니다.")

            edu_dept = row.get('교육 운영전공') or row.get('교육운영전공', '')
            if is_micro and pd.notna(edu_dept) and str(edu_dept).strip():
                st.caption(f"🏫 운영전공: {str(edu_dept).strip()}")


def display_courses(major, program_type):
    """과목 정보 표시 - 수정 버전"""
    if not major or major == "선택 안 함":
        return False
    
    if COURSES_DATA.empty:
        st.info("교과목 데이터가 없습니다.")
        return False
    
    is_micro = "소단위" in program_type or "마이크로" in program_type
    
    def match_program_type_for_courses(type_value):
        """제도유형 매칭 - 개선 버전"""
        type_str = str(type_value).strip()
        type_list = [t.strip() for t in type_str.split(',')]
        
        if is_micro:
            return any(kw in type_str.lower() for kw in ['소단위', '마이크로', 'md'])
        
        if program_type == "부전공":
            return "부전공" in type_list and "융합부전공" not in type_list
        
        if program_type == "융합전공":
            return "융합전공" in type_list
        
        if program_type == "융합부전공":
            return "융합부전공" in type_list
        
        if program_type == "연계전공":
            return "연계전공" in type_list
        
        return program_type in type_list
    
    clean_major = major
    display_major = major
    
    if major.endswith(')') and '(' in major:
        last_open_paren = major.rfind('(')
        if last_open_paren > 0:
            clean_major = major[:last_open_paren].strip()
            display_major = clean_major
    
    courses = COURSES_DATA[
        (COURSES_DATA['전공명'] == clean_major) & 
        (COURSES_DATA['제도유형'].apply(match_program_type_for_courses))
    ]
    
    if courses.empty and is_micro:
        keyword = clean_major.replace('전공', '').replace('과정', '').replace('전문가', '').replace('MD', '').replace(' ', '').strip()
        type_matched = COURSES_DATA[COURSES_DATA['제도유형'].apply(match_program_type_for_courses)]
        
        for course_major in type_matched['전공명'].unique():
            cm_str = str(course_major)
            if 'MD' in cm_str or 'md' in cm_str.lower():
                cm_keyword = cm_str.replace('MD', '').replace('md', '').replace(' ', '').strip()
                if len(keyword) >= 2 and len(cm_keyword) >= 2:
                    if keyword[:2] in cm_keyword or cm_keyword[:2] in keyword:
                        courses = type_matched[type_matched['전공명'] == course_major]
                        display_major = cm_str
                        break

    if courses.empty:
        keyword = clean_major.replace('전공', '').replace('과정', '').replace('(', '').replace(')', '')[:4]
        if keyword:
            courses = COURSES_DATA[
                (COURSES_DATA['전공명'].str.contains(keyword, na=False, regex=False)) & 
                (COURSES_DATA['제도유형'].apply(match_program_type_for_courses))
            ]
            if not courses.empty:
                display_major = courses['전공명'].iloc[0]
    
    display_program_type = "소단위전공과정(마이크로디그리)" if is_micro else program_type
    
    if not courses.empty:
        st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📚 교과목 안내</p>', unsafe_allow_html=True)
        
        years = sorted([int(y) for y in courses['학년'].unique() if pd.notna(y)])
        
        if years:
            tabs = st.tabs([f"{year}학년" for year in years])
            
            for idx, year in enumerate(years):
                with tabs[idx]:
                    year_courses = courses[courses['학년'] == year]
                    semesters = sorted([int(s) for s in year_courses['학기'].unique() if pd.notna(s)])
                    
                    for semester in semesters:
                        st.markdown(f'<p style="font-size: 1.1rem; font-weight: 600; margin: 16px 0 12px 0;">📅 {semester}학기</p>', unsafe_allow_html=True)
                        semester_courses = year_courses[year_courses['학기'] == semester]
                        
                        required = semester_courses[semester_courses['이수구분'].str.contains('필수', na=False)]
                        elective = semester_courses[semester_courses['이수구분'].str.contains('선택', na=False)]
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            if not required.empty:
                                st.markdown("**🔴 전공필수**")
                                render_course_list(required, is_micro)
                            else:
                                st.caption("이 학기에는 전공필수 과목이 없습니다.")
                                                                    
                        with col2:
                            if not elective.empty:
                                st.markdown("**🟢 전공선택**")
                                render_course_list(elective, is_micro)
                            else:
                                st.caption("이 학기에는 전공선택 과목이 없습니다.")                        
                        st.divider()
        else:
            semesters = sorted([int(s) for s in courses['학기'].unique() if pd.notna(s)])
            
            if semesters:
                for semester in semesters:
                    st.markdown(f'<p style="font-size: 1.1rem; font-weight: 600; margin: 16px 0 12px 0;">📅 {semester}학기</p>', unsafe_allow_html=True)
                    semester_courses = courses[courses['학기'] == semester]
                    
                    has_required = not semester_courses[semester_courses['이수구분'].str.contains('필수', na=False)].empty
                    has_elective = not semester_courses[semester_courses['이수구분'].str.contains('선택', na=False)].empty
                    
                    if has_required or has_elective:
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            required = semester_courses[semester_courses['이수구분'].str.contains('필수', na=False)]
                            if not required.empty:
                                st.markdown("**🔴 전공필수**")
                                render_course_list(required, is_micro)
                        
                        with col2:
                            elective = semester_courses[semester_courses['이수구분'].str.contains('선택', na=False)]
                            if not elective.empty:
                                st.markdown("**🟢 전공선택**")
                                render_course_list(elective, is_micro)
                    
                    # 🔥 수정: else 블록에서도 render_course_list 사용!
                    else:
                        st.markdown("**📚 교과목 목록**")
                        render_course_list(semester_courses, is_micro)
                    
                    st.divider()
            else:
                # 🔥 수정: 여기서도 render_course_list 사용!
                st.markdown("**📚 교과목 목록**")
                render_course_list(courses, is_micro)
        
        st.markdown("---")
        display_major_contact(display_major, program_type)
        return True
    else:
        st.info(f"'{display_major}' 교과목 정보가 없습니다.")
        return False

def display_major_contact(major, program_type="전공"):
    """전공 연락처 표시 - 마이크로디그리 지원"""
    
    # 🔥 마이크로디그리 체크
    is_micro = "소단위" in program_type or "마이크로" in program_type
    
    # 마이크로디그리인 경우 MICRODEGREE_INFO에서 찾기
    if is_micro and not MICRODEGREE_INFO.empty:
        clean_major = major
        
        # 괄호 제거
        if major.endswith(')') and '(' in major:
            last_open_paren = major.rfind('(')
            if last_open_paren > 0:
                clean_major = major[:last_open_paren].strip()
        
        # MD 제거
        clean_major = clean_major.replace(' MD', '').replace('MD', '').strip()
        
        # 🔥 MICRODEGREE_INFO에서 검색
        contact_row = pd.DataFrame()
        
        # 1차: 정확한 과정명 매칭
        if '과정명' in MICRODEGREE_INFO.columns:
            contact_row = MICRODEGREE_INFO[MICRODEGREE_INFO['과정명'] == major]
        
        # 2차: 괄호 제거 후 매칭
        if contact_row.empty:
            contact_row = MICRODEGREE_INFO[MICRODEGREE_INFO['과정명'] == clean_major]
        
        # 3차: 부분 매칭
        if contact_row.empty:
            keyword = clean_major.replace('전공', '').replace('과정', '').replace('전문가', '')
            if keyword:
                contact_row = MICRODEGREE_INFO[
                    MICRODEGREE_INFO['과정명'].str.contains(keyword, na=False, regex=False)
                ]
        
        # 마이크로디그리 정보 표시
        if not contact_row.empty:
            row = contact_row.iloc[0]
            
            course_name = row.get('과정명', major)
            edu_major = row.get('교육운영전공', '')
            phone = row.get('연락처', '')
            location = row.get('위치', row.get('사무실위치', ''))
            
            contact_parts = [f"🎓 **과정명**: {course_name}"]
            
            if pd.notna(edu_major) and str(edu_major).strip():
                contact_parts.append(f"🏛️ **교육운영전공**: {edu_major}")
            
            if pd.notna(phone) and str(phone).strip():
                contact_parts.append(f"📞 **연락처**: {phone}")
            
            if pd.notna(location) and str(location).strip():
                contact_parts.append(f"📍 **사무실 위치**: {location}")
            
            st.info(f"**📋 소단위전공과정 문의처**\n\n" + "\n\n".join(contact_parts))
            return
    
    # 일반 전공인 경우 MAJORS_INFO에서 찾기
    if MAJORS_INFO.empty:
        st.info(f"📞 **문의**: 교육학사혁신팀 031-670-5035")
        return
    
    edu_major = None
    clean_major = major
    if major.endswith(')') and '(' in major:
        last_open_paren = major.rfind('(')
        if last_open_paren > 0:
            edu_major = major[last_open_paren+1:-1].strip()
            clean_major = major[:last_open_paren].strip()
    
    clean_major = clean_major.replace(' MD', '').replace('MD', '').strip()
    
    contact_row = pd.DataFrame()
    if edu_major:
        contact_row = MAJORS_INFO[MAJORS_INFO['전공명'] == edu_major]

    if contact_row.empty:
        contact_row = MAJORS_INFO[MAJORS_INFO['전공명'] == clean_major]
    
    if contact_row.empty:
        keyword = clean_major.replace('전공', '').replace('과정', '').replace('(', '').replace(')', '')[:4]
        if keyword:
            contact_row = MAJORS_INFO[MAJORS_INFO['전공명'].str.contains(keyword, na=False, regex=False)]
    
    if not contact_row.empty:
        row = contact_row.iloc[0]
        major_name = row.get('전공명', major)
        phone = row.get('연락처', '')
        location = row.get('사무실위치', row.get('위치', ''))
        
        contact_title = f"{program_type} 문의처"
        
        contact_parts = [f"🎓 **전공명**: {major_name}"]
        if pd.notna(phone) and str(phone).strip():
            contact_parts.append(f"📞 **연락처**: {phone}")
        if pd.notna(location) and str(location).strip():
            contact_parts.append(f"📍 **사무실 위치**: {location}")
        
        st.info(f"**📋 {contact_title}**\n\n" + "\n\n".join(contact_parts))
    else:
        st.info(f"📞 **문의**: 교육학사혁신팀 031-670-5035")

def render_student_design_samples():
    """학생설계전공: 샘플 교육과정 안내"""
    st.info(
        "학생설계전공은 기존 전공을 선택하는 제도가 아니라, "
        "학생이 교과목을 조합해 나만의 전공과정을 설계하는 제도입니다. "
        "아래는 구성 방법을 이해하기 위한 **샘플**입니다."
    )

    df = load_excel_data('data/student_design_samples.xlsx')
    required = ['샘플명', '구분', '과목명', '학점']
    if df.empty or any(c not in df.columns for c in required):
        st.warning("샘플 교육과정 자료가 아직 등록되지 않았습니다. (data/student_design_samples.xlsx)")
        return

    names = list(dict.fromkeys(df['샘플명'].dropna().astype(str)))
    selected = st.selectbox("샘플 교육과정 선택", names, key="sd_sample")
    sub = df[df['샘플명'].astype(str) == selected].copy()
    sub['학점'] = pd.to_numeric(sub['학점'], errors='coerce').fillna(0).astype(int)

    # 학년/학기 열 합치기
    if '학년' in sub.columns and '학기' in sub.columns:
        sub['학년'] = pd.to_numeric(sub['학년'], errors='coerce')
        sub['학기'] = pd.to_numeric(sub['학기'], errors='coerce')
        sub['학년/학기'] = sub.apply(
            lambda r: f"{int(r['학년'])}학년 {int(r['학기'])}학기"
            if pd.notna(r['학년']) and pd.notna(r['학기']) else "-", axis=1)
        sub = sub.sort_values(['학년', '학기'], kind='stable')
    else:
        sub['학년/학기'] = "-"

    # 샘플 설명 (첫 행 기준, 줄바꿈 → 문단 구분)
    if '설명' in sub.columns:
        desc = sub['설명'].dropna().astype(str)
        if not desc.empty:
            text = desc.iloc[0].replace("\\n", "\n").replace("\n", "\n\n")
            with st.expander("📖 이 샘플의 취지 보기", expanded=False):
                st.markdown(text)

    # 요약 지표
    req_credit = int(sub.loc[sub['구분'] == '전공필수', '학점'].sum())
    ele_credit = int(sub.loc[sub['구분'] == '전공선택', '학점'].sum())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("샘플 총 학점", f"{sub['학점'].sum()}학점")
    c2.metric("과목 수", f"{len(sub)}과목")
    c3.metric("전공필수", f"{req_credit}학점")
    c4.metric("전공선택", f"{ele_credit}학점")

    show_cols = [c for c in ['출처전공', '과목명', '학점', '학년/학기', '구분'] if c in sub.columns]

    def show_table(part, drop_col):
        cols = [c for c in show_cols if c != drop_col]
        view = part[cols].copy()
        if '학점' in view.columns:
            view['학점'] = view['학점'].astype(str) + '학점'   # 문자로 바꿔 왼쪽 정렬
        widths = {'출처전공': 'medium', '과목명': 'medium',
                  '학점': 'small', '학년/학기': 'small', '구분': 'small'}
        cfg = {c: st.column_config.TextColumn(c, width=widths.get(c, 'medium'))
               for c in view.columns}
        st.dataframe(view, hide_index=True, use_container_width=True,
                     column_config=cfg)

    tab_a, tab_b, tab_c = st.tabs(["📋 구분별", "🏫 출처전공별", "📅 학년·학기별"])

    with tab_a:
        order = [g for g in ['전공필수', '전공선택'] if g in set(sub['구분'])]
        order += [g for g in sub['구분'].unique() if g not in order]
        for g in order:
            part = sub[sub['구분'] == g]
            st.markdown(f"**{g}** ({part['학점'].sum()}학점)")
            show_table(part, '구분')

    with tab_b:
        st.caption("여러 전공의 과목을 조합해 하나의 전공과정을 만든 모습입니다.")
        for g in sub['출처전공'].unique():
            part = sub[sub['출처전공'] == g]
            st.markdown(f"**{g}** ({len(part)}과목, {part['학점'].sum()}학점)")
            show_table(part, '출처전공')

    with tab_c:
        for g in sub['학년/학기'].unique():
            part = sub[sub['학년/학기'] == g]
            st.markdown(f"**{g}** ({part['학점'].sum()}학점)")
            show_table(part, '학년/학기')

    st.caption("※ 샘플은 이해를 돕기 위한 예시입니다. 실제 구성은 학생이 희망하는 교과목으로 구성할 수 있습니다.")

def render_program_overview():
    """제도 한눈에 비교: 공통 안내 + 다전공 학점별 묶음 카드 / 표 전환"""
    import re
    from html import escape as _esc

    programs = ALL_DATA.get('programs') or {}
    if not programs:
        return

    # 난이도(⭐) 기준 문구: 기준이 정해지면 아래에 입력 (비워 두면 표시 안 함)
    DIFFICULTY_NOTE = ""
    COLORS = [("#eef2ff", "#4f46e5"), ("#ecfdf5", "#059669"),
              ("#fff7ed", "#ea580c"), ("#f5f3ff", "#7c3aed")]

    def short_credit(text):
        m = re.search(r'(\d+(?:\s*[~\-]\s*\d+)?)\s*학점', str(text))
        return m.group(1).replace(" ", "") if m else ""

    def clean(v):
        s = str(v).replace("학점", "").strip()
        return s[:-2] if s.endswith(".0") else s

    def md(t):  # 마크다운에서 ~ 가 취소선으로 바뀌는 것 방지
        return str(t).replace("~", "\\~")

    # ---------- 제목 ----------
    st.markdown(
        '<p style="font-size:clamp(1.5rem, 4vw, 2rem); margin-bottom:12px; color:#1f2937; font-weight:600;">'
        '📊 제도 한눈에 비교</p>', unsafe_allow_html=True)

    # ---------- 공통 안내 (모든 제도가 같은 값일 때만 위로 올림) ----------
    def common(key):
        vals = {str(i.get(key, '-')).strip() for i in programs.values()}
        return vals.pop() if len(vals) == 1 else None

    common_q = common('qualification')
    common_c = common('graduation_certification')
    parts = []
    if common_q and common_q != '-':
        parts.append(f"<b>✅ 신청자격:</b> {_esc(common_q)}")
    if common_c and common_c != '-':
        parts.append(f"<b>🎓 졸업인증:</b> {_esc(common_c)}")
    if parts:
        items_html = ''.join(
            f'<span style="white-space:normal;">{p}</span>' for p in parts)
        st.markdown(
            '<div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; '
            'padding:12px 18px; margin-bottom:14px; font-size:1.05rem; color:#334155; '
            'display:flex; flex-wrap:wrap; align-items:center; gap:6px 24px;">'
            '<b style="width:100%; color:#1e293b;">모든 제도 공통</b>'
            + items_html + '</div>',
            unsafe_allow_html=True)

    # ---------- 제도별 표시값 만들기 ----------
    items = []
    for name, info in programs.items():
        multi = clean(info.get('multi_display') or short_credit(info.get('credits_multi', '')))
        primary = clean(info.get('primary_display') or short_credit(info.get('credits_primary', '')))
        desc = str(info.get('description', '') or '')
        summary = info.get('summary') or (desc[:60] + '…' if len(desc) > 60 else desc)
        group = info.get('group') or (f"다전공 이수 기준: {multi}학점" if multi else "기타")
        items.append(dict(name=name, info=info, multi=multi, primary=primary,
                          summary=summary, group=group))

    groups = {}
    for it in items:
        groups.setdefault(it['group'], []).append(it)

    def gkey(g):  # 학점 숫자가 큰 묶음부터, 숫자 없는 묶음은 맨 뒤
        m = re.search(r'(\d+)', g)
        return -int(m.group(1)) if m else 1
    ordered = sorted(groups.items(), key=lambda kv: gkey(kv[0]))

    # ---------- 보기 방식 ----------
    view = st.radio("보기 방식", ["카드로 보기", "표로 비교"], horizontal=True,
                    key="prog_view", label_visibility="collapsed")

    if view == "카드로 보기":
        for gi, (gname, members) in enumerate(ordered):
            bg, fg = COLORS[gi % len(COLORS)]
            st.markdown(
                f'<div style="background:{bg}; color:{fg}; font-weight:700; font-size:0.95rem; '
                f'border-radius:8px; padding:6px 12px; margin:14px 0 8px 0;">{_esc(gname)}</div>',
                unsafe_allow_html=True)
            for start in range(0, len(members), 3):
                cols = st.columns(3)
                for k, it in enumerate(members[start:start + 3]):
                    info = it['info']
                    with cols[k]:
                        h = []
                        h.append(
                            '<div style="display:flex;justify-content:space-between;'
                            'align-items:flex-start;gap:8px;">'
                            f'<span style="font-size:1.05rem;font-weight:700;color:#1f2937;">🎓 {_esc(it["name"])}</span>'
                            '<span style="font-size:0.85rem;white-space:nowrap;">'
                            f'<span style="color:#6b7280;">난이도: </span><span style="color:#f59e0b;">{_esc(str(info.get("difficulty", "")))}</span></span>'
                            '</div>')
                        h.append(f'<div style="color:#6b7280;font-size:0.85rem;line-height:1.5;margin:6px 0 10px 0;">{_esc(str(it["summary"]))}</div>')
                        if it['multi']:
                            h.append(
                                f'<div style="color:{fg};font-weight:800;font-size:1.6rem;line-height:1.2;">'
                                f'{_esc(it["multi"])}<span style="font-size:0.9rem;font-weight:600;"> 학점</span>'
                                '<span style="font-size:0.8rem;font-weight:500;color:#6b7280;"> 다전공</span></div>')
                        line = []
                        if it['primary']:
                            line.append(f'<b>원전공</b> {_esc(it["primary"])}학점')
                        exam = str(info.get('graduation_exam', '-'))
                        if exam and exam != '-':
                            line.append(f'<b>졸업시험</b> {_esc(exam)}')
                        degree = str(info.get('degree', '-'))
                        if degree and degree != '-':
                            line.append(f'<b>학위표기</b> <span style="color:#2563eb;">{_esc(degree[:40])}</span>')
                        h.append('<div style="font-size:0.85rem;line-height:1.8;color:#374151;margin-top:8px;">'
                                 + '<br>'.join(line) + '</div>')
                        st.markdown(
                            '<div style="border:1px solid #e5e7eb;border-radius:12px;padding:14px;'
                            'background:white;box-shadow:0 2px 4px rgba(0,0,0,0.05);margin-bottom:6px;">'
                            + ''.join(h) + '</div>', unsafe_allow_html=True)

                        with st.expander("자세히 보기"):
                            st.markdown(md(
                                "**이수학점**\n"
                                f"- 교양: {info.get('credits_general', '-') or '-'}\n"
                                f"- 원전공: {info.get('credits_primary', '-') or '-'}\n"
                                f"- 다전공: {info.get('credits_multi', '-') or '-'}"))
                            if not common_q:
                                st.markdown(md(f"**신청자격**  \n{info.get('qualification', '-') or '-'}"))
                            if not common_c:
                                st.markdown(md(f"**졸업인증**  \n{info.get('graduation_certification', '-') or '-'}"))
                            feats = [f for f in info.get('features', []) if str(f).strip()]
                            if feats:
                                st.markdown("**특징**")
                                for f in feats:
                                    st.write(f"✔️ {f}")
                            if info.get('notes'):
                                st.warning(f"💡 {info['notes']}")
    else:
        rows = []
        for gname, members in ordered:
            for it in members:
                info = it['info']
                rows.append({
                    "제도": it['name'],
                    "다전공 학점": f"{it['multi']}학점" if it['multi'] else "-",
                    "원전공 학점": f"{it['primary']}학점" if it['primary'] else "-",
                    "졸업시험": info.get('graduation_exam', '-'),
                    "학위표기": info.get('degree', '-'),
                    "난이도": info.get('difficulty', ''),
                })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
        st.caption("※ 상세 내용은 아래 '상세 정보 조회'에서 확인하세요.")

    if DIFFICULTY_NOTE:
        st.caption(DIFFICULTY_NOTE)

# ============================================================
# 🖥️ 메인 UI
# ============================================================

def main():
    initialize_session_state()
    
    # 다크모드 대응 CSS
    st.markdown("""
    <style>
    /* 다크모드 감지 및 메뉴 색상 조정 */
    @media (prefers-color-scheme: dark) {
        [data-testid="stSidebar"] .nav-link {
            color: #fafafa !important;
        }
        [data-testid="stSidebar"] h3 {
            color: #fafafa !important;
        }
    }
    
    /* 라이트모드 */
    @media (prefers-color-scheme: light) {
        [data-testid="stSidebar"] .nav-link {
            color: #262730 !important;
        }
        [data-testid="stSidebar"] h3 {
            color: #262730 !important;
        }
    }
    
    /* Streamlit 다크모드 클래스 대응 */
    [data-theme="dark"] .nav-link {
        color: #fafafa !important;
    }
    [data-theme="dark"] h3 {
        color: #fafafa !important;
    }
    
    [data-theme="light"] .nav-link {
        color: #262730 !important;
    }
    [data-theme="light"] h3 {
        color: #262730 !important;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # 사이드바
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 10px 0;'>
            <p style='font-size: 3rem; margin-bottom: 0;'>🎓</p>
            <p style='margin-top: 0; font-size: 1.3rem; font-weight: 600;'>HKNU 다전공</p>
        </div>
        """, unsafe_allow_html=True)
        
        menu = option_menu(
            menu_title=None,
            options=["다전공 제도 안내", "다전공 비교 분석", "학사 자료실"], 
            icons=["journal-bookmark-fill", "calculator-fill", "folder-fill"],
            default_index=0,
            styles={
                "container": {"padding": "0!important"},
                "icon": {"color": "orange", "font-size": "18px"}, 
                "nav-link": {
                    "font-size": "15px", 
                    "text-align": "left", 
                    "margin": "0px"
                },
                "nav-link-selected": {
                    "background-color": "#0091FF",
                    "color": "white"
                },
            }
        )
        
        st.divider()
        
        
        # 다전공 제도 소개
        st.markdown("""
        <div style="background-color: #f0f8f5; border-left: 4px solid #11998e; 
                    padding: 15px; border-radius: 8px; margin-bottom: 10px;">
            <p style="color: #333; margin: 0 0 10px 0; font-size: 0.95rem; font-weight: 600;">
                📚 다전공 제도란?
            </p>
            <p style="color: #555; font-size: 0.82rem; margin: 0; line-height: 1.6;">
                주전공 외에 복수, 융합전공 등<br>
                다양한 학위를 취득하여<br>
                융합형 인재로 성장할 수 있도록<br>
                지원하는 유연학사제도입니다.
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # 교육학사혁신팀 연락처
        st.markdown("""
        <div style="background-color: #fff3e0; border-left: 4px solid #ff9800; 
                    padding: 12px; border-radius: 8px; margin-bottom: 12px;">
            <p style="color: #333; font-size: 0.8rem; margin: 0; line-height: 1.5;">
                📞 <strong>교육학사혁신팀</strong><br>
                <span style="color: #555; font-size: 0.75rem;">031-670-5035</span>
            </p>
        </div>
        """, unsafe_allow_html=True)
        
    
    if menu == "다전공 제도 안내":
        # st.markdown("""
        # <p style="font-size: 2rem; margin-bottom: 20px; color: #1f2937; font-weight: 600;">
        #     📊 제도 한눈에 비교
        # </p>
        # """, unsafe_allow_html=True)
        
        # # 제도 비교 카드
        # if 'programs' in ALL_DATA and ALL_DATA['programs']:
        #     cols = st.columns(3)
        #     for idx, (program, info) in enumerate(ALL_DATA['programs'].items()):
        #         with cols[idx % 3]:
        #             desc = info.get('description', '')[:50] + '...' if len(info.get('description', '')) > 50 else info.get('description', '-')
        #             qual = info.get('qualification', '-')[:30] + '...' if len(str(info.get('qualification', '-'))) > 30 else info.get('qualification', '-')
                    
        #             html = f"""<div style="border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px; background: white; box-shadow: 0 2px 4px rgba(0,0,0,0.05); min-height: 400px; margin-bottom: 12px;"><p style="margin: 0 0 8px 0; color: #1f2937; font-size: 1rem; font-weight: 600;">🎓 {program}</p><p style="color: #6b7280; font-size: 11px; margin-bottom: 10px; line-height: 1.4;">{desc}</p><hr style="margin: 8px 0; border-top: 1px solid #e5e7eb;"><div style="font-size: 12px; margin-bottom: 8px;"><strong>📖 이수학점</strong><br><span style="font-size: 11px; line-height: 1.6;">• 본전공: {info.get('credits_primary', '-')}<br>• 다전공: {info.get('credits_multi', '-')}</span></div><div style="font-size: 12px; margin-bottom: 6px;"><strong>✅ 신청자격</strong><br><span style="font-size: 11px; color: #4b5563;">{qual}</span></div><div style="font-size: 12px; margin-bottom: 6px;"><strong>🎓 졸업요건</strong><br><span style="font-size: 11px;">졸업인증: {info.get('graduation_certification', '-')}<br>졸업시험: {info.get('graduation_exam', '-')}</span></div><div style="font-size: 12px; margin-bottom: 6px;"><strong>📜 학위표기</strong><br><span style="font-size: 11px; color: #2563eb;">{str(info.get('degree', '-'))[:30]}</span></div><div style="text-align: right; margin-top: 10px;"><span style="font-size: 11px;">난이도: </span><span style="color: #f59e0b;">{info.get('difficulty', '⭐⭐⭐')}</span></div></div>"""
        #             st.markdown(html, unsafe_allow_html=True)
        
        # st.divider()
        render_program_overview()

        st.divider()
        st.markdown('<p style="font-size:clamp(1.5rem, 4vw, 2rem); font-weight:600; color:#1f2937; margin:20px 0 16px 0;">🔍 상세 정보 조회</p>', unsafe_allow_html=True)
        st.caption("궁금한 제도를 선택하면 기본 정보와 특징, 전공, 교과목 목록을 볼 수 있습니다.")

        prog_keys = list(ALL_DATA['programs'].keys()) if 'programs' in ALL_DATA else []
        selected_program = st.selectbox("제도 선택", prog_keys)
        
        if selected_program:
            info = ALL_DATA['programs'][selected_program]
            
            tab1, tab2 = st.tabs(["📝 기본 정보", "✅ 특징"])
            with tab1:
                col1, col2 = st.columns([2, 1])
                with col1:
                    st.info(f"**개요**\n\n{info.get('description', '-')}")
                    
                    credits_text = f"""**이수학점**
- 교양: {info.get('credits_general', '-')}
- 원전공: {info.get('credits_primary', '-')}
- 다전공: {info.get('credits_multi', '-')}"""
                    st.markdown(credits_text)
                    
                    graduation_text = f"""**졸업요건**
- 졸업인증: {info.get('graduation_certification', '-')}
- 졸업시험: {info.get('graduation_exam', '-')}"""
                    st.markdown(graduation_text)
                with col2:
                    st.success(f"**신청자격**\n\n{info.get('qualification', '-')}")
                    st.write(f"**학위표기**: {info.get('degree', '-')}")
            with tab2:
                for f in info.get('features', []):
                    st.write(f"✔️ {f}")
                if info.get('notes'):
                    st.warning(f"💡 {info['notes']}")
            
            st.divider()
             
            # 학생설계전공: 전공 목록 대신 샘플 교육과정 표시
            if selected_program == "학생설계전공":
                render_student_design_samples()
                st.stop()

            # 전공 목록
            available_majors = {}
            
            def match_program_type(type_value, selected_prog):
                type_str = str(type_value).strip()
                if "소단위" in selected_prog or "마이크로" in selected_prog:
                    return any(kw in type_str.lower() for kw in ['소단위', '마이크로', 'md'])
                if selected_prog == "부전공":
                    return "부전공" in type_str and "융합부전공" not in type_str
                if selected_prog == "융합전공":
                    return "융합전공" in type_str
                return selected_prog in type_str
            
            if not COURSES_DATA.empty and '제도유형' in COURSES_DATA.columns:
                mask = COURSES_DATA['제도유형'].apply(lambda x: match_program_type(x, selected_program))
                for major in COURSES_DATA[mask]['전공명'].unique():
                    available_majors[major] = None
            
            if not MAJORS_INFO.empty and '제도유형' in MAJORS_INFO.columns:
                mask = MAJORS_INFO['제도유형'].apply(lambda x: match_program_type(x, selected_program))
                for _, row in MAJORS_INFO[mask].iterrows():
                    if selected_program == "융합부전공":
                        continue
                    major_name = row['전공명']
                    edu_major = row.get('교육운영전공')
                    if pd.notna(edu_major) and str(edu_major).strip():
                        available_majors[major_name] = str(edu_major).strip()
                    elif major_name not in available_majors:
                        available_majors[major_name] = None
            
            if available_majors:
                target_programs = ["복수전공", "부전공", "융합전공", "융합부전공", "연계전공"]
    
                # 🔥 구분 명확히
                is_microdegree = any(sp in selected_program for sp in ["소단위", "마이크로"])
                is_linked = "연계전공" in selected_program
                is_convergence = any(sp in selected_program for sp in ["융합전공", "융합부전공"])
    
                # [수정] 카테고리 설정 로직 변경
                category_majors = {}

                if is_microdegree or is_convergence:
                    # 융합전공, 마이크로는 '전체' 하나로 통일
                    category_majors = {"전체": sorted(available_majors.keys())}
                elif is_linked:
                    # 🔥 [핵심 수정] 연계전공을 '계열' 별로 분류하는 로직 추가
                    target_col = '계열' if '계열' in MAJORS_INFO.columns else ('단과대학' if '단과대학' in MAJORS_INFO.columns else None)
                
                    if target_col:
                        for major_name in available_majors.keys():
                            # MAJORS_INFO에서 해당 전공의 행을 찾음
                            major_row = MAJORS_INFO[MAJORS_INFO['전공명'] == major_name]
                        
                            if not major_row.empty:
                                # 해당 전공의 계열 정보를 가져옴 (여러 개일 경우 첫 번째 것 사용)
                                cat_val = major_row.iloc[0].get(target_col)
                                category = str(cat_val).strip() if pd.notna(cat_val) else "기타"
                            else:
                                category = "기타"
                        
                            if category not in category_majors:
                                category_majors[category] = []
                            category_majors[category].append(major_name)
                    
                        # 딕셔너리 키 정렬 (가나다순)
                        category_majors = dict(sorted(category_majors.items()))
                    else:
                        # 계열 컬럼을 못 찾으면 전체로 표시
                        category_majors = {"전체": sorted(available_majors.keys())}
                else:
                    category_majors = get_majors_by_category(selected_program)
    
                if selected_program in target_programs:
                    # 🔥 1. 연계전공: 단일 컬럼만
                    if is_linked:
                        major_options_with_dividers = ["선택 안 함"]

                        for category in sorted(category_majors.keys()):
                            divider = f"━━━━━━ {category} ━━━━━━"
                            major_options_with_dividers.append(divider)
                            for major in sorted(category_majors[category]):
                                major_options_with_dividers.append(major)

                        selected_major = st.selectbox(
                        f"🎓 이수하려는 {selected_program}",
                        major_options_with_dividers
                        )

                        # [수정 3] 구분선 선택 시 경고 및 null 처리
                        if selected_major and "━━━" in selected_major:
                            st.warning("⚠️ 계열 구분선이 아닌 구체적인 전공명을 선택해주세요.")
                            selected_major = None
                            
                        my_primary = "선택 안 함"
                        admission_year = datetime.now(timezone(timedelta(hours=9))).year
        
                    # 🔥 2. 융합전공: 전공 + 본전공 + 학번
                    elif is_convergence or len(category_majors) <= 1:
                        col_m1, col_m2, col_m3 = st.columns([3, 3, 1.5])
                        with col_m1:
                            all_majors = []
                            for majors in category_majors.values():
                                all_majors.extend(majors)
                            selected_major = st.selectbox(f"이수하려는 {selected_program}", sorted(set(all_majors)))
                        with col_m2:
                            primary_categories = get_majors_by_category("복수전공")
                            if len(primary_categories) > 1:
                                primary_options_with_dividers = ["선택 안 함"]
                                for category in sorted(primary_categories.keys()):
                                    divider = f"━━━━━━ {category} ━━━━━━"
                                    primary_options_with_dividers.append(divider)
                                    for major in sorted(primary_categories[category]):
                                        primary_options_with_dividers.append(major)
                                my_primary = st.selectbox(
                                    "나의 본전공",
                                    primary_options_with_dividers,
                                    key=f"special_primary_{selected_program}"
                                )
                                if my_primary and "━━━" in my_primary:
                                    st.warning("⚠️ 계열 구분선이 아닌 구체적인 전공명을 선택해주세요.")
                                    my_primary = "선택 안 함"
                            else:
                                primary_list = []
                                if not PRIMARY_REQ.empty:
                                    primary_list = sorted(PRIMARY_REQ['전공명'].unique().tolist())
                                my_primary = st.selectbox("나의 본전공", ["선택 안 함"] + primary_list)
                        with col_m3:
                            admission_year = st.number_input(
                                "📅 본인 학번",
                                min_value=2020,
                                max_value=datetime.now(timezone(timedelta(hours=9))).year,
                                value=datetime.now(timezone(timedelta(hours=9))).year,
                                key=f"special_admission_year_{selected_program}"
                            )
        
                    # 🔥 3. 복수전공/부전공: 일반 처리 (기존 코드)
                    else:
                        major_options_with_dividers = ["선택 안 함"]
                        major_to_category = {}
            
                        for category in sorted(category_majors.keys()):
                            divider = f"━━━━━━ {category} ━━━━━━"
                            major_options_with_dividers.append(divider)
                            for major in sorted(category_majors[category]):
                                major_options_with_dividers.append(major)
                                major_to_category[major] = category
            
                        primary_categories = get_majors_by_category("복수전공")
                        primary_options_with_dividers = ["선택 안 함"]
            
                        for category in sorted(primary_categories.keys()):
                            divider = f"━━━━━━ {category} ━━━━━━"
                            primary_options_with_dividers.append(divider)
                            for major in sorted(primary_categories[category]):
                                primary_options_with_dividers.append(major)
            
                        col1, col2, col3 = st.columns([3, 3, 1.5])
            
                        with col1:
                            selected_major = st.selectbox(
                                f"🎓 이수하려는 {selected_program}",
                                major_options_with_dividers,
                                key=f"major_select_{selected_program}"
                            )
            
                        with col2:
                            my_primary = st.selectbox(
                                "🏠 나의 본전공",
                                primary_options_with_dividers,
                                key=f"primary_select_{selected_program}"
                            )
            
                        with col3:
                            admission_year = st.number_input(
                                "📅 본인 학번",
                                min_value=2020,
                                max_value=datetime.now(timezone(timedelta(hours=9))).year,
                                value=datetime.now(timezone(timedelta(hours=9))).year,
                                key=f"admission_year_{selected_program}"
                            )
            
                        if selected_major and "━━━" in selected_major:
                            st.warning("⚠️ 계열 구분선이 아닌 구체적인 전공명을 선택해주세요.")
                            selected_major = None
            
                        if my_primary and "━━━" in my_primary:
                            st.warning("⚠️ 계열 구분선이 아닌 구체적인 전공명을 선택해주세요.")
                            my_primary = "선택 안 함"
        
                else:
                    # 🔥 소단위전공과정(마이크로디그리) - MICRODEGREE_INFO 사용
                    field_majors = {}
                    major_to_edu_major = {}
                    
                    if not MICRODEGREE_INFO.empty and '과정명' in MICRODEGREE_INFO.columns:
                        group_column = '계열' if '계열' in MICRODEGREE_INFO.columns else None
                        
                        for _, row in MICRODEGREE_INFO.iterrows():
                            if group_column:
                                field = row.get(group_column, '기타')
                                if pd.isna(field) or str(field).strip() == '':
                                    field = '기타'
                                field = str(field).strip()
                            else:
                                field = '전체'
                            
                            course_name = row.get('과정명', '')
                            edu_major = row.get('교육운영전공', '')
                            
                            if pd.notna(edu_major) and str(edu_major).strip():
                                display_name = f"{course_name}({str(edu_major).strip()})"
                                major_to_edu_major[display_name] = str(edu_major).strip()
                            else:
                                display_name = course_name
                                major_to_edu_major[display_name] = course_name
                            
                            if field not in field_majors:
                                field_majors[field] = []
                            if display_name not in field_majors[field]:
                                field_majors[field].append(display_name)
                    
                    if field_majors and len(field_majors) > 1:
                        major_options_with_dividers = ["선택 안 함"]
                        
                        for field in sorted(field_majors.keys()):
                            divider = f"━━━━━━ {field} ━━━━━━"
                            major_options_with_dividers.append(divider)
                            for major in sorted(field_majors[field]):
                                major_options_with_dividers.append(major)
                        
                        selected_major = st.selectbox(
                            f"🎓 이수하려는 {selected_program}",
                            major_options_with_dividers,
                            key=f"micro_major_{selected_program}"
                        )
                        
                        if selected_major and "━━━" in selected_major:
                            st.warning("⚠️ 분야 구분선이 아닌 구체적인 전공명을 선택해주세요.")
                            selected_major = None
                    elif field_majors:
                        all_majors = []
                        for majors in field_majors.values():
                            all_majors.extend(majors)
                        
                        selected_major = st.selectbox(
                            f"🎓 이수하려는 {selected_program}",
                            ["선택 안 함"] + sorted(all_majors),
                            key=f"micro_major_{selected_program}"
                        )
                    else:
                        if category_majors and category_majors.get("전체"):
                            all_majors = category_majors["전체"]
                        else:
                            all_majors = sorted(available_majors.keys())
                        
                        if all_majors:
                            selected_major = st.selectbox(
                                f"🎓 이수하려는 {selected_program}",
                                all_majors,
                                key=f"micro_major_{selected_program}"
                            )
                        else:
                            st.warning(f"⚠️ {selected_program}에 해당하는 전공을 찾을 수 없습니다.")
                            selected_major = None
                    
                    my_primary = "선택 안 함"
                    admission_year = datetime.now(timezone(timedelta(hours=9))).year
                
                if selected_major:
                    if selected_program in target_programs and "연계전공" not in selected_program:
                        col_l, col_r = st.columns(2)
                        with col_l:
                            st.markdown(f'<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">🎯 {selected_program} 이수학점</p>', unsafe_allow_html=True)
                            if not GRADUATION_REQ.empty:
                                req_data = GRADUATION_REQ[
                                    (GRADUATION_REQ['전공명'] == selected_major) & 
                                    (GRADUATION_REQ['제도유형'].str.contains(selected_program, na=False))
                                ].copy()
                                if not req_data.empty:
                                    req_data['기준학번'] = pd.to_numeric(req_data['기준학번'], errors='coerce')
                                    applicable = req_data[req_data['기준학번'] <= admission_year].sort_values('기준학번', ascending=False)
                                    if not applicable.empty:
                                        row = applicable.iloc[0]
                                        st.write(f"전공필수: **{int(row.get('다전공_전공필수', 0))}**학점")
                                        st.write(f"전공선택: **{int(row.get('다전공_전공선택', 0))}**학점")
                                        st.markdown(f'<p style="font-size: 1.1rem; font-weight: 600; margin: 12px 0;">👉 합계 {int(row.get("다전공_계", 0))}학점</p>', unsafe_allow_html=True)
                        
                        with col_r:
                            st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">🏠 본전공 이수학점 변화(신입학 기준)</p>', unsafe_allow_html=True)
                            if my_primary != "선택 안 함" and not PRIMARY_REQ.empty:
                                # 신입학 기준으로 필터링
                                pri_data = PRIMARY_REQ[
                                    (PRIMARY_REQ['전공명'] == my_primary) & 
                                    (PRIMARY_REQ['입학구분'] == '신입학')
                                ].copy()
                                if not pri_data.empty:
                                    pri_data['기준학번'] = pd.to_numeric(pri_data['기준학번'], errors='coerce')
                                    pri_valid = pri_data[pri_data['기준학번'] <= admission_year].sort_values('기준학번', ascending=False)
                                    
                                    found_req = False

                                    for _, p_row in pri_valid.iterrows():
                                        if selected_program in str(p_row['제도유형']):
                                            # ✅ [수정 핵심] NaN(빈값) 처리를 위한 안전한 변환 로직
                                            def safe_int(val):
                                                try:
                                                    # 값이 없거나 NaN이면 0 반환
                                                    if pd.isna(val) or str(val).strip() == "":
                                                        return 0
                                                    # 실수형(3.0)도 정수(3)로 변환
                                                    return int(float(val))
                                                except:
                                                    return 0

                                            p_req = safe_int(p_row.get('본전공변화_전공필수'))
                                            p_sel = safe_int(p_row.get('본전공변화_전공선택'))
                                            p_total = safe_int(p_row.get('본전공변화_계'))

                                            st.write(f"전공필수: **{p_req}**학점")
                                            st.write(f"전공선택: **{p_sel}**학점")
                                            st.markdown(f'<p style="font-size: 1.1rem; font-weight: 600; margin: 12px 0;">👉 합계 {p_total}학점</p>', unsafe_allow_html=True)
                                            
                                            # 선택한 학번과 적용된 기준학번이 다르면 안내 문구 표시
                                            applied_year = int(p_row.get('기준학번', 0))
                                            if admission_year != applied_year:
                                                st.info(f"ℹ️ {applied_year}학번 기준 ({admission_year}학번 기준은 추후 업데이트 예정)")
                                            
                                            found_req = True
                                            break
                                    
                                    if not found_req:
                                        st.info("해당 학번/과정에 대한 본전공 요건 정보가 없습니다.")
                            else:
                                st.info("본전공을 선택하면 변동 학점을 확인할 수 있습니다.")
                    
                    st.divider()

                    if not MAJORS_INFO.empty and '전공설명' in MAJORS_INFO.columns:
                        # 선택된 전공에 해당하는 행 찾기
                        desc_row = MAJORS_INFO[MAJORS_INFO['전공명'] == selected_major]
                        
                        if not desc_row.empty:
                            # 전공설명 값 가져오기
                            description = desc_row.iloc[0].get('전공설명')
                            
                            # 내용이 비어있지 않다면(NaN이나 빈 문자열이 아니면) 출력
                            if pd.notna(description) and str(description).strip():
                                st.markdown(f'<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📘 ({selected_program}) {selected_major} 전공 소개</p>', unsafe_allow_html=True)
                                st.info(str(description).strip())

                    if selected_program == "융합전공":
                        st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">📋 이수체계도</p>', unsafe_allow_html=True)
                        display_curriculum_image(selected_major, selected_program)
                        display_courses(selected_major, selected_program)
                    elif "소단위" in selected_program or "마이크로" in selected_program:
                        st.markdown('<p style="font-size: 1.3rem; font-weight: 600; margin: 20px 0 16px 0;">🖼️ 과정 안내 이미지</p>', unsafe_allow_html=True)
                        display_curriculum_image(selected_major, selected_program)
                        display_courses(selected_major, selected_program)
                    else:
                        display_courses(selected_major, selected_program)
            else:
                st.warning(f"⚠️ {selected_program}에 해당하는 전공 목록을 찾을 수 없습니다.")
                st.info("💡 데이터 파일에 해당 제도의 전공 정보가 있는지 확인해주세요.")

    # 🎯 다전공 비교 분석
    elif menu == "다전공 비교 분석":
        from simulation import render_simulation_page
        render_simulation_page()

    # 학사 서식
    elif menu == "학사 자료실":
        from forms_page import render_forms_page
        render_forms_page()
        
if __name__ == "__main__":
    initialize_session_state()
    main()
