"""
============================================================
📁 학사 자료실 (forms_page.py)
============================================================
- 메뉴: "학사 자료실" (multimajor.py 에서 render_forms_page() 호출)
- 목록: data/forms.xlsx  (forms 시트 / tabs 시트)
- 파일: forms/ 폴더 (서식 HWP·PDF, 홍보물 PDF)
- 홍보물: forms 시트의 '미리보기' 칸이 Y 이면 PDF 를 화면에 이미지로 보여 줌
          (pymupdf 가 없으면 자동으로 'PDF 받기' 버튼만 표시)
- 엑셀만 고치면 화면이 바뀝니다. 학생 제출 기능은 없습니다.
============================================================
"""

import os
import re
import unicodedata
import pandas as pd
import streamlit as st

try:                       # 홍보물 PDF → 화면 이미지 변환용 (requirements.txt: pymupdf)
    try:
        import pymupdf as fitz      # 새 버전 이름
    except ImportError:
        import fitz                 # 옛 버전 이름
    HAS_PDF_VIEW = True
except Exception:
    fitz = None
    HAS_PDF_VIEW = False

# ------------------------------------------------------------
# 설정
# ------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
FORMS_XLSX = os.path.join(BASE_DIR, "data", "forms.xlsx")
FORMS_DIR = os.path.join(BASE_DIR, "forms")
DEFAULT_NOTICE_URL = "https://www.hknu.ac.kr/kor/562/subview.do"
DEFAULT_NOTICE_NAME = "학사공지"
PLACEHOLDER = "[입력]"          # 엑셀에서 아직 정하지 못한 칸의 표시

MIME = {"hwp": "application/x-hwp", "pdf": "application/pdf"}

# 홍보물 화면 표시 설정
SIDE_BY_SIDE_MAX = 2        # 세로형 이 쪽 수 이하 → 나란히 / 넘거나 가로형이 섞이면 → 쪽 탭
OPEN_FIRST_PROMO = True     # 첫 번째 홍보물은 펼친 상태로 시작
RENDER_WIDTH = {"portrait": 1200, "landscape": 1600}   # 변환 이미지 가로 픽셀
JPEG_QUALITY = 82
PORTRAIT_COLS = [1, 2, 1]   # 세로형 한 쪽을 가운데에 놓을 때 비율(가운데가 이미지)

NOTICE_TEXT = (
    "📌 서식은 수시로 개정될 수 있습니다. **작성 전에 '공고' 버튼을 눌러 최신 서식인지 확인**해 주세요. "
    "공고에 서식이 따로 첨부되어 있으면 공고의 서식을 사용하세요. "
    "이 사이트에서는 서식을 받기만 할 수 있으며, 작성한 서식은 각 제출처에 직접 제출해 주세요."
)


# ------------------------------------------------------------
# 작은 도우미 함수
# ------------------------------------------------------------
def _clean(v):
    """엑셀 값 → 앞뒤 공백 없는 문자열 (빈 값은 '')"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "nat", "none") else s


def _yes(v):
    return _clean(v).upper() in ("Y", "예", "사용")


def _to_int(v, default=9999):
    """'1', '1.0', 1 → 1 / 숫자가 아니면 default"""
    try:
        return int(float(_clean(v)))
    except (ValueError, TypeError):
        return default


def _md(text):
    """마크다운에서 ~ 와 $ 가 서식으로 바뀌는 것을 막음"""
    return str(text).replace("~", "\\~").replace("$", "\\$")


def _is_url(v):
    return str(v).lower().startswith(("http://", "https://"))


def _fmt_date(v):
    """개정일 표시: 2025.12.03. 형태로 통일 (텍스트·날짜 모두 처리)"""
    s = _clean(v)
    m = re.match(r"^(\d{4})[-./](\d{1,2})[-./](\d{1,2})", s)
    if m:
        return f"{m.group(1)}.{int(m.group(2)):02d}.{int(m.group(3)):02d}."
    return s


def _download_name(form_name, rev_date, ext):
    """학생이 받는 파일명: 서식명_개정일.확장자 (쉼표·공백·특수문자 정리)"""
    name = re.sub(r"\s*,\s*", "-", form_name)
    name = re.sub(r"[\\/:*?\"<>|()]", "", name)
    name = re.sub(r"\s+", "_", name).strip("_")
    date = _fmt_date(rev_date).rstrip(".")
    base = f"{name}_{date}" if date else name
    return f"{base}.{ext}"


def _norm_key(filename):
    """파일명 비교용: 한글 자모 분리(맥) 통일 + 대소문자 무시"""
    return unicodedata.normalize("NFC", filename).lower()


def _file_index():
    """forms/ 폴더의 실제 파일 {비교용이름: 실제이름}"""
    if not os.path.isdir(FORMS_DIR):
        return {}
    return {_norm_key(f): f for f in os.listdir(FORMS_DIR)
            if os.path.isfile(os.path.join(FORMS_DIR, f))}


def _find_file(filename, index):
    """엑셀에 적힌 파일명 → 실제 파일 경로 (없으면 None)"""
    if not filename:
        return None
    real = index.get(_norm_key(filename))
    return os.path.join(FORMS_DIR, real) if real else None


@st.cache_data(show_spinner=False)
def _read_bytes(path, mtime):
    """파일 내용 읽기 (mtime 이 바뀌면 다시 읽음)"""
    with open(path, "rb") as f:
        return f.read()


def _mtime(path):
    try:
        return os.path.getmtime(path)
    except OSError:
        return 0


def _folder_mtime():
    """forms/ 폴더에 파일을 넣거나 바꾸면 값이 달라짐 (목록 다시 읽기용)"""
    if not os.path.isdir(FORMS_DIR):
        return 0
    stamps = [_mtime(FORMS_DIR)]
    for f in os.listdir(FORMS_DIR):
        stamps.append(_mtime(os.path.join(FORMS_DIR, f)))
    return max(stamps)


# ------------------------------------------------------------
# 홍보물 PDF → 이미지 (한 번 변환하면 기억, PDF 를 바꾸면 다시 변환)
# ------------------------------------------------------------
def _pix_to_jpeg(pix):
    try:
        return pix.tobytes("jpeg", jpg_quality=JPEG_QUALITY)
    except TypeError:                      # 오래된 pymupdf: 품질 지정 불가
        return pix.tobytes("jpeg")
    except Exception:                      # jpeg 미지원 → png
        return pix.tobytes("png")


@st.cache_data(show_spinner=False)
def _pdf_page_count(path, mtime):
    doc = fitz.open(path)
    try:
        return len(doc)
    finally:
        doc.close()


@st.cache_data(show_spinner=False, max_entries=30)
def _pdf_pages(path, mtime):
    """PDF → [(이미지 bytes, 가로형 여부), ...]"""
    pages = []
    doc = fitz.open(path)
    try:
        for page in doc:
            w, h = page.rect.width, page.rect.height
            landscape = w > h
            target = RENDER_WIDTH["landscape" if landscape else "portrait"]
            zoom = min(target / w, 4.0) if w else 1.0
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
            pages.append((_pix_to_jpeg(pix), landscape))
    finally:
        doc.close()
    return pages


# ------------------------------------------------------------
# 데이터 로드
# ------------------------------------------------------------
@st.cache_data(show_spinner=False)
def _load_forms(xlsx_mtime, folder_mtime=0):
    """forms.xlsx → (탭 목록, 서식 목록, 점검 메시지)"""
    problems = []
    if not os.path.exists(FORMS_XLSX):
        return [], [], ["data/forms.xlsx 파일이 없습니다."]
    try:
        forms_df = pd.read_excel(FORMS_XLSX, sheet_name="forms", dtype=str)
        tabs_df = pd.read_excel(FORMS_XLSX, sheet_name="tabs", dtype=str)
    except Exception as e:
        return [], [], [f"forms.xlsx 를 읽지 못했습니다: {e}"]

    # ----- 탭 -----
    tabs = []
    for _, r in tabs_df.iterrows():
        name = _clean(r.get("탭이름"))
        if not name:
            continue
        link = _clean(r.get("공고게시판링크"))
        tabs.append({
            "name": name,
            "desc": _clean(r.get("탭설명")),
            "board": _clean(r.get("공고게시판명")) or DEFAULT_NOTICE_NAME,
            "link": link if _is_url(link) else DEFAULT_NOTICE_URL,
            "order": _to_int(r.get("탭순서")),
        })
    tabs.sort(key=lambda t: t["order"])
    tab_names = {t["name"] for t in tabs}

    # ----- 서식·홍보물 -----
    index = _file_index()
    forms = []
    seen_codes = set()
    need_viewer = False
    for _, r in forms_df.iterrows():
        if not _yes(r.get("사용")):
            continue
        code = _clean(r.get("서식코드"))
        name = _clean(r.get("서식명"))
        tab = _clean(r.get("탭"))
        if not name or not tab:
            problems.append(f"{code or '(코드 없음)'}: 서식명 또는 탭이 비어 있습니다.")
            continue
        if code and code in seen_codes:
            problems.append(f"{code}: 서식코드가 중복되어 두 번째 행은 건너뜁니다.")
            continue
        seen_codes.add(code)
        if tab not in tab_names:
            problems.append(f"{code}: 탭 '{tab}' 이(가) tabs 시트에 없어 맨 뒤에 추가했습니다.")
            tabs.append({"name": tab, "desc": "", "board": DEFAULT_NOTICE_NAME,
                         "link": DEFAULT_NOTICE_URL, "order": 9999})
            tab_names.add(tab)

        files = {}
        for ext in ("hwp", "pdf"):
            fname = _clean(r.get(f"{ext}파일"))
            if not fname:
                continue
            path = _find_file(fname, index)
            if path:
                files[ext] = path
            else:
                problems.append(f"{code}: {ext.upper()} 파일을 찾을 수 없습니다 → forms/{fname}")

        if not files:
            problems.append(f"{code}: 받을 수 있는 파일이 하나도 없습니다(학생 화면에는 '준비 중'으로 표시).")

        preview = _yes(r.get("미리보기"))
        if preview:
            need_viewer = True
            if "pdf" not in files:
                problems.append(f"{code}: 미리보기=Y 인데 PDF 파일이 없어 화면에 보여 줄 수 없습니다.")

        notice = _clean(r.get("공고링크"))
        submit = _clean(r.get("제출처"))
        forms.append({
            "code": code or name,
            "name": name,
            "tab": tab,
            "desc": _clean(r.get("설명")),
            "program": _clean(r.get("관련제도")),
            "submit": "" if submit == PLACEHOLDER else submit,
            "files": files,
            "notice": notice if _is_url(notice) else "",
            "rev": _fmt_date(r.get("개정일")),
            "order": _to_int(r.get("순서")),
            "preview": preview,
        })
    forms.sort(key=lambda f: f["order"])

    if need_viewer and not HAS_PDF_VIEW:
        problems.append("pymupdf 가 설치되지 않아 홍보물은 'PDF 받기' 버튼만 표시합니다 "
                        "(requirements.txt 에 pymupdf 한 줄 추가).")

    # 서식이 하나도 없는 탭은 숨김
    used = {f["tab"] for f in forms}
    tabs = [t for t in tabs if t["name"] in used]

    # 점검 결과는 터미널(배포 시 로그)에도 남김
    if problems:
        print("[학사 자료실] 점검 필요 항목:")
        for p in problems:
            print("  -", p)
    return tabs, forms, problems


# ------------------------------------------------------------
# 화면 조각
# ------------------------------------------------------------
def _card():
    try:
        return st.container(border=True)
    except TypeError:          # 오래된 Streamlit 버전
        return st.container()


def _link_button(label, url):
    if hasattr(st, "link_button"):
        st.link_button(label, url, use_container_width=True)
    else:
        st.markdown(f"[{label}]({url})")


def _toggle(label, key, value):
    if hasattr(st, "toggle"):
        return st.toggle(label, key=key, value=value)
    return st.checkbox(label, key=key, value=value)


def _desc_line(form):
    """설명(제출처: ○○ 사무실)"""
    line = form["desc"]
    if form["submit"]:
        line = f"{line}(제출처: {form['submit']})" if line else f"제출처: {form['submit']}"
    return line


def _download(form, ext, label, key_prefix):
    path = form["files"][ext]
    st.download_button(
        label,
        data=_read_bytes(path, _mtime(path)),
        file_name=_download_name(form["name"], form["rev"], ext),
        mime=MIME[ext],
        key=f"{key_prefix}_{form['code']}_{ext}",
        use_container_width=True,
    )


def _render_form_card(form, tab):
    with _card():
        st.markdown(f"**📄 {_md(form['name'])}**")
        line = _desc_line(form)
        if line:
            st.caption(_md(line))

        # 공고 버튼: 서식별 공고링크가 있으면 그 링크, 없으면 탭의 게시판 주소
        notice_url = form["notice"] or tab["link"]

        if not form["files"]:
            st.caption("🛠️ 서식 파일 준비 중입니다. 공고를 확인해 주세요.")
            _link_button("📢 공고", notice_url)
            return

        # 있는 버튼만 한 줄로 배치 (HWP / PDF / 공고)
        exts = [(ext, label) for ext, label in (("hwp", "📥 HWP"), ("pdf", "📥 PDF"))
                if form["files"].get(ext)]
        cols = st.columns(len(exts) + 1)
        for col, (ext, label) in zip(cols, exts):
            with col:
                _download(form, ext, label, f"dl_{tab['name']}")
        with cols[-1]:
            _link_button("📢 공고", notice_url)


def _show_image(data, caption):
    st.image(data, caption=caption)


def _show_pdf_pages(form):
    """홍보물 PDF 를 쪽별 이미지로 표시 (2쪽 이하 세로형은 나란히, 그 외는 쪽 탭)"""
    path = form["files"]["pdf"]
    try:
        with st.spinner("홍보물을 불러오는 중입니다..."):
            pages = _pdf_pages(path, _mtime(path))
    except Exception as e:
        print(f"[학사 자료실] {form['code']}: 미리보기 변환 실패 → {e}")
        st.caption("⚠️ 화면 미리보기를 만들지 못했습니다. 'PDF 받기'로 확인해 주세요.")
        return
    if not pages:
        return

    n = len(pages)
    all_portrait = not any(landscape for _, landscape in pages)
    st.caption("🔍 이미지에 마우스를 올리면 오른쪽 위에 크게 보기 버튼이 나타납니다. "
               "닫을 때는 오른쪽 위 ✕ 또는 Esc 키를 누르세요(뒤로 가기를 누르면 사이트를 벗어납니다).")

    if all_portrait and n <= SIDE_BY_SIDE_MAX:
        cols = st.columns(n) if n > 1 else [st.columns(PORTRAIT_COLS)[1]]
        for i, (col, (img, _)) in enumerate(zip(cols, pages), start=1):
            with col:
                _show_image(img, f"{i}/{n}쪽" if n > 1 else None)
        return

    for i, (ui, (img, landscape)) in enumerate(
            zip(st.tabs([f"{i}쪽" for i in range(1, n + 1)]), pages), start=1):
        with ui:
            if landscape:
                _show_image(img, f"{i}/{n}쪽")
            else:
                with st.columns(PORTRAIT_COLS)[1]:
                    _show_image(img, f"{i}/{n}쪽")


def _render_promo_card(form, tab, open_default):
    """홍보물 카드: 제목·설명 + PDF 받기 + (화면으로 보기)"""
    pdf = form["files"].get("pdf")
    with _card():
        title = f"**🖼️ {_md(form['name'])}**"
        if pdf and HAS_PDF_VIEW:
            try:
                title += f"  ·  {_pdf_page_count(pdf, _mtime(pdf))}쪽"
            except Exception:
                pass
        st.markdown(title)
        line = _desc_line(form)
        if line:
            st.caption(_md(line))

        if not pdf:
            st.caption("🛠️ 홍보물 파일 준비 중입니다.")
            return

        # 버튼은 왼쪽에 작게 (PDF 받기 / 관련 공고)
        cols = st.columns(4)
        with cols[0]:
            _download(form, "pdf", "📥 PDF 받기", f"pr_{tab['name']}")
        if form["notice"]:
            with cols[1]:
                _link_button("📢 관련 공고", form["notice"])

        if HAS_PDF_VIEW:
            if _toggle("👀 화면으로 보기", key=f"pv_{tab['name']}_{form['code']}",
                       value=open_default):
                _show_pdf_pages(form)


def _match(form, query):
    if not query:
        return True
    q = query.replace(" ", "").lower()
    hay = " ".join([form["name"], form["desc"], form["program"], form["code"]])
    return q in hay.replace(" ", "").lower()


# ------------------------------------------------------------
# 메인: multimajor.py 에서 호출
# ------------------------------------------------------------
def render_forms_page():
    st.markdown(
        '<p style="font-size:clamp(1.5rem, 4vw, 2rem); margin-bottom:12px; '
        'color:#1f2937; font-weight:600;">📁 학사 자료실</p>',
        unsafe_allow_html=True)
    tabs, forms, problems = _load_forms(_mtime(FORMS_XLSX), _folder_mtime())

    # 관리자 점검: 주소 끝에 ?admin=1 을 붙이면 점검 목록이 보임
    try:
        is_admin = st.query_params.get("admin") == "1"
    except Exception:
        is_admin = False
    if is_admin and problems:
        with st.expander(f"🔧 관리자 점검 ({len(problems)}건)", expanded=True):
            for p in problems:
                st.write(f"- {p}")

    if not forms:
        st.info("등록된 서식이 아직 없습니다. 학교 홈페이지 학사공지를 확인해 주세요.")
        _link_button(f"📢 {DEFAULT_NOTICE_NAME} 바로가기", DEFAULT_NOTICE_URL)
        return

    st.info(NOTICE_TEXT)

    query = st.text_input("자료 검색", placeholder="🔍 서식명·제도명으로 검색 (예: 복수전공, 포기서)",
                          key="forms_query", label_visibility="collapsed").strip()

    labels = []
    for t in tabs:
        n = sum(1 for f in forms if f["tab"] == t["name"] and _match(f, query))
        labels.append(f"{t['name']} ({n})")

    for t, ui in zip(tabs, st.tabs(labels)):
        with ui:
            if t["desc"]:
                st.caption(_md(t["desc"]))

            items = [f for f in forms if f["tab"] == t["name"] and _match(f, query)]
            if not items:
                st.caption("검색 결과가 없습니다." if query else "등록된 자료가 없습니다.")
                continue

            # 서식: 2열 배치 (신청서·포기서가 순서대로 한 줄에 나란히 놓임)
            grid = [f for f in items if not f["preview"]]
            for i in range(0, len(grid), 2):
                row_cols = st.columns(2)
                for col, f in zip(row_cols, grid[i:i + 2]):
                    with col:
                        _render_form_card(f, t)

            # 홍보물(미리보기=Y): 화면 폭 전체 사용, 첫 번째만 펼쳐서 시작
            promos = [f for f in items if f["preview"]]
            for k, f in enumerate(promos):
                _render_promo_card(f, t, open_default=(OPEN_FIRST_PROMO and k == 0 and not query))

    st.caption("※ 자료 내용은 참고용입니다. 최종 기준은 학칙·학사운영규정·학교 공지사항입니다.")