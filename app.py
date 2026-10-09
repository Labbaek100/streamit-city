import streamlit as st
import os
import json
import io
from datetime import datetime
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from google import genai
from google.genai import types
from google.genai.errors import APIError

# python-docx optional import for Word document downloads
try:
    import docx
    from docx.shared import Pt, Inches, RGBColor
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

# ==========================================
# 1. PAGE CONFIGURATION & SESSION STATE
# ==========================================
st.set_page_config(
    page_title="CityCraft AI",
    page_icon="🏙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dark/Light Theme state initialization
if "theme" not in st.session_state:
    st.session_state.theme = "dark"

# Chat history and module-specific data initialization
if "messages" not in st.session_state:
    st.session_state.messages = []
if "concept_output" not in st.session_state:
    st.session_state.concept_output = {}
if "scenario_feedback" not in st.session_state:
    st.session_state.scenario_feedback = {}

# Theme toggler
def toggle_theme():
    st.session_state.theme = "light" if st.session_state.theme == "dark" else "dark"

IS_DARK = st.session_state.theme == "dark"

# Reset all session state data
def reset_session():
    st.session_state.messages = []
    st.session_state.concept_output = {}
    st.session_state.scenario_feedback = {}
    st.toast("학습 진행 상황이 초기화되었습니다.", icon="🔄")

# ==========================================
# 2. DOCUMENT GENERATION HELPERS
# ==========================================
def create_docx_document(title: str, subtitle: str, sections: list) -> io.BytesIO:
    """Word (.docx) 문서 생성 헬퍼 함수"""
    if not HAS_DOCX:
        return None
    
    doc = docx.Document()
    
    # Document Title
    h1 = doc.add_heading(title, level=0)
    
    # Subtitle / Metadata
    meta_p = doc.add_paragraph()
    meta_p.add_run(f"{subtitle}\n").bold = True
    meta_p.add_run(f"생성 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} | 작성 플랫폼: CityCraft AI").italic = True
    doc.add_paragraph("―" * 40)
    
    for sec in sections:
        level = sec.get("level", 1)
        doc.add_heading(sec.get("heading", ""), level=level)
        content = sec.get("content", "")
        
        for line in content.split("\n"):
            line = line.strip()
            if not line:
                continue
            if line.startswith("### "):
                doc.add_heading(line.replace("### ", ""), level=3)
            elif line.startswith("## "):
                doc.add_heading(line.replace("## ", ""), level=2)
            elif line.startswith("# "):
                doc.add_heading(line.replace("# ", ""), level=1)
            elif line.startswith(("- ", "* ")):
                doc.add_paragraph(line[2:].strip(), style='List Bullet')
            elif len(line) > 2 and line[0].isdigit() and line[1] == '.':
                doc.add_paragraph(line[2:].strip(), style='List Number')
            else:
                doc.add_paragraph(line)
        doc.add_paragraph("")
        
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

def build_portfolio_markdown() -> str:
    """전체 학습 내용(개념 교안, 시나리오 평가, AI 튜터링)을 취합한 종합 마크다운 보고서 생성"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    lines = [
        "# 🏙️ CityCraft AI 도시계획 종합 학습 포트폴리오",
        f"> 생성 일시: {timestamp}",
        "> 플랫폼: CityCraft AI (도시계획 대학 인터랙티브 학습 공간)\n",
        "---",
        "\n## 📌 1단계: 핵심 개념 교안 학습 내역\n"
    ]
    
    if st.session_state.concept_output:
        for concept_name, content in st.session_state.concept_output.items():
            lines.append(f"### [개념] {concept_name}\n")
            lines.append(content)
            lines.append("\n---\n")
    else:
        lines.append("*진행된 개념 학습 내역이 없습니다.*\n\n---\n")
        
    lines.append("## 🏙️ 2단계: 가상 도시 문제해결 시나리오 평가 리포트\n")
    if st.session_state.scenario_feedback:
        for sc_name, data in st.session_state.scenario_feedback.items():
            lines.append(f"### [시나리오] {sc_name}\n")
            lines.append("#### 📝 학생 제출 제안서")
            lines.append(data.get("proposal", ""))
            lines.append("\n#### 🎓 전문가(AI 기술사) 정밀 평가 및 피드백")
            lines.append(data.get("feedback", ""))
            lines.append("\n---\n")
    else:
        lines.append("*제출 및 평가된 시나리오가 없습니다.*\n\n---\n")
        
    lines.append("## 💬 3단계: 1:1 도시계획 전문 AI 튜터 상담 대화록\n")
    if st.session_state.messages:
        for msg in st.session_state.messages:
            sender = "👤 학생" if msg["role"] == "user" else "👨‍🏫 도시계획 전문 AI 튜터"
            lines.append(f"**{sender}**:\n{msg['content']}\n")
    else:
        lines.append("*진행된 튜터링 대화 내역이 없습니다.*\n")
        
    return "\n".join(lines)

# ==========================================
# 3. DESIGN SYSTEM & CSS INJECTION
# ==========================================
bg = "#09090b" if IS_DARK else "#ffffff"
bg_subtle = "#0c0c0f" if IS_DARK else "#f9fafb"
card = "#0c0c0f" if IS_DARK else "#ffffff"
border = "#1e1e24" if IS_DARK else "#e4e4e7"
border_subtle = "#16161a" if IS_DARK else "#f0f0f2"
text = "#fafafa" if IS_DARK else "#09090b"
text_muted = "#a1a1aa" if IS_DARK else "#71717a"
text_dim = "#52525b" if IS_DARK else "#a1a1aa"
accent = "#2563eb"
shadow = "none" if IS_DARK else "0 1px 3px rgba(0,0,0,0.04), 0 1px 2px rgba(0,0,0,0.03)"

custom_css = f"""
<style>
    /* Hide default Streamlit elements */
    header[data-testid="stHeader"], #MainMenu, footer, [data-testid="stToolbar"],
    [data-testid="stDecoration"], [data-testid="stStatusWidget"], .stDeployButton {{
        display: none !important;
    }}
    
    /* Apply base styling to app body */
    html, body, [data-testid="stAppViewContainer"], [data-testid="stApp"], .main, .block-container, section[data-testid="stMain"] {{
        background-color: {bg} !important;
        color: {text} !important;
        font-family: 'DM Sans', -apple-system, sans-serif !important;
    }}
    
    .block-container {{
        padding: 2rem 3rem 3rem !important;
        max-width: 1400px !important;
    }}

    /* Card styling */
    .content-card {{
        background-color: {card} !important;
        border: 1px solid {border} !important;
        border-radius: 10px !important;
        padding: 1.5rem !important;
        box-shadow: {shadow} !important;
        margin-bottom: 1.5rem !important;
    }}
    
    .metric-value {{
        font-size: 2.2rem;
        font-weight: 700;
        color: {text};
        font-family: 'JetBrains Mono', monospace;
    }}
    
    /* Navigation styling */
    [data-baseweb="tab-list"] {{
        gap: 6px !important;
        background: {bg_subtle} !important;
        border: 1px solid {border} !important;
        border-radius: 10px !important;
        padding: 4px !important;
    }}
    
    button[data-baseweb="tab"] {{
        background: transparent !important;
        color: {text_muted} !important;
        font-size: 0.85rem !important;
        font-weight: 500 !important;
        padding: 0.6rem 1.2rem !important;
        border: 1px solid transparent !important;
        border-radius: 8px !important;
        transition: all 0.2s ease !important;
    }}
    
    button[data-baseweb="tab"][aria-selected="true"] {{
        color: {text} !important;
        background: {card} !important;
        border-color: {border} !important;
    }}
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 4. SIDEBAR CONFIGURATION
# ==========================================
with st.sidebar:
    st.markdown("### 🏙️ **CityCraft AI**")
    st.markdown("대학생을 위한 도시계획 인터랙티브 학습 공간")
    st.markdown("---")
    
    # API Key Settings
    default_key = os.environ.get("GEMINI_API_KEY", "")
    api_key = st.text_input("Gemini API Key", value=default_key, type="password", help="Google AI Studio에서 발급받은 API Key를 입력하세요.")
    
    if not api_key:
        st.warning("⚠️ Google AI Studio에서 발급받은 API Key를 입력해주세요. [API Key 발급하기](https://aistudio.google.com/)")
    else:
        st.success("API Key가 준비되었습니다.", icon="🔑")
        
    st.markdown("---")
    
    # Navigation / Tab Select
    menu = st.radio(
        "학습 단계 선택",
        [
            "📌 1단계: 핵심 개념 백과 (Core Concepts)",
            "🏙️ 2단계: 가상 도시 문제해결 시나리오 (Scenario Workshop)",
            "💬 3단계: 1:1 도시계획 전문 AI 튜터 (AI Consultation)"
        ]
    )
    
    st.markdown("---")
    
    # Supported Google AI Gemini Models
    MODEL_MAP = {
        "Gemini 3.8 Flash Low (요청 권장)": "gemini-3.8-flash",
        "Gemini 3.8 Flash (최신 기본)": "gemini-3.8-flash",
        "Gemini 3.8 Flash Low (ID: gemini-3.8-flash-low)": "gemini-3.8-flash-low",
        "Gemini 2.0 Flash": "gemini-2.0-flash",
        "Gemini 1.5 Flash (레거시 표준)": "gemini-1.5-flash",
        "Gemini 1.5 Pro (레거시 전문가)": "gemini-1.5-pro",
        "직접 입력 (Custom ID)": "custom"
    }
    
    selected_model_label = st.selectbox(
        "Gemini 모델 선택",
        options=list(MODEL_MAP.keys()),
        index=0,
        help="Google AI Studio 모델 목록입니다. 요청하신 Gemini 3.8 Flash Low 버전이 기본 적용됩니다."
    )
    
    if MODEL_MAP[selected_model_label] == "custom":
        model_choice = st.text_input("모델 ID 직접 입력", value="gemini-3.8-flash", help="사용하고자 하는 정확한 Gemini 모델 ID를 입력하세요.")
    else:
        model_choice = MODEL_MAP[selected_model_label]
        
    st.caption(f"선택된 모델 ID: `{model_choice}`")
    
    st.markdown("---")
    st.markdown("### 📥 **종합 학습 포트폴리오**")
    total_items = len(st.session_state.concept_output) + len(st.session_state.scenario_feedback) + (1 if st.session_state.messages else 0)
    
    if total_items > 0:
        portfolio_md = build_portfolio_markdown()
        st.download_button(
            label="📄 종합 포트폴리오 (.md)",
            data=portfolio_md,
            file_name=f"도시계획_종합포트폴리오_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
            mime="text/markdown",
            use_container_width=True,
            key="dl_portfolio_md_sidebar"
        )
        
        if HAS_DOCX:
            # Word portfolio sections
            p_sections = []
            for c_name, c_content in st.session_state.concept_output.items():
                p_sections.append({"heading": f"[1단계 개념] {c_name}", "content": c_content, "level": 1})
            for s_name, s_data in st.session_state.scenario_feedback.items():
                p_sections.append({
                    "heading": f"[2단계 시나리오] {s_name}",
                    "content": f"### 학생 제안서\n{s_data.get('proposal', '')}\n\n### 전문가 피드백\n{s_data.get('feedback', '')}",
                    "level": 1
                })
            if st.session_state.messages:
                chat_txt = "\n\n".join([f"{'학생' if m['role']=='user' else 'AI 튜터'}: {m['content']}" for m in st.session_state.messages])
                p_sections.append({"heading": "[3단계 1:1 AI 튜터링 대화록]", "content": chat_txt, "level": 1})
                
            portfolio_docx = create_docx_document(
                title="CityCraft AI 도시계획 종합 학습 포트폴리오",
                subtitle="개념 교안, 시나리오 평가, 전문 상담 종합 리포트",
                sections=p_sections
            )
            if portfolio_docx:
                st.download_button(
                    label="📘 종합 포트폴리오 (.docx)",
                    data=portfolio_docx,
                    file_name=f"도시계획_종합포트폴리오_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="dl_portfolio_docx_sidebar"
                )
    else:
        st.caption("진행된 학습 내역이 없습니다. 단계를 진행하면 다운로드할 수 있습니다.")
        
    st.markdown("---")
    if st.button("학습 진행 리셋 (Reset)", use_container_width=True, on_click=reset_session):
        pass

# ==========================================
# 5. APP HEADER & THEME TOGGLER
# ==========================================
col_header, col_theme = st.columns([10, 2])
with col_header:
    st.title("🏙️ CityCraft AI")
    st.markdown("대학 전공 수준의 맞춤형 콘텐츠와 시나리오 기반의 양방향 피드백을 제공하는 도시계획 학습 어시스턴트입니다.")
with col_theme:
    theme_btn_label = "☀️ Light Mode" if IS_DARK else "🌙 Dark Mode"
    st.button(theme_btn_label, on_click=toggle_theme, use_container_width=True)

st.markdown("---")

# SYSTEM INSTRUCTION FOR GEMINI
SYSTEM_INSTRUCTION = """
당신은 20년 경력의 도시계획학 교수이자 도시계획 기술사입니다.
도시공학 및 도시계획을 공부하는 대학생들에게 친절하면서도 학술적 깊이를 갖춘 전문적인 설명을 제공합니다.
답변 시 실제 도시 사례, 법적/제도적 배경, 최신 스마트시티 기술 트렌드를 함께 연계하여 설명하세요.
마크다운 형식의 깔끔한 구조로 시각적으로 파악하기 쉽게 출력해주세요.
"""

def get_gemini_client():
    if not api_key:
        return None
    try:
        # Initialize the official google-genai client
        client = genai.Client(api_key=api_key)
        return client
    except Exception as e:
        st.error(f"클라이언트 초기화 오류: {e}")
        return None

def get_available_models(client):
    """API Key로 실제 사용 가능한 Gemini 모델 목록 조회"""
    try:
        models = []
        for m in client.models.list():
            name = getattr(m, 'name', '')
            if name.startswith('models/'):
                name = name[len('models/'):]
            if name:
                models.append(name)
        return models
    except Exception:
        return []

def call_gemini_api(client, contents, chosen_model, temperature=0.7):
    """
    Gemini API 호출 래퍼 함수:
    404 NOT_FOUND(모델 미지원) 발생 시 실제 API에서 사용 가능한 모델을 동적으로 탐색하여 자동 fallback 처리
    """
    config_params = {
        "system_instruction": SYSTEM_INSTRUCTION,
        "temperature": temperature
    }
    
    # ThinkingConfig 지원 시 low thinking budget 설정 시도
    if hasattr(types, "ThinkingConfig"):
        try:
            config_params["thinking_config"] = types.ThinkingConfig(thinking_budget=1024)
        except Exception:
            pass

    config = types.GenerateContentConfig(**config_params)

    # 1. 1차 시도 (요청된 모델)
    try:
        response = client.models.generate_content(
            model=chosen_model,
            contents=contents,
            config=config
        )
        return response.text
    except Exception as e:
        err_msg = str(e)
        # 만약 thinking_config 파라미터가 거부된 경우 기본 config로 1회 재시도
        if "thinking_config" in err_msg or "ThinkingConfig" in err_msg or "INVALID_ARGUMENT" in err_msg:
            try:
                base_config = types.GenerateContentConfig(
                    system_instruction=SYSTEM_INSTRUCTION,
                    temperature=temperature
                )
                response = client.models.generate_content(
                    model=chosen_model,
                    contents=contents,
                    config=base_config
                )
                return response.text
            except Exception as retry_err:
                err_msg = str(retry_err)
                e = retry_err

        # 2. 404 NOT_FOUND 발생 시 지능형 Fallback 처리
        if "404" in err_msg or "NOT_FOUND" in err_msg:
            available = get_available_models(client)
            fallback_candidates = []
            
            # 동적 모델 목록 우선 검사
            for m in available:
                if "3.8" in m and "flash" in m and m != chosen_model and m not in fallback_candidates:
                    fallback_candidates.append(m)
            for m in available:
                if "flash" in m and m != chosen_model and m not in fallback_candidates:
                    fallback_candidates.append(m)
            for m in available:
                if "gemini" in m and m != chosen_model and m not in fallback_candidates:
                    fallback_candidates.append(m)
                    
            # 정적 예비 후보군 (동적 조회가 비어있을 때 대비)
            static_fallbacks = ["gemini-3.8-flash", "gemini-3.8-flash-low", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-1.5-pro"]
            for sf in static_fallbacks:
                if sf != chosen_model and sf not in fallback_candidates:
                    fallback_candidates.append(sf)
                    
            for fb_model in fallback_candidates:
                try:
                    fb_config = types.GenerateContentConfig(
                        system_instruction=SYSTEM_INSTRUCTION,
                        temperature=temperature
                    )
                    response = client.models.generate_content(
                        model=fb_model,
                        contents=contents,
                        config=fb_config
                    )
                    st.info(f"ℹ️ 선택된 모델(`{chosen_model}`)을 사용할 수 없어 사용 가능한 모델(`{fb_model}`)로 자동 전환하여 완료했습니다.")
                    return response.text
                except Exception:
                    continue
                    
        # 모든 fallback 시도 실패 시 원래 에러 반환
        raise e

# ==========================================
# 6. NAVIGATION MODULE CONTROLLER
# ==========================================

# ------------------------------------------
# MODULE 1: CORE CONCEPTS
# ------------------------------------------
if "1단계" in menu:
    st.subheader("📌 1단계: 도시계획 핵심 개념 백과 (Core Concepts)")
    
    # Sub-tabs for Core Concepts
    concept_tabs = st.tabs([
        "🌐 토지이용 & 용도지역제",
        "🚇 대중교통지향형 개발 (TOD)",
        "🏙️ 스마트시티 & 디지털트윈",
        "🌱 지속가능 도시재생 & ESG"
    ])
    
    categories_prompt = [
        "토지이용 및 용도지역제 (Zoning & Land Use Regulation)",
        "대중교통지향형 개발 (TOD & Mobility)",
        "스마트시티 & 디지털 트윈 (Smart City & Digital Twin)",
        "지속가능 도시재생 및 환경 (Urban Regeneration & ESG)"
    ]
    
    for idx, tab in enumerate(concept_tabs):
        with tab:
            category_name = categories_prompt[idx]
            st.markdown(f"### {category_name} 학습")
            st.markdown(f"`{category_name}` 관련 세부 심화 분석 정보와 국내외 우수 사례를 분석합니다.")
            
            # Sub-concept select list
            if idx == 0:
                sub_concepts = ["용도지역 지구제 (Zoning)", "개발권양도제 (TDR)", "복합용도개발 (Mixed-Use Development)"]
            elif idx == 1:
                sub_concepts = ["대중교통지향형 개발 (TOD)", "15분 도시 (15-Minute City)", "교통수요관리 (TDM)"]
            elif idx == 2:
                sub_concepts = ["디지털 트윈 (Digital Twin)", "자율주행 및 MaaS", "스마트 그리드 & 에너지 인프라"]
            else:
                sub_concepts = ["도시재생 뉴딜사업", "바르셀로나 슈퍼블록 (Superblocks)", "저영향개발 (LID) 기법"]
                
            selected_sub = st.selectbox("탐구할 세부 개념 선택", sub_concepts, key=f"sel_{idx}")
            
            if st.button("📚 심화 설명 및 사례 생성", key=f"btn_{idx}"):
                if not api_key:
                    st.warning("사이드바에 Gemini API Key를 입력해주셔야 서비스 생성이 가능합니다.")
                else:
                    client = get_gemini_client()
                    if client:
                        with st.spinner("AI 교수가 심화 학습자료를 작성하고 있습니다..."):
                            prompt = f"""
                            도시계획 전공 대학생에게 다음 세부 개념에 대해 대학 강의 교안 수준으로 분석해주세요:
                            개념: {selected_sub} (카테고리: {category_name})
                            
                            반드시 다음 항목들을 포함하여 자세하고 학술적인 보고서 형태로 작성해주세요:
                            1. 전공 수준의 정확한 정의 및 이론적 배경
                            2. 핵심 매커니즘 / 작동 원리
                            3. 장점 및 현실적인 한계점 (실패 요인 및 법적/제도적 제약 등)
                            4. 국내외 대표 성공 사례 분석 (예: 서울시 2040 플랜, 바르셀로나 슈퍼블록 등 구체적 사업명 명시)
                            5. 미래 도시 발전 방향 제언
                            """
                            
                            try:
                                result_text = call_gemini_api(
                                    client=client,
                                    contents=prompt,
                                    chosen_model=model_choice,
                                    temperature=0.7
                                )
                                st.session_state.concept_output[selected_sub] = result_text
                            except Exception as e:
                                st.error(f"콘텐츠 생성 중 오류가 발생했습니다: {e}")
                                
            # Show cached concept output
            if selected_sub in st.session_state.concept_output:
                st.markdown("---")
                st.markdown(f"### 📖 {selected_sub} 심화 교안")
                st.markdown(st.session_state.concept_output[selected_sub])
                
                # Download buttons for Concept Material
                st.markdown("##### 📥 이 개념 교안 다운로드")
                concept_text = st.session_state.concept_output[selected_sub]
                concept_md = f"# {selected_sub} 심화 교안\n\n- **분야**: {category_name}\n- **생성 일시**: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n- **출처**: CityCraft AI\n\n---\n\n{concept_text}"
                
                col_c_d1, col_c_d2 = st.columns(2)
                with col_c_d1:
                    st.download_button(
                        label="📄 마크다운 다운로드 (.md)",
                        data=concept_md,
                        file_name=f"{selected_sub}_심화교안.md",
                        mime="text/markdown",
                        key=f"dl_concept_md_{selected_sub}",
                        use_container_width=True
                    )
                if HAS_DOCX:
                    with col_c_d2:
                        concept_docx = create_docx_document(
                            title=f"{selected_sub} 심화 교안",
                            subtitle=f"도시계획 학술 자료 ({category_name})",
                            sections=[{"heading": selected_sub, "content": concept_text, "level": 1}]
                        )
                        if concept_docx:
                            st.download_button(
                                label="📘 워드 문서 다운로드 (.docx)",
                                data=concept_docx,
                                file_name=f"{selected_sub}_심화교안.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key=f"dl_concept_docx_{selected_sub}",
                                use_container_width=True
                            )

# ------------------------------------------
# MODULE 2: SCENARIO WORKSHOP
# ------------------------------------------
elif "2단계" in menu:
    st.subheader("🏙️ 2단계: 가상 도시 문제해결 시나리오 (Scenario Workshop)")
    st.markdown("제시된 현실적 도시 문제 상황에 대해 나만의 정책 및 인프라 구상을 제출하고, 전문 기술사 수준의 피드백을 받으세요.")
    
    scenarios = {
        "시나리오 A: 구도심 쇠퇴와 청년 인구 유출 극복 방안": {
            "background": "지방 거점 대도시인 'A시'의 원도심 지역은 인구 고령화율이 25%에 도달하였고, 일자리 부족으로 청년층 유출이 매년 심화되고 있습니다. 역사 문화 자원은 풍부하지만 정주 여건이 불량하고 낙후된 저층 주거지가 밀집해 있습니다.",
            "target": "원도심 활성화 및 청년층 유입 촉진을 위한 토지이용, 산업유치, 정주환경 종합 대책 수립"
        },
        "시나리오 B: 신규 광역급행철도(GTX) 역세권 복합개발 계획": {
            "background": "수도권 배후 주거도시의 중심부에 새로운 GTX 환승역 개통이 예정되어 있습니다. 현재 역 주변은 노후화된 저밀 상가와 환승 편의성이 부족한 대중교통 거점으로 구성되어 있어, 광역 거점으로서의 입지 잠재력을 살리지 못하고 있습니다.",
            "target": "용적률 완화, 복합 용도 조닝, 대중교통 연계성 극대화 및 활성화 거점 조성 전략 개발"
        },
        "시나리오 C: 기후위기 대응 도심 열섬현상 및 침수 방지 인프라 설계": {
            "background": "매년 여름철 기록적인 폭우와 폭염으로 피해를 보는 분지 지형의 인구 밀집 구도심 지구입니다. 도로 포장율이 85%에 달하고 녹지율이 5% 미만인 저지대로서 우수 유출량이 매우 크고 열섬 현상이 극심합니다.",
            "target": "저영향개발(LID), 바람길 구축, 그린인프라 도입 및 방재 안전성 향상을 위한 기후 복원형 도시 공간 설계"
        }
    }
    
    selected_scenario_name = st.selectbox("도전할 학습 시나리오 선택", list(scenarios.keys()))
    scenario_info = scenarios[selected_scenario_name]
    
    # Styled scenario prompt container
    st.markdown(f"""
    <div class="content-card">
        <h4>📋 시나리오 상황 및 대상지 현황</h4>
        <p><strong>배경 상황:</strong> {scenario_info['background']}</p>
        <p><strong>수행 목표:</strong> {scenario_info['target']}</p>
    </div>
    """, unsafe_allow_html=True)
    
    student_proposal = st.text_area(
        "📝 본인이 구상한 도시계획 정책 및 아이디어 제안서 작성",
        placeholder="이곳에 구체적인 토지이용 계획, 인프라 구축, 교통 연계, ESG 도입안 등 제안할 아이디어를 대학 리포트 수준으로 상세하게 작성해주세요. 구체적일수록 심도 깊은 피드백이 제공됩니다.",
        height=250
    )
    
    if st.button("📊 전문가 피드백 및 제안서 평가하기"):
        if not api_key:
            st.warning("사이드바에 Gemini API Key를 입력해주셔야 평가가 가능합니다.")
        elif len(student_proposal.strip()) < 20:
            st.warning("조금 더 상세한 제안서를 입력해주세요. 최소 20자 이상 작성이 필요합니다.")
        else:
            client = get_gemini_client()
            if client:
                with st.spinner("도시계획 전문가 AI 교수가 제안서를 정밀 평가 중입니다..."):
                    prompt = f"""
                    [시나리오]
                    주제: {selected_scenario_name}
                    현황: {scenario_info['background']}
                    목표: {scenario_info['target']}
                    
                    [학생의 제안 내용]
                    {student_proposal}
                    
                    도시계획학 교수 및 도시계획 기술사 페르소나를 기반으로 위 제안서를 평가하고 대안을 분석해주세요.
                    출력 형식은 다음의 4가지 핵심 기준에 입각한 마크다운 분석표와 세부 조언 형태로 제공해야 합니다:
                    
                    1. **실현 가능성 (Feasibility)**: 정책의 현실성, 예상 소요 예산/재원 조달 가능성, 현행 법제도(국토계획법 등)와의 정합성 등
                    2. **공공성 및 주민 수용성 (Public Interest)**: 젠트리피케이션 대책, 주민 참여 절차, 공공 기여 및 혜택 배분성
                    3. **친환경/지속가능성 (Sustainability)**: 탄소 배출 저감, 그린 인프라 확보 여부, 장기적 도시 탄력성
                    4. **보완점 및 발전 방향 (Recommendations)**: 제안의 치명적 결함 지적, 아이디어를 심화하기 위해 찾아볼 만한 구체적인 논문/사례 추천
                    
                    각 기준별 평가 결과는 직관적인 평점(예: S / A / B / C / F 등급)과 함께 상세한 이유를 덧붙여 서술해주세요.
                    """
                    
                    try:
                        result_feedback = call_gemini_api(
                            client=client,
                            contents=prompt,
                            chosen_model=model_choice,
                            temperature=0.5
                        )
                        st.session_state.scenario_feedback[selected_scenario_name] = {
                            "proposal": student_proposal,
                            "feedback": result_feedback
                        }
                    except Exception as e:
                        st.error(f"피드백 생성 중 오류가 발생했습니다: {e}")
                        
    # Show cached feedback
    if selected_scenario_name in st.session_state.scenario_feedback:
        saved_data = st.session_state.scenario_feedback[selected_scenario_name]
        st.markdown("---")
        st.markdown(f"### 📋 {selected_scenario_name} 평가 결과")
        
        # User proposed plan preview
        with st.expander("내가 제출한 제안서 보기"):
            st.info(saved_data["proposal"])
            
        st.markdown("#### 🎓 AI 기술사 정밀 피드백")
        st.markdown(saved_data["feedback"])
        
        # Download buttons for Scenario Evaluation
        st.markdown("##### 📥 전문가 평가 내용 및 제안서 다운로드")
        eval_md = f"""# {selected_scenario_name} - 평가 결과서

> 생성 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  
> 평가 주관: CityCraft AI 도시계획 기술사 평가위원회  

---

## 1. 대상지 개요
- **배경 상황**: {scenario_info['background']}
- **수행 목표**: {scenario_info['target']}

---

## 2. 학생 제출 제안서
{saved_data['proposal']}

---

## 3. 전문가(AI 기술사) 정밀 평가 및 피드백
{saved_data['feedback']}
"""
        col_sc_d1, col_sc_d2 = st.columns(2)
        with col_sc_d1:
            st.download_button(
                label="📄 평가서 마크다운 다운로드 (.md)",
                data=eval_md,
                file_name=f"{selected_scenario_name}_평가결과.md",
                mime="text/markdown",
                key=f"dl_eval_md_{selected_scenario_name}",
                use_container_width=True
            )
        if HAS_DOCX:
            with col_sc_d2:
                eval_docx = create_docx_document(
                    title=f"도시계획 시나리오 평가 리포트",
                    subtitle=f"{selected_scenario_name}",
                    sections=[
                        {"heading": "1. 시나리오 대상지 현황", "content": f"배경: {scenario_info['background']}\n목표: {scenario_info['target']}", "level": 1},
                        {"heading": "2. 학생 정책 제안서", "content": saved_data['proposal'], "level": 1},
                        {"heading": "3. 전문가 정밀 평가 및 피드백", "content": saved_data['feedback'], "level": 1}
                    ]
                )
                if eval_docx:
                    st.download_button(
                        label="📘 평가서 워드 문서 다운로드 (.docx)",
                        data=eval_docx,
                        file_name=f"{selected_scenario_name}_평가결과.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_eval_docx_{selected_scenario_name}",
                        use_container_width=True
                    )

# ------------------------------------------
# MODULE 3: AI TUTOR CONSULTATION
# ------------------------------------------
else:
    st.subheader("💬 3단계: 1:1 도시계획 전문 AI 튜터 (AI Consultation)")
    st.markdown("용적률 및 건폐율 계산, 국토의 계획 및 이용에 관한 법률, 도시계획 수립 절차 등 전공 지식에 관한 어떤 질문이든 튜터링을 제공합니다.")
    
    # Download Chat History if available
    if st.session_state.messages:
        chat_md_text = f"# 💬 도시계획 전문 AI 튜터 1:1 상담 대화록\n\n> 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n---\n\n"
        for msg in st.session_state.messages:
            speaker = "👤 학생" if msg["role"] == "user" else "👨‍🏫 도시계획 전문 AI 튜터"
            chat_md_text += f"### {speaker}\n{msg['content']}\n\n"
            
        col_chat_d1, col_chat_d2 = st.columns(2)
        with col_chat_d1:
            st.download_button(
                label="📥 상담 대화록 다운로드 (.md)",
                data=chat_md_text,
                file_name=f"도시계획_튜터링_대화록_{datetime.now().strftime('%Y%m%d_%H%M')}.md",
                mime="text/markdown",
                key="dl_chat_md",
                use_container_width=True
            )
        if HAS_DOCX:
            with col_chat_d2:
                chat_docx = create_docx_document(
                    title="도시계획 전문 AI 튜터 1:1 상담 대화록",
                    subtitle="CityCraft AI 인터랙티브 튜터링 기록",
                    sections=[
                        {
                            "heading": "👤 학생 질문" if m["role"] == "user" else "👨‍🏫 AI 튜터 답변",
                            "content": m["content"],
                            "level": 2
                        }
                        for m in st.session_state.messages
                    ]
                )
                if chat_docx:
                    st.download_button(
                        label="📥 상담 대화록 다운로드 (.docx)",
                        data=chat_docx,
                        file_name=f"도시계획_튜터링_대화록_{datetime.now().strftime('%Y%m%d_%H%M')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_chat_docx",
                        use_container_width=True
                    )
        st.markdown("---")

    # Display chat history
    for message in st.session_state.messages:
        with st.chat_message(message["role"]):
            st.markdown(message["content"])
            
    # Interactive input
    if prompt := st.chat_input("도시계획학 전공 질문을 입력해보세요. (예: 주거지역 용적률 한도와 기부채납 인센티브 계산 방식은?)"):
        # Add to local chat memory
        st.session_state.messages.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
            
        # Call Gemini response
        if not api_key:
            st.error("사이드바에 Gemini API Key를 먼저 입력해주세요.")
        else:
            client = get_gemini_client()
            if client:
                with st.chat_message("assistant"):
                    message_placeholder = st.empty()
                    message_placeholder.markdown("👨‍🏫 답변을 작성하고 있습니다...")
                    
                    # Convert simple text list history to Gemini Content structure
                    contents_payload = []
                    # Keep history for context
                    for msg in st.session_state.messages:
                        contents_payload.append(
                            types.Content(
                                role=msg["role"],
                                parts=[types.Part.from_text(text=msg["content"])]
                            )
                        )
                        
                    try:
                        assistant_response = call_gemini_api(
                            client=client,
                            contents=contents_payload,
                            chosen_model=model_choice,
                            temperature=0.7
                        )
                        message_placeholder.markdown(assistant_response)
                        st.session_state.messages.append({"role": "assistant", "content": assistant_response})
                        st.rerun()
                    except Exception as e:
                        message_placeholder.empty()
                        st.error(f"답변 생성 오류가 발생했습니다: {e}")
