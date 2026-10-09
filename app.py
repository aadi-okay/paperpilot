import os
import io
import re
import streamlit as st
from dotenv import load_dotenv
from openai import OpenAI
from docx import Document
from docx.shared import Inches, Pt

# ---------------- CONFIGURATION ----------------

load_dotenv(dotenv_path=".env", override=True)

st.set_page_config(
    page_title="PaperPilot | AI Paper Generator",
    page_icon="📘",
    layout="wide",
    initial_sidebar_state="expanded",
)
# ---------------- STYLING ----------------

st.markdown("""
<style>
    /* Overall page */
    .stApp {
        background: #f6f8fc;
        color: #17233c;
    }

    .block-container {
        max-width: 1250px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background: #111c34;
        border-right: 1px solid #24324d;
    }

    [data-testid="stSidebar"] * {
        color: #f8fafc;
    }

    [data-testid="stSidebar"] hr {
        border-color: #34425d;
    }

    /* Main headings */
    h1, h2, h3 {
        color: #17233c !important;
        letter-spacing: -0.5px;
    }

    p, label, .stMarkdown {
        color: #42516b;
    }

    /* Hero banner */
    .hero {
        background: linear-gradient(120deg, #172b50 0%, #315fc1 100%);
        padding: 34px;
        border-radius: 20px;
        color: #ffffff;
        margin-bottom: 28px;
        box-shadow: 0 12px 35px rgba(32, 65, 130, 0.12);
    }

    .hero h1 {
        color: #ffffff !important;
        font-size: 2.35rem;
        font-weight: 750;
        line-height: 1.2;
        margin-bottom: 12px;
    }

    .hero p {
        color: #e5edff !important;
        font-size: 1.05rem;
        margin-bottom: 0;
    }

    /* Dashboard cards */
    .metric-card {
        background: #ffffff;
        padding: 24px;
        border: 1px solid #e5eaf3;
        border-radius: 16px;
        min-height: 125px;
        box-shadow: 0 4px 15px rgba(25, 45, 80, 0.035);
    }

    .metric-card small {
        color: #64748b;
        font-size: 0.76rem;
        font-weight: 700;
        letter-spacing: 0.7px;
    }

    .metric-card h2 {
        color: #172b50 !important;
        font-size: 1.8rem;
        margin-top: 12px;
        margin-bottom: 0;
    }

    /* Buttons */
    div.stButton > button,
    div.stDownloadButton > button {
        border-radius: 10px;
        min-height: 44px;
        font-weight: 600;
        transition: all 0.15s ease;
    }

    div.stButton > button[kind="primary"],
    div.stDownloadButton > button {
        background: #315fc1;
        color: #ffffff;
        border: 1px solid #315fc1;
    }

    div.stButton > button[kind="primary"]:hover,
    div.stDownloadButton > button:hover {
        background: #244da6;
        border-color: #244da6;
    }

    /* Form controls */
    input, textarea {
        border-radius: 9px !important;
    }

    div[data-baseweb="select"] > div {
        border-radius: 9px;
    }

    /* Expanders */
    [data-testid="stExpander"] {
        background: #ffffff;
        border: 1px solid #e5eaf3;
        border-radius: 12px;
    }

    /* Alerts */
    [data-testid="stAlert"] {
        border-radius: 12px;
    }

    /* Download buttons and forms */
    [data-testid="stForm"] {
        background: #ffffff;
        padding: 24px;
        border: 1px solid #e5eaf3;
        border-radius: 16px;
    }

    /* Small screens */
    @media (max-width: 768px) {
        .hero {
            padding: 24px;
        }

        .hero h1 {
            font-size: 1.8rem;
        }

        .block-container {
            padding-left: 1rem;
            padding-right: 1rem;
        }
    }
</style>
""", unsafe_allow_html=True)



# ---------------- SESSION STATE ----------------

if "papers" not in st.session_state:
    st.session_state.papers = []

if "current_paper" not in st.session_state:
    st.session_state.current_paper = None

# ---------------- SIDEBAR ----------------

with st.sidebar:
    st.markdown("# 📘 PaperPilot")
    st.caption("AI-powered assessment workspace")
    st.divider()

    page = st.radio(
        "WORKSPACE",
        ["Dashboard", "Create Paper", "My Papers", "Settings"],
    )

    st.divider()
    st.caption("Prototype v0.1")
    st.caption("AI-generated content requires teacher review.")

# ---------------- HELPERS ----------------

def get_client():
    api_key = os.getenv("OPENROUTER_API_KEY")
    if not api_key:
        raise RuntimeError(
            "API key not found. Check your .env file."
        )

    return OpenAI(
        base_url="https://openrouter.ai/api/v1",
        api_key=api_key,
        timeout=120.0,
        max_retries=1,
    )


def build_prompt(
    grade,
    subject,
    chapters,
    marks,
    duration,
    difficulty,
    question_types,
    include_answers,
    instructions,
):
    answer_setting = (
        "Include a separate answer key with worked solutions where appropriate."
        if include_answers
        else "Do not include answers or solutions."
    )

    return f"""
Create a DRAFT school examination question paper.

Board: CBSE
Class: {grade}
Subject: {subject}
Chapters or topics: {chapters}
Maximum marks: {marks}
Duration: {duration} minutes
Difficulty: {difficulty}
Requested question types: {", ".join(question_types)}

Additional teacher instructions:
{instructions if instructions.strip() else "None"}

Requirements:
1. Write clear, age-appropriate questions.
2. Include a title and general instructions.
3. Give every question a number and marks allocation.
4. Use only the requested chapters or topics.
5. Make the marks allocations add up to exactly {marks}.
6. Include a sensible mix of the requested question types.
7. Avoid ambiguous wording and duplicate questions.
8. Do not claim official CBSE approval or verified syllabus compliance.
9. Clearly separate the question paper and answer key.
10. If a numerical question is included, solve it and check the arithmetic.

{answer_setting}

Return readable Markdown.
Use ordinary text for equations when practical.
Before finishing, verify that the sum of all question marks is {marks}.
"""


def create_docx(title, body, include_answers):
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    styles = doc.styles
    styles["Normal"].font.name = "Arial"
    styles["Normal"].font.size = Pt(10)

    heading = doc.add_heading(title, 0)
    heading.alignment = 1

    doc.add_paragraph(
        "DRAFT — Please review questions, answers, marks and syllabus alignment."
    )

    for line in body.splitlines():
        text = line.strip()

        if not text:
            continue

        if text.startswith("# "):
            doc.add_heading(text[2:], level=1)
        elif text.startswith("## "):
            doc.add_heading(text[3:], level=2)
        elif text.startswith("### "):
            doc.add_heading(text[4:], level=3)
        elif text.startswith("- ") or text.startswith("* "):
            doc.add_paragraph(text[2:], style="List Bullet")
        else:
            text = re.sub(r"\*\*(.*?)\*\*", r"\1", text)
            text = re.sub(r"\*(.*?)\*", r"\1", text)
            doc.add_paragraph(text)

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()


def generate_paper(
    grade,
    subject,
    chapters,
    marks,
    duration,
    difficulty,
    question_types,
    include_answers,
    instructions,
):
    client = get_client()

    prompt = build_prompt(
        grade,
        subject,
        chapters,
        marks,
        duration,
        difficulty,
        question_types,
        include_answers,
        instructions,
    )

    response = client.chat.completions.create(
        model="openrouter/free",
        messages=[
            {
                "role": "system",
                "content": (
                    "You are PaperPilot, an educational drafting assistant. "
                    "Follow the teacher's constraints carefully. "
                    "Never invent a claim of official curriculum approval."
                ),
            },
            {"role": "user", "content": prompt},
        ],
    )

    return response.choices[0].message.content


# ---------------- DASHBOARD ----------------

if page == "Dashboard":
    st.markdown("""
    <div class="hero">
        <h1>Good work starts with a great paper.</h1>
        <p>Create, review and export draft question papers with AI assistance.</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            f'<div class="metric-card"><small>PAPERS CREATED</small>'
            f'<h2>{len(st.session_state.papers)}</h2></div>',
            unsafe_allow_html=True,
        )

    with col2:
        st.markdown(
            '<div class="metric-card"><small>AI WORKSPACE</small>'
            '<h2>Ready to use</h2></div>',
            unsafe_allow_html=True,
        )

    with col3:
        st.markdown(
            '<div class="metric-card"><small>EXPORT FORMAT</small>'
            '<h2>Word (.docx)</h2></div>',
            unsafe_allow_html=True,
        )

    st.subheader("Welcome to your workspace")
    st.write(
        "Generate a question-paper draft, review the content, "
        "and download it as an editable Word document."
    )

    if st.button("＋ Create a new paper", type="primary"):
        st.session_state.current_paper = None
        st.session_state.navigate_to_create = True
        st.rerun()

    st.subheader("Recent papers")

    if not st.session_state.papers:
        st.info(
            "No papers yet. Select **Create Paper** in the sidebar "
            "to generate your first draft."
        )
    else:
        for paper in reversed(st.session_state.papers[-5:]):
            with st.expander(
                f"{paper['subject']} · Class {paper['grade']} · "
                f"{paper['marks']} marks"
            ):
                st.write(paper["content"])


# ---------------- CREATE PAPER ----------------

elif page == "Create Paper":
    st.title("Create a question paper")
    st.write(
        "Configure your assessment. PaperPilot will generate a draft "
        "for you to review."
    )

    with st.form("paper_form"):
        col1, col2 = st.columns(2)

        with col1:
            grade = st.selectbox(
                "Class",
                ["6", "7", "8", "9", "10"],
                index=4,
            )

            subject = st.selectbox(
                "Subject",
                [
                    "Science",
                    "Mathematics",
                    "English",
                    "Social Science",
                    "Computer Applications",
                ],
            )

            chapters = st.text_area(
                "Chapters or topics",
                placeholder="Example: Electricity, Magnetic Effects of Electric Current",
                height=100,
            )

            marks = st.selectbox(
                "Total marks",
                [20, 40, 50, 80, 100],
                index=2,
            )

        with col2:
            duration = st.selectbox(
                "Duration",
                [30, 45, 60, 90, 120, 180],
                index=3,
                format_func=lambda x: f"{x} minutes",
            )

            difficulty = st.select_slider(
                "Difficulty",
                options=["Easy", "Medium", "Hard", "Mixed"],
                value="Mixed",
            )

            question_types = st.multiselect(
                "Question types",
                [
                    "Multiple choice",
                    "Very short answer",
                    "Short answer",
                    "Long answer",
                    "Case-based",
                    "Numerical",
                    "Assertion and reasoning",
                ],
                default=[
                    "Multiple choice",
                    "Short answer",
                    "Long answer",
                ],
            )

            include_answers = st.checkbox(
                "Include answer key",
                value=True,
            )

        instructions = st.text_area(
            "Additional instructions (optional)",
            placeholder="Example: Include 5 MCQs and at least 2 application-based questions.",
        )

        submitted = st.form_submit_button(
            "Generate question paper",
            type="primary",
            use_container_width=True,
        )

    if submitted:
        if not chapters.strip():
            st.error("Please enter at least one chapter or topic.")
        elif not question_types:
            st.error("Select at least one question type.")
        else:
            with st.spinner(
                "Generating your draft. This may take a little while..."
            ):
                try:
                    content = generate_paper(
                        grade,
                        subject,
                        chapters,
                        marks,
                        duration,
                        difficulty,
                        question_types,
                        include_answers,
                        instructions,
                    )

                    paper = {
                        "grade": grade,
                        "subject": subject,
                        "chapters": chapters,
                        "marks": marks,
                        "duration": duration,
                        "difficulty": difficulty,
                        "content": content,
                    }

                    st.session_state.current_paper = paper
                    st.session_state.papers.append(paper)
                    st.success("Draft generated! Review it carefully below.")

                except Exception as exc:
                    st.error(
                        f"Generation failed: {type(exc).__name__}: "
                        f"{str(exc)[:700]}"
                    )

    paper = st.session_state.current_paper

    if paper:
        st.divider()
        st.subheader("Question-paper preview")

        st.warning(
            "AI-generated draft: verify factual accuracy, mark totals, "
            "answer keys and syllabus alignment before using it."
        )

        st.markdown(
            f"**Class {paper['grade']} · {paper['subject']} · "
            f"{paper['marks']} marks · {paper['duration']} minutes**"
        )

        st.markdown(paper["content"])

        docx_data = create_docx(
            f"Class {paper['grade']} {paper['subject']} Question Paper",
            paper["content"],
            True,
        )

        st.download_button(
            "Download Word document (.docx)",
            data=docx_data,
            file_name="PaperPilot_Question_Paper.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            type="primary",
        )


# ---------------- MY PAPERS ----------------

elif page == "My Papers":
    st.title("My papers")
    st.write("Papers generated during this app session.")

    if not st.session_state.papers:
        st.info("Your generated papers will appear here.")
    else:
        for index, paper in enumerate(
            reversed(st.session_state.papers), start=1
        ):
            with st.expander(
                f"{index}. Class {paper['grade']} {paper['subject']} "
                f"· {paper['marks']} marks"
            ):
                st.markdown(paper["content"])


# ---------------- SETTINGS ----------------

elif page == "Settings":
    st.title("Settings")

    st.subheader("AI connection")

    if os.getenv("OPENROUTER_API_KEY"):
        st.success("An API key is configured locally.")
    else:
        st.error("No API key found. Check the .env file.")

    st.caption(
        "The API key is never displayed in this interface. "
        "Keep the .env file private."
    )

    st.subheader("About PaperPilot")
    st.write(
        "PaperPilot is a prototype that drafts educational assessments. "
        "It does not guarantee official CBSE approval or curriculum compliance."
    )

    st.info(
        "Generated content should be checked by a teacher before use."
    )