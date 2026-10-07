import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Inches, Mm, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import io
import re

st.set_page_config(page_title="Flashcard PPT Generator", page_icon="📚", layout="centered")

st.title("📚 Flashcard PowerPoint Generator")
st.markdown("Upload your CSV question bank to automatically generate a fresh PowerPoint flashcard deck.")

# Initialize session state
if "df" not in st.session_state:
    st.session_state.df = None


def clean_value(value):
    """Normalize values to a display-safe string."""
    if value is None:
        return ""
    if isinstance(value, float) and pd.isna(value):
        return ""
    return str(value)


def add_line_breaks_before_numbers(text):
    """Insert a double line break before numbered list items to improve readability."""
    if text is None:
        return ""
    text = str(text)
    return re.sub(r'([^\n])\s+(\d+\.)', r'\1\n\n\2', text)


def create_flashcard_grid(slide, slide_width, slide_height, topics, questions=None, answers=None, is_answer_sheet=False):
    """Create a 2-column by 4-row grid of flashcards on a blank slide."""
    rows = 4
    cols = 2
    margin = Mm(5)
    gap = Mm(3)

    usable_width = slide_width - (margin * 2) - (gap * (cols - 1))
    usable_height = slide_height - (margin * 2) - (gap * (rows - 1))
    card_width = usable_width / cols
    card_height = usable_height / rows

    for row in range(rows):
        for col in range(cols):
            slot_index = row * cols + col
            if is_answer_sheet:
                # Mirror horizontally: map current display position to original card position
                source_index = row * cols + (cols - 1 - col)
                if source_index >= len(topics):
                    continue
                title_value = topics[source_index]
                content_value = answers[source_index] if answers is not None and source_index < len(answers) else ""
            else:
                if slot_index >= len(topics):
                    continue
                title_value = topics[slot_index]
                content_value = questions[slot_index] if questions is not None and slot_index < len(questions) else ""

            left = margin + col * (card_width + gap)
            top = margin + row * (card_height + gap)
            create_single_flashcard(
                slide=slide,
                left=left,
                top=top,
                width=card_width,
                height=card_height,
                title=clean_value(title_value),
                content=clean_value(content_value),
                is_answer_sheet=is_answer_sheet,
            )


def create_single_flashcard(slide, left, top, width, height, title, content, is_answer_sheet=False):
    """Draw one flashcard with a black title bar and a content box."""
    if is_answer_sheet:
        content_left = left
        content_top = top
        content_width = width
        content_height = height
    else:
        outer_card = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
        outer_card.fill.solid()
        outer_card.fill.fore_color.rgb = RGBColor(255, 255, 255)
        outer_card.line.color.rgb = RGBColor(0, 0, 0)
        outer_card.line.width = Pt(0.5)

        title_area_height = height * 0.24
        title_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, title_area_height)
        title_box.fill.solid()
        title_box.fill.fore_color.rgb = RGBColor(0, 0, 0)
        title_box.line.color.rgb = RGBColor(0, 0, 0)
        title_box.line.width = Pt(0.5)

        title_tf = title_box.text_frame
        title_tf.clear()
        title_tf.word_wrap = True
        title_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        title_tf.margin_left = Mm(3)
        title_tf.margin_right = Mm(3)
        title_tf.margin_top = Mm(1.5)
        title_tf.margin_bottom = Mm(1.5)
        p = title_tf.paragraphs[0]
        p.text = title
        p.alignment = PP_ALIGN.LEFT
        p.space_before = Pt(0)
        p.space_after = Pt(0)
        for run in p.runs:
            run.font.size = Pt(16)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

        content_left = left + Mm(2.5)
        content_top = top + title_area_height + Mm(1.2)
        content_width = width - Mm(5)
        content_height = height - title_area_height - Mm(2.4)

    content_box = slide.shapes.add_textbox(content_left, content_top, content_width, content_height)
    content_box.fill.background()
    content_box.line.color.rgb = RGBColor(255, 255, 255)
    content_box.line.width = Pt(0)

    body = add_line_breaks_before_numbers(content)
    body_tf = content_box.text_frame
    body_tf.clear()
    body_tf.word_wrap = True
    body_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    body_tf.margin_left = Mm(0)
    body_tf.margin_right = Mm(0)
    body_tf.margin_top = Mm(0)
    body_tf.margin_bottom = Mm(0)

    if body.strip():
        lines = body.splitlines()
        for idx, line in enumerate(lines):
            p = body_tf.paragraphs[0] if idx == 0 else body_tf.add_paragraph()
            p.text = line
            p.alignment = PP_ALIGN.LEFT
            p.space_before = Pt(0)
            p.space_after = Pt(0)
            for run in p.runs:
                run.font.size = Pt(12)
                run.font.bold = False
                run.font.color.rgb = RGBColor(0, 0, 0)
    else:
        p = body_tf.paragraphs[0]
        p.text = ""


def create_flashcard_presentation(topics, questions, answers, show_answers=True):
    """Create a fresh portrait A4 deck from scratch."""
    prs = Presentation()
    prs.slide_width = Inches(8.27)
    prs.slide_height = Inches(11.69)
    blank_layout = prs.slide_layouts[6]

    num_cards = len(topics)
    cards_per_slide = 8
    num_pairs = (num_cards + cards_per_slide - 1) // cards_per_slide

    for pair_index in range(num_pairs):
        start = pair_index * cards_per_slide
        end = min(start + cards_per_slide, num_cards)

        slide_topics = topics[start:end]
        slide_questions = questions[start:end]
        slide_answers = answers[start:end]

        question_slide = prs.slides.add_slide(blank_layout)
        create_flashcard_grid(
            slide=question_slide,
            slide_width=prs.slide_width,
            slide_height=prs.slide_height,
            topics=slide_topics,
            questions=slide_questions,
            answers=None,
            is_answer_sheet=False,
        )

        answer_slide = prs.slides.add_slide(blank_layout)
        answer_values = slide_answers if show_answers else ["" for _ in slide_answers]
        create_flashcard_grid(
            slide=answer_slide,
            slide_width=prs.slide_width,
            slide_height=prs.slide_height,
            topics=slide_topics,
            questions=None,
            answers=answer_values,
            is_answer_sheet=True,
        )

    return prs


def parse_csv_flexible(csv_text):
    """Parse pasted Topic,Question,Answer rows without breaking when answers contain commas."""
    import csv

    csv_text = csv_text.strip()
    if csv_text.startswith("```csv"):
        csv_text = csv_text[6:]
    if csv_text.startswith("```"):
        csv_text = csv_text[3:]
    if csv_text.endswith("```"):
        csv_text = csv_text[:-3]

    lines = [line.strip() for line in csv_text.strip().splitlines() if line.strip()]
    if not lines:
        raise ValueError("No data was pasted.")
    if lines[0].lower().startswith("topic,") or lines[0].startswith('"Topic"'):
        lines = lines[1:]

    rows = []
    reader = csv.reader(lines)

    for idx, row in enumerate(reader, start=1):
        if len(row) < 3:
            raise ValueError(
                f"Row {idx} is not in a valid Topic,Question,Answer format. "
                f"Found {len(row)} parts: {row}"
            )
        topic = row[0].strip().strip('"')
        question = row[1].strip().strip('"')
        answer = row[2].strip().strip('"')
        rows.append({"Topic": topic, "Question": question, "Answer": answer})

    return pd.DataFrame(rows)


# Sidebar controls
st.sidebar.header("Configuration")
show_answers = st.sidebar.checkbox("Include Answers on Answer Slides", value=True, help="Uncheck to generate blank answer slides.")

# Sidebar input method selection
st.sidebar.header("Data Input")
input_method = st.sidebar.radio("How would you like to input your data?", ["Upload CSV File", "Paste CSV Data"])

if input_method == "Upload CSV File":
    uploaded_csv = st.sidebar.file_uploader("Upload CSV Question Bank", type=["csv"])
    if uploaded_csv is not None:
        try:
            st.session_state.df = pd.read_csv(uploaded_csv)
        except Exception as e:
            st.error(f"Error reading CSV file: {e}")
else:
    with st.sidebar.form("csv_input_form"):
        csv_text = st.text_area(
            "Paste your data here in this format: Topic,Question,Answer",
            height=200,
            placeholder="Topic,Question,Answer\nDensity,1. What is density?,1. Density = mass / volume, ρ = m / V [cite: 1]"
        )
        submitted = st.form_submit_button("📤 Parse CSV Data")

    if submitted:
        if csv_text.strip():
            try:
                st.session_state.df = parse_csv_flexible(csv_text)
            except Exception as e:
                st.error(f"Error parsing CSV data: {e}")
        else:
            st.warning("Please paste CSV data before parsing.")


if st.session_state.df is not None:
    try:
        st.success(f"Successfully loaded CSV with {len(st.session_state.df)} rows.")
        st.subheader("CSV Data Preview")
        st.dataframe(st.session_state.df.head())

        if len(st.session_state.df.columns) < 3:
            st.error("The uploaded CSV must have at least 3 columns: Topic, Question, and Answer.")
        else:
            t_col = st.session_state.df.columns[0]
            q_col = st.session_state.df.columns[1]
            a_col = st.session_state.df.columns[2]

            topics = st.session_state.df[t_col].astype(str).tolist()
            questions = st.session_state.df[q_col].astype(str).tolist()
            answers = st.session_state.df[a_col].astype(str).tolist()

            if st.button("Generate Flashcard Presentation"):
                try:
                    prs = create_flashcard_presentation(topics, questions, answers, show_answers=show_answers)

                    output_buffer = io.BytesIO()
                    prs.save(output_buffer)
                    output_buffer.seek(0)

                    num_cards = len(questions)
                    num_pairs = (num_cards + 7) // 8
                    st.success(f"Generated presentation with {len(prs.slides)} slides ({num_pairs} question/answer page sets)!")

                    st.download_button(
                        label="📥 Download Generated Flashcards (.pptx)",
                        data=output_buffer,
                        file_name="generated_flashcards.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    )
                except Exception as e:
                    st.error(f"Error generating PowerPoint: {e}")
                    import traceback
                    st.error(traceback.format_exc())

    except Exception as e:
        st.error(f"Error processing CSV data: {e}")
else:
    st.info("Please upload a CSV file or paste CSV data via the sidebar to get started.")
