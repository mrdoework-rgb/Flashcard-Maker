import streamlit as st
import pandas as pd
from pptx import Presentation
import copy
import io
import re
from io import StringIO

st.set_page_config(page_title="Flashcard PPT Generator", page_icon="📚", layout="centered")

st.title("📚 Flashcard PowerPoint Generator")
st.markdown("Upload your CSV question bank to automatically populate and scale your PowerPoint flashcard template.")

# Initialize session state
if "df" not in st.session_state:
    st.session_state.df = None

# Helper function to duplicate slides
def clone_slide(prs, source_slide):
    blank_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(blank_layout)
    for shape in list(new_slide.shapes):
        sp = shape.element
        sp.getparent().remove(sp)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide


def add_line_breaks_before_numbers(text):
    """Add double line break before numbered items (e.g., '1. ', '2. ', etc.)"""
    return re.sub(r'([^\n])\s+(\d+\.)', r'\1\n\n\2', text)


def replace_tokens_in_text_frame(text_frame, replacements):
    """Replace placeholder tokens in a text frame without losing the formatting of the slide template."""
    if text_frame is None:
        return

    # Update runs first (best preserves formatting)
    for paragraph in text_frame.paragraphs:
        runs = list(paragraph.runs)
        if not runs:
            current = paragraph.text
            updated = current
            for token, value in replacements.items():
                updated = updated.replace(token, value)
            if updated != current:
                paragraph.text = updated
            continue

        for run in runs:
            current = run.text
            updated = current
            for token, value in replacements.items():
                updated = updated.replace(token, value)
            if updated != current:
                run.text = updated


def replace_tokens_in_shape(shape, replacements):
    """Replace placeholder tokens in a shape or table cell."""
    if shape.has_table:
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                replace_tokens_in_text_frame(cell.text_frame, replacements)
    elif shape.has_text_frame:
        replace_tokens_in_text_frame(shape.text_frame, replacements)


def replace_tokens_in_slide(slide, replacements):
    """Replace placeholder tokens across all text-bearing shapes on a slide."""
    for shape in slide.shapes:
        replace_tokens_in_shape(shape, replacements)


def get_cell_indices_left_to_right(num_rows, num_cols, num_items):
    """Generate cell indices filling left-to-right, then down."""
    indices = []
    for row in range(num_rows):
        for col in range(num_cols):
            cell_index = row * num_cols + col
            if cell_index < num_items:
                indices.append((row, col, cell_index))
    return indices


def get_cell_indices_mirror_horizontal(num_rows, num_cols, num_items):
    """Generate cell indices mirrored horizontally for double-sided printing."""
    indices = []
    for row in range(num_rows):
        for col in range(num_cols):
            mirror_col = num_cols - 1 - col
            cell_index = row * num_cols + col
            if cell_index < num_items:
                indices.append((row, mirror_col, cell_index))
    return indices


def parse_csv_flexible(csv_text):
    """Parse pasted Topic,Question,Answer rows without breaking when answers contain commas.
    Handles both quoted and unquoted CSV formats."""
    import csv
    
    # Strip ```csv tags if present
    csv_text = csv_text.strip()
    if csv_text.startswith("```csv"):
        csv_text = csv_text[6:]  # Remove ```csv
    if csv_text.startswith("```"):
        csv_text = csv_text[3:]  # Remove ``` in case just ``` is there
    if csv_text.endswith("```"):
        csv_text = csv_text[:-3]  # Remove closing ```
    
    csv_text = csv_text.strip()
    
    lines = [line.strip() for line in csv_text.splitlines() if line.strip()]
    if not lines:
        raise ValueError("No data was pasted.")

    # Check if header is present and skip it
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
        
        # Take first 3 columns and strip quotes/whitespace
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

# Default template path
TEMPLATE_PATH = "Flashcard template.pptx"

if st.session_state.df is not None:
    try:
        st.success(f"Successfully loaded CSV with {len(st.session_state.df)} rows.")

        st.subheader("CSV Data Preview")
        st.dataframe(st.session_state.df.head())

        if len(st.session_state.df.columns) < 3:
            st.error("The uploaded CSV must have at least 3 columns: Topic, Question (2nd column), and Answer (3rd column).")
        else:
            t_col = st.session_state.df.columns[0]
            q_col = st.session_state.df.columns[1]
            a_col = st.session_state.df.columns[2]

            topics = st.session_state.df[t_col].astype(str).tolist()
            questions = st.session_state.df[q_col].astype(str).tolist()
            answers = st.session_state.df[a_col].astype(str).tolist()

            if st.button("Generate Flashcard Presentation"):
                try:
                    prs = Presentation(TEMPLATE_PATH)

                    num_questions = len(questions)
                    num_pairs = (num_questions + 7) // 8

                    while len(prs.slides) < num_pairs * 2:
                        clone_slide(prs, prs.slides[0])
                        clone_slide(prs, prs.slides[1])

                    for pair_idx in range(num_pairs):
                        q_slide = prs.slides[pair_idx * 2]
                        a_slide = prs.slides[pair_idx * 2 + 1]

                        batch_t = topics[pair_idx * 8 : (pair_idx + 1) * 8]
                        batch_q = questions[pair_idx * 8 : (pair_idx + 1) * 8]
                        batch_a = answers[pair_idx * 8 : (pair_idx + 1) * 8]

                        # Replace question placeholders: {Title1} {Question Bank1}, {Title2} {Question Bank2}, etc.
                        for card_idx, (title, question) in enumerate(zip(batch_t, batch_q), start=1):
                            replacements = {
                                f"{{Title{card_idx}}}": title,
                                f"{{Question Bank{card_idx}}}": add_line_breaks_before_numbers(question),
                            }
                            replace_tokens_in_slide(q_slide, replacements)

                        # Replace answer placeholders: {Answer Bank1}, {Answer Bank2}, etc.
                        for card_idx, answer in enumerate(batch_a, start=1):
                            answer_value = add_line_breaks_before_numbers(answer) if show_answers else ""
                            replacements = {
                                f"{{Answer Bank{card_idx}}}": answer_value,
                            }
                            replace_tokens_in_slide(a_slide, replacements)

                    output_buffer = io.BytesIO()
                    prs.save(output_buffer)
                    output_buffer.seek(0)

                    st.success(f"Generated presentation with {len(prs.slides)} slides ({num_pairs} question/answer page sets)!")

                    st.download_button(
                        label="📥 Download Generated Flashcards (.pptx)",
                        data=output_buffer,
                        file_name="generated_flashcards.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )

                except Exception as e:
                    st.error(f"Error generating PowerPoint: {e}")

    except Exception as e:
        st.error(f"Error processing CSV data: {e}")
else:
    st.info("Please upload a CSV file or paste CSV data via the sidebar to get started.")
