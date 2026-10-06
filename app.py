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
    # Replace pattern: if there's text before a number, add double line breaks
    return re.sub(r'([^\n])\s+(\d+\.)', r'\1\n\n\2', text)

def replace_text_in_cell(cell, old_text, new_text):
    """Replace placeholder text in a cell while preserving formatting."""
    if cell.text_frame is None:
        return
    
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)

def get_cell_indices_left_to_right(num_rows, num_cols, num_items):
    """Generate cell indices filling left-to-right, then down"""
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
    """Parse pasted Topic,Question,Answer rows without breaking when answers contain commas."""
    lines = [line.strip() for line in csv_text.splitlines() if line.strip()]
    if not lines:
        raise ValueError("No data was pasted.")

    if lines[0].lower().startswith("topic,"):
        lines = lines[1:]

    rows = []
    for idx, line in enumerate(lines, start=1):
        parts = [p.strip() for p in line.split(",", 2)]
        if len(parts) != 3:
            raise ValueError(
                f"Row {idx} is not in a valid Topic,Question,Answer format. "
                f"Found {len(parts)} parts: {parts}"
            )
        topic, question, answer = parts
        rows.append({"Topic": topic, "Question": question, "Answer": answer})

    return pd.DataFrame(rows)


# Sidebar controls
st.sidebar.header("Configuration")
show_answers = st.sidebar.checkbox("Include Answers on Answer Slides", value=True, help="Uncheck to generate blank answer slides.")

# Sidebar input method selection
st.sidebar.header("Data Input")
input_method = st.sidebar.radio("How would you like to input your data?", ["Upload CSV File", "Paste CSV Data"])

df = None

if input_method == "Upload CSV File":
    uploaded_csv = st.sidebar.file_uploader("Upload CSV Question Bank", type=["csv"])
    if uploaded_csv is not None:
        try:
            df = pd.read_csv(uploaded_csv)
        except Exception as e:
            st.error(f"Error reading CSV file: {e}")
else:
    csv_text = st.sidebar.text_area(
        "Paste your data here in this format: Topic,Question,Answer",
        height=200,
        placeholder="Topic,Question,Answer\nDensity,1. What is density?,1. Density = mass / volume, ρ = m / V [cite: 1]"
    )
    if csv_text.strip():
        try:
            df = parse_csv_flexible(csv_text)
        except Exception as e:
            st.error(f"Error parsing CSV data: {e}")

# Default template path
TEMPLATE_PATH = "Flashcard template.pptx"

if df is not None:
    try:
        st.success(f"Successfully loaded CSV with {len(df)} rows.")

        st.subheader("CSV Data Preview")
        st.dataframe(df.head())

        if len(df.columns) < 3:
            st.error("The uploaded CSV must have at least 3 columns: Topic, Question (2nd column), and Answer (3rd column).")
        else:
            t_col = df.columns[0]
            q_col = df.columns[1]
            a_col = df.columns[2]

            topics = df[t_col].astype(str).tolist()
            questions = df[q_col].astype(str).tolist()
            answers = df[a_col].astype(str).tolist()

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

                        # Process Question Slide
                        for shape in q_slide.shapes:
                            if shape.has_table:
                                table = shape.table
                                num_rows = len(table.rows)
                                num_cols = len(table.rows[0].cells)

                                # Replace placeholders in question slide
                                for idx, (title, question) in enumerate(zip(batch_t, batch_q)):
                                    placeholder_title = f"{{Title{idx+1}}}"
                                    placeholder_question = f"{{Question Bank{idx+1}}}"
                                    
                                    # Find and replace in all cells
                                    for row in table.rows:
                                        for cell in row.cells:
                                            replace_text_in_cell(cell, placeholder_title, title)
                                            replace_text_in_cell(cell, placeholder_question, add_line_breaks_before_numbers(question))

                        # Process Answer Slide
                        for shape in a_slide.shapes:
                            if shape.has_table:
                                table = shape.table
                                num_rows = len(table.rows)
                                num_cols = len(table.rows[0].cells)

                                # Get indices for horizontally mirrored filling
                                cell_mapping = get_cell_indices_mirror_horizontal(num_rows, num_cols, len(batch_a))

                                # Replace answer placeholders (already mirrored in template)
                                for idx in range(len(batch_a)):
                                    placeholder_answer = f"{{Answer Bank{idx+1}}}"
                                    answer_content = batch_a[idx] if show_answers else ""
                                    
                                    for row in table.rows:
                                        for cell in row.cells:
                                            replace_text_in_cell(cell, placeholder_answer, add_line_breaks_before_numbers(answer_content))

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
