import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import copy
import io
from io import StringIO

st.set_page_config(page_title="Flashcard PPT Generator", page_icon="📚", layout="centered")

st.title("📚 Flashcard PowerPoint Generator")
st.markdown("Upload your CSV question bank to automatically populate and scale your PowerPoint flashcard template.")

# Helper function to duplicate slides
def clone_slide(prs, source_slide):
    blank_layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(blank_layout)
    for shape in list(new_shape.shapes):
        sp = shape.element
        sp.getparent().remove(sp)
    for shape in source_slide.shapes:
        new_el = copy.deepcopy(shape.element)
        new_slide.shapes._spTree.append(new_el)
    return new_slide

def add_line_breaks_before_numbers(text):
    """Add double line break before numbered items (e.g., '1. ', '2. ', etc.)"""
    import re
    # Replace pattern: if there's text before a number, add double line breaks
    return re.sub(r'([^\n])\s+(\d+\.)', r'\1\n\n\2', text)

def format_question_cell_with_title(cell, title, content):
    """Format question cell with centered, bold, white text on black highlight title, 
    then double line break before content."""
    cell.text = ""
    tf = cell.text_frame
    tf.clear()
    tf.word_wrap = True

    # Title paragraph: bold, centered, white text on black background
    title_paragraph = tf.paragraphs[0]
    title_paragraph.alignment = PP_ALIGN.CENTER
    title_run = title_paragraph.add_run()
    title_run.text = str(title)
    title_run.font.bold = True
    title_run.font.size = Pt(11)
    title_run.font.color.rgb = RGBColor(255, 255, 255)  # White text
    
    # Apply black background to the title run using XML
    try:
        rPr = title_run._r.get_or_add_rPr()
        shd_elm = rPr.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd')
        if shd_elm is not None:
            rPr.remove(shd_elm)
        
        # Create highlight element
        from lxml import etree
        shd = etree.SubElement(
            rPr,
            '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd',
            fill='000000'
        )
    except Exception:
        # If XML manipulation fails, just skip the background
        pass

    # Add a blank paragraph to create the double line break
    tf.add_paragraph()

    # Content paragraph
    content_paragraph = tf.add_paragraph()
    content_run = content_paragraph.add_run()
    content_run.text = add_line_breaks_before_numbers(str(content))


def format_answer_cell(cell, content):
    """Format answer cell with content only (no title)."""
    cell.text = ""
    tf = cell.text_frame
    tf.clear()
    tf.word_wrap = True

    # Content paragraph
    content_paragraph = tf.paragraphs[0]
    content_run = content_paragraph.add_run()
    content_run.text = add_line_breaks_before_numbers(str(content))


def get_cell_indices_left_to_right(num_rows, num_cols, num_items):
    """Generate cell indices filling left-to-right, then down"""
    indices = []
    for row in range(num_rows):
        for col in range(num_cols):  # Left to right
            cell_index = row * num_cols + col
            if cell_index < num_items:
                indices.append((row, col, cell_index))
    return indices

def get_cell_indices_mirror_horizontal(num_rows, num_cols, num_items):
    """Generate cell indices mirrored horizontally for double-sided printing.
    Flips left-right so answers align with questions on the back when printed.
    For a 2x4 grid, position [0,0] maps to [0,3], [0,1] maps to [0,2], etc."""
    indices = []
    for row in range(num_rows):
        for col in range(num_cols):  # Left to right
            # Mirror column position: if col=0, mirror_col=3; if col=1, mirror_col=2, etc.
            mirror_col = num_cols - 1 - col
            cell_index = row * num_cols + col
            if cell_index < num_items:
                indices.append((row, mirror_col, cell_index))
    return indices

def parse_csv_flexible(csv_text):
    """
    Parse pasted flashcard data in the form:
    Topic,Question,Answer

    This supports the common case where the answer text contains commas, and standard
    CSV quoting is not used. It splits only on the first two commas, so the third field
    retains the rest of the answer text.
    """
    lines = [line.strip() for line in csv_text.splitlines() if line.strip()]
    if not lines:
        raise ValueError("No data was pasted.")

    # Remove a possible header row such as: Topic,Front (Question),Back (Answer)
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
    # Paste CSV data
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
        
        # Display preview
        st.subheader("CSV Data Preview")
        st.dataframe(df.head())
        
        # Validate columns
        if len(df.columns) < 3:
            st.error("The uploaded CSV must have at least 3 columns: Topic, Question (2nd column), and Answer (3rd column).")
        else:
            # Extract topic (col 1 / index 0), questions (col 2 / index 1), answers (col 3 / index 2)
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
                    
                    # Ensure enough slide pairs exist (each pair: 1 Question slide + 1 Answer slide)
                    while len(prs.slides) < num_pairs * 2:
                        clone_slide(prs, prs.slides[0])
                        clone_slide(prs, prs.slides[1])
                        
                    for pair_idx in range(num_pairs):
                        q_slide = prs.slides[pair_idx * 2]
                        a_slide = prs.slides[pair_idx * 2 + 1]
                        
                        batch_t = topics[pair_idx * 8 : (pair_idx + 1) * 8]
                        batch_q = questions[pair_idx * 8 : (pair_idx + 1) * 8]
                        batch_a = answers[pair_idx * 8 : (pair_idx + 1) * 8]
                        
                        # Fill Question Slide Table placeholders (left to right, top to bottom)
                        for shape in q_slide.shapes:
                            if shape.has_table:
                                table = shape.table
                                num_rows = len(table.rows)
                                num_cols = len(table.rows[0].cells)
                                
                                for r_idx, row in enumerate(table.rows):
                                    for c_idx, cell in enumerate(row.cells):
                                        idx = r_idx * num_cols + c_idx
                                        if idx < len(batch_q):
                                            format_question_cell_with_title(cell, batch_t[idx], batch_q[idx])
                                        else:
                                            cell.text = ""
                                            
                        # Fill Answer Slide Table placeholders (horizontally mirrored for double-sided printing)
                        for shape in a_slide.shapes:
                            if shape.has_table:
                                table = shape.table
                                num_rows = len(table.rows)
                                num_cols = len(table.rows[0].cells)
                                
                                # Get indices for horizontally mirrored filling
                                cell_mapping = get_cell_indices_mirror_horizontal(num_rows, num_cols, len(batch_a))
                                
                                # First, clear all cells
                                for row in table.rows:
                                    for cell in row.cells:
                                        cell.text = ""
                                
                                # Then fill according to mirror mapping
                                for r_idx, c_idx, item_idx in cell_mapping:
                                    if item_idx < len(batch_a):
                                        answer_content = (
                                            add_line_breaks_before_numbers(batch_a[item_idx])
                                            if show_answers else ""
                                        )
                                        # Answer side: no title, just content
                                        format_answer_cell(table.rows[r_idx].cells[c_idx], answer_content)
                                            
                    # Save to BytesIO buffer
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
