import streamlit as st
import pandas as pd
from pptx import Presentation
import copy
import io

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

# Sidebar controls
st.sidebar.header("Configuration")
show_answers = st.sidebar.checkbox("Include Answers on Answer Slides", value=True, help="Uncheck to generate blank answer slides.")

# File uploader for CSV
uploaded_csv = st.sidebar.file_uploader("Upload CSV Question Bank", type=["csv"])

# Default template path
TEMPLATE_PATH = "Flashcard template.pptx"

if uploaded_csv is not None:
    try:
        df = pd.read_csv(uploaded_csv)
        st.success(f"Successfully loaded CSV with {len(df)} rows.")
        
        # Display preview
        st.subheader("CSV Data Preview")
        st.dataframe(df.head())
        
        # Validate columns
        if len(df.columns) < 3:
            st.error("The uploaded CSV must have at least 3 columns: Topic, Question (2nd column), and Answer (3rd column).")
        else:
            # Extract questions (col 2 / index 1) and answers (col 3 / index 2)
            q_col = df.columns[1]
            a_col = df.columns[2]
            
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
                        
                        batch_q = questions[pair_idx * 8 : (pair_idx + 1) * 8]
                        batch_a = answers[pair_idx * 8 : (pair_idx + 1) * 8]
                        
                        # Fill Question Slide Table placeholders
                        for shape in q_slide.shapes:
                            if shape.has_table:
                                table = shape.table
                                for r_idx, row in enumerate(table.rows):
                                    for c_idx, cell in enumerate(row.cells):
                                        idx = r_idx * 2 + c_idx
                                        if idx < len(batch_q):
                                            cell.text = batch_q[idx]
                                        else:
                                            cell.text = ""
                                            
                        # Fill Answer Slide Table placeholders
                        for shape in a_slide.shapes:
                            if shape.has_table:
                                table = shape.table
                                for r_idx, row in enumerate(table.rows):
                                    for c_idx, cell in enumerate(row.cells):
                                        idx = r_idx * 2 + c_idx
                                        if idx < len(batch_a):
                                            cell.text = batch_a[idx] if show_answers else ""
                                        else:
                                            cell.text = ""
                                            
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
        st.error(f"Error reading CSV file: {e}")
else:
    st.info("Please upload your CSV file via the sidebar to get started.")
