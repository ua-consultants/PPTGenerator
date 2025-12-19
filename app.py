# app.py — Slide-by-Slide PPT Generator
import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import os
import tempfile
from io import BytesIO

# Initialize session state
if 'slides' not in st.session_state:
    st.session_state.slides = []  # List of lists: [[img1, img2], [img3, img4], ...]

st.set_page_config(page_title="🪑 Inspection PPT Builder", layout="wide")
st.title("🪑 Furniture Inspection PPT Builder")
st.markdown("Add images **slide by slide**. When done, generate your report.")

# Show current slides
if st.session_state.slides:
    st.subheader(f"📋 Preview ({len(st.session_state.slides)} slide(s))")
    for i, slide_images in enumerate(st.session_state.slides):
        with st.expander(f"Slide {i+1} ({len(slide_images)} image(s))"):
            cols = st.columns(min(len(slide_images), 4))
            for j, img in enumerate(slide_images):
                cols[j % 4].image(img, width=120)

# Main action area
if not st.session_state.slides:
    st.info("👆 Click below to add your first slide.")
    action = "add_first"
else:
    action = st.radio(
        "What would you like to do next?",
        ["Add another slide", "Generate and download PPT"],
        horizontal=True
    )

# Handle actions
if not st.session_state.slides or action == "Add another slide":
    st.subheader("🖼️ Upload images for this slide")
    uploaded_files = st.file_uploader(
        "Drag & drop or browse images (JPG/PNG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True,
        key="slide_uploader"
    )
    
    if uploaded_files:
        st.success(f"✅ {len(uploaded_files)} image(s) selected for this slide.")
        # Show thumbnails
        cols = st.columns(min(len(uploaded_files), 4))
        for i, f in enumerate(uploaded_files):
            cols[i % 4].image(f, width=100)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Add This Slide"):
                # Save to temp dir and store paths
                temp_dir = tempfile.mkdtemp()
                image_paths = []
                for f in uploaded_files:
                    path = os.path.join(temp_dir, f.name)
                    with open(path, "wb") as fp:
                        fp.write(f.getvalue())
                    image_paths.append(path)
                st.session_state.slides.append(image_paths)
                st.rerun()
        with col2:
            st.button("↩️ Cancel")

# Generate PPT
if st.session_state.slides and (not st.session_state.slides or action == "Generate and download PPT"):
    if st.button("📥 Generate and Download PPT", type="primary"):
        if not st.session_state.slides:
            st.error("❌ No slides to generate!")
        else:
            # Create PPT
            prs = Presentation()
            # Find blank layout
            blank_layout = None
            for layout in prs.slide_layouts:
                if "blank" in layout.name.lower():
                    blank_layout = layout
                    break
            if blank_layout is None:
                blank_layout = prs.slide_layouts[0]

            for slide_images in st.session_state.slides:
                slide = prs.slides.add_slide(blank_layout)
                num_imgs = len(slide_images)
                
                # Auto grid
                if num_imgs <= 4:
                    rows, cols = (2, 2)
                elif num_imgs <= 6:
                    rows, cols = (2, 3)
                elif num_imgs <= 9:
                    rows, cols = (3, 3)
                else:
                    rows, cols = (4, 4)
                
                max_width = 9.0
                max_height = 6.5
                img_width = max_width / cols
                img_height = max_height / rows

                for idx, img_path in enumerate(slide_images):
                    row = idx // cols
                    col = idx % cols
                    left = Inches(col * img_width + 0.15)
                    top = Inches(row * img_height + 0.3)
                    try:
                        slide.shapes.add_picture(
                            img_path,
                            left, top,
                            width=Inches(img_width - 0.3),
                            height=Inches(img_height - 0.3)
                        )
                    except Exception as e:
                        st.warning(f"⚠️ Failed to add image")

            # Save and download
            ppt_buffer = BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)

            st.download_button(
                label="⬇️ Download Inspection Report (PPTX)",
                data=ppt_buffer,
                file_name="Furniture_Inspection_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

# Reset button (optional)
if st.session_state.slides:
    if st.button("🔄 Start Over"):
        st.session_state.slides = []
        st.rerun()