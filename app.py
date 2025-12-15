import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import os
import tempfile
import shutil
from io import BytesIO

# Initialize session state
if 'slide_groups' not in st.session_state:
    st.session_state.slide_groups = []  # List of lists: [[img1, img2], [img3], ...]
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = []  # Currently selected for current slide

st.set_page_config(page_title="Furniture Inspection PPT Builder", layout="wide")
st.title("🪑 Furniture Inspection PPT Builder")
st.markdown("Upload inspection photos, select images for each slide, and generate a PowerPoint report.")

# === Step 1: Upload images ===
uploaded_files = st.file_uploader(
    "📁 Upload inspection photos (JPG/PNG)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True,
    help="Select multiple images. You can upload from a folder."
)

if uploaded_files:
    # Create a temporary directory to store uploaded images
    temp_dir = tempfile.mkdtemp()
    image_paths = []
    
    for f in uploaded_files:
        # Avoid duplicate filenames
        safe_name = f.name.replace(" ", "_").replace("-", "_")
        path = os.path.join(temp_dir, safe_name)
        with open(path, "wb") as fp:
            fp.write(f.getvalue())
        image_paths.append(path)

    # === Step 2: Display image grid with checkboxes ===
    st.subheader("🖼️ Select images for the current slide")
    cols = st.columns(4)  # 4 images per row
    selected_for_slide = []

    for i, img_path in enumerate(image_paths):
        col = cols[i % 4]
        with col:
            st.image(img_path, use_column_width=True)
            selected = st.checkbox("Select", key=f"select_{i}")
            if selected:
                selected_for_slide.append(img_path)

    st.session_state.selected_images = selected_for_slide

    # === Step 3: Slide management buttons ===
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("✅ Add Selected to New Slide"):
            if st.session_state.selected_images:
                st.session_state.slide_groups.append(st.session_state.selected_images.copy())
                st.session_state.selected_images = []
                st.success(f"✅ Added {len(st.session_state.slide_groups[-1])} image(s) to a new slide!")
            else:
                st.warning("⚠️ No images selected!")
    
    with col2:
        if st.button("🗑️ Clear Current Selection"):
            st.session_state.selected_images = []
            st.experimental_rerun()  # Optional: refresh UI

    with col3:
        if st.button("🔁 Reset All Slides"):
            st.session_state.slide_groups = []
            st.session_state.selected_images = []
            st.experimental_rerun()

    # === Step 4: Show current slide groups ===
    if st.session_state.slide_groups:
        st.subheader("📋 Current Slides")
        for idx, group in enumerate(st.session_state.slide_groups):
            with st.expander(f"Slide {idx + 1} ({len(group)} image(s))"):
                preview_cols = st.columns(min(len(group), 4))
                for i, img in enumerate(group):
                    preview_cols[i % 4].image(img, width=120)

    # === Step 5: Generate PowerPoint ===
    if st.button("📥 Generate & Download PowerPoint"):
        if not st.session_state.slide_groups:
            st.error("❌ No slides to generate! Add at least one slide first.")
        else:
            # Create PowerPoint
            prs = Presentation()
            
            # Try to find a blank layout
            blank_layout = None
            for layout in prs.slide_layouts:
                if "blank" in layout.name.lower():
                    blank_layout = layout
                    break
            if blank_layout is None:
                blank_layout = prs.slide_layouts[0]  # fallback

            for group in st.session_state.slide_groups:
                slide = prs.slides.add_slide(blank_layout)
                num_imgs = len(group)
                
                # Determine grid: 1–4 → 2 rows max, 5–12 → 3 rows
                if num_imgs <= 4:
                    rows = 2
                elif num_imgs <= 9:
                    rows = 3
                else:
                    rows = 4
                cols_grid = (num_imgs + rows - 1) // rows  # ceil division

                max_width = 9.0   # inches (standard slide width ~10")
                max_height = 6.5  # inches (leave margin)

                img_width = max_width / cols_grid
                img_height = max_height / rows

                for idx, img_path in enumerate(group):
                    row = idx // cols_grid
                    col = idx % cols_grid
                    left = Inches(col * img_width + 0.15)
                    top = Inches(row * img_height + 0.3)
                    # Add picture (handle potential size errors)
                    try:
                        slide.shapes.add_picture(
                            img_path,
                            left, top,
                            width=Inches(img_width - 0.3),
                            height=Inches(img_height - 0.3)
                        )
                    except Exception as e:
                        st.warning(f"⚠️ Could not add image {os.path.basename(img_path)}: {str(e)}")

            # Save to BytesIO
            ppt_buffer = BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)

            st.download_button(
                label="⬇️ Download Inspection Report (PPTX)",
                data=ppt_buffer,
                file_name="Furniture_Inspection_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

            # Optional: cleanup (comment out if you want to debug)
            shutil.rmtree(temp_dir, ignore_errors=True)

else:
    st.info("👆 Please upload images to get started.")