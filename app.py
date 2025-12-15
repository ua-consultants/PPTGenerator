import streamlit as st
from pptx import Presentation
from pptx.util import Inches
import os
import tempfile
import shutil
from io import BytesIO

# Initialize session state
if 'slide_groups' not in st.session_state:
    st.session_state.slide_groups = []
if 'selected_images' not in st.session_state:
    st.session_state.selected_images = []
if 'confirm_duplicate_add' not in st.session_state:
    st.session_state.confirm_duplicate_add = False
if 'pending_duplicate_images' not in st.session_state:
    st.session_state.pending_duplicate_images = []
if 'pending_new_images' not in st.session_state:
    st.session_state.pending_new_images = []

st.set_page_config(page_title="Furniture Inspection PPT Builder", layout="wide")
st.title("🪑 Furniture Inspection PPT Builder")
st.markdown("Upload inspection photos, select images for each slide, and generate a PowerPoint report.")

# Helper: Get all used images
def get_all_used_images():
    used = set()
    for group in st.session_state.slide_groups:
        used.update(group)
    return used

# === Step 1: Upload images ===
uploaded_files = st.file_uploader(
    "📁 Upload inspection photos (JPG/PNG)", 
    type=["jpg", "jpeg", "png"], 
    accept_multiple_files=True,
    help="Select multiple images. You can upload from a folder."
)

if uploaded_files:
    temp_dir = tempfile.mkdtemp()
    image_paths = []
    
    for f in uploaded_files:
        safe_name = f.name.replace(" ", "_").replace("-", "_")
        path = os.path.join(temp_dir, safe_name)
        with open(path, "wb") as fp:
            fp.write(f.getvalue())
        image_paths.append(path)

    # === Step 2: Display image grid ===
    st.subheader("🖼️ Select images for the current slide")
    cols = st.columns(4)
    selected_for_slide = []

    for i, img_path in enumerate(image_paths):
        col = cols[i % 4]
        with col:
            st.image(img_path, use_column_width=True)
            selected = st.checkbox("Select", key=f"select_{i}")
            if selected:
                selected_for_slide.append(img_path)

    st.session_state.selected_images = selected_for_slide

    # === Step 3: Buttons ===
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("✅ Add Selected to New Slide"):
            if not st.session_state.selected_images:
                st.warning("⚠️ No images selected!")
            else:
                # Check for duplicates
                used_images = get_all_used_images()
                duplicates = [img for img in st.session_state.selected_images if img in used_images]
                non_duplicates = [img for img in st.session_state.selected_images if img not in used_images]

                if duplicates:
                    # Store pending images and show confirmation
                    st.session_state.pending_duplicate_images = duplicates
                    st.session_state.pending_new_images = non_duplicates
                    st.session_state.confirm_duplicate_add = True
                    st.rerun()
                else:
                    # No duplicates → add directly
                    st.session_state.slide_groups.append(st.session_state.selected_images.copy())
                    st.session_state.selected_images = []
                    st.success(f"✅ Added {len(st.session_state.slide_groups[-1])} image(s) to a new slide!")

    with col2:
        if st.button("🗑️ Clear Current Selection"):
            st.session_state.selected_images = []
            st.rerun()

    with col3:
        if st.button("🔁 Reset All Slides"):
            st.session_state.slide_groups = []
            st.session_state.selected_images = []
            st.session_state.confirm_duplicate_add = False
            st.rerun()

    # === Duplicate Confirmation Dialog ===
    if st.session_state.confirm_duplicate_add:
        st.warning("⚠️ Duplicate Images Detected!")
        dup_count = len(st.session_state.pending_duplicate_images)
        st.write(f"The following {dup_count} image(s) are already used in previous slides:")
        
        # Show thumbnails of duplicates
        dup_cols = st.columns(min(dup_count, 4))
        for i, img in enumerate(st.session_state.pending_duplicate_images):
            dup_cols[i % 4].image(img, width=100)
        
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("✅ Yes, Add Duplicates"):
                # Combine non-dup + dup
                final_images = st.session_state.pending_new_images + st.session_state.pending_duplicate_images
                st.session_state.slide_groups.append(final_images)
                st.session_state.selected_images = []
                st.session_state.confirm_duplicate_add = False
                st.session_state.pending_duplicate_images = []
                st.session_state.pending_new_images = []
                st.success(f"✅ Added {len(final_images)} image(s) (including duplicates) to a new slide!")
                st.rerun()
        with col_b:
            if st.button("❌ No, Skip Duplicates"):
                if st.session_state.pending_new_images:
                    st.session_state.slide_groups.append(st.session_state.pending_new_images)
                    st.success(f"✅ Added {len(st.session_state.pending_new_images)} new image(s) only.")
                else:
                    st.info("ℹ️ No new images to add.")
                st.session_state.selected_images = []
                st.session_state.confirm_duplicate_add = False
                st.session_state.pending_duplicate_images = []
                st.session_state.pending_new_images = []
                st.rerun()

    # === Show current slides ===
    if st.session_state.slide_groups:
        st.subheader("📋 Current Slides")
        for idx, group in enumerate(st.session_state.slide_groups):
            with st.expander(f"Slide {idx + 1} ({len(group)} image(s))"):
                preview_cols = st.columns(min(len(group), 4))
                for i, img in enumerate(group):
                    preview_cols[i % 4].image(img, width=120)

    # === Generate PPT ===
    if st.button("📥 Generate & Download PowerPoint"):
        if not st.session_state.slide_groups:
            st.error("❌ No slides to generate!")
        else:
            prs = Presentation()
            blank_layout = None
            for layout in prs.slide_layouts:
                if "blank" in layout.name.lower():
                    blank_layout = layout
                    break
            if blank_layout is None:
                blank_layout = prs.slide_layouts[0]

            for group in st.session_state.slide_groups:
                slide = prs.slides.add_slide(blank_layout)
                num_imgs = len(group)
                
                if num_imgs <= 4:
                    rows = 2
                elif num_imgs <= 9:
                    rows = 3
                else:
                    rows = 4
                cols_grid = (num_imgs + rows - 1) // rows

                max_width = 9.0
                max_height = 6.5
                img_width = max_width / cols_grid
                img_height = max_height / rows

                for idx, img_path in enumerate(group):
                    row = idx // cols_grid
                    col = idx % cols_grid
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
                        st.warning(f"⚠️ Failed to add {os.path.basename(img_path)}")

            ppt_buffer = BytesIO()
            prs.save(ppt_buffer)
            ppt_buffer.seek(0)

            st.download_button(
                label="⬇️ Download Inspection Report (PPTX)",
                data=ppt_buffer,
                file_name="Furniture_Inspection_Report.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )
            shutil.rmtree(temp_dir, ignore_errors=True)

else:
    st.info("👆 Please upload images to get started.")