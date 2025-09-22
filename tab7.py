import io
import os
import zipfile
import streamlit as st
from PIL import Image as PILImage

from helper import extract_zip_to_temp

ALLOWED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
ALLOWED_TEXT_EXTENSIONS = {'.txt', '.md', '.csv'}
ALLOWED_VIDEO_EXTENSIONS = {'.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.webm', '.m4v', '.3gp', '.ogv'}
ALLOWED_AUDIO_EXTENSIONS = {'.mp3', '.wav', '.aac', '.flac', '.ogg', '.m4a', '.wma'}
ALL_ALLOWED_EXTENSIONS = ALLOWED_IMAGE_EXTENSIONS.union(ALLOWED_TEXT_EXTENSIONS).union(ALLOWED_VIDEO_EXTENSIONS).union(ALLOWED_AUDIO_EXTENSIONS)


def center_on_canvas_tool():
    st.header("Center on Transparent Canvas")
    st.markdown("Center small images on a transparent 500x500 canvas.")
    
    # File upload - add unique key
    uploaded_file = st.file_uploader(
        "Choose a ZIP file with images",
        type=['zip'],
        help="Upload a zip file containing images to process.",
        key="canvas_center_upload"  # Added unique key
    )
    
    if uploaded_file:
        # Options
        canvas_size = st.number_input(
            "Canvas size (pixels):",
            min_value=100,
            max_value=5000,
            value=500,
            step=10,
            key="canvas_size_input"  # Added unique key
        )
        
        # Add unique key to the button
        if st.button("Process Images", type="primary", key="process_canvas_images"):
            try:
                # Extract zip file
                with st.spinner("Extracting zip file..."):
                    temp_dir = extract_zip_to_temp(uploaded_file)
                    input_folder = temp_dir
                    items = os.listdir(temp_dir)
                    if len(items) == 1 and os.path.isdir(os.path.join(temp_dir, items[0])):
                        input_folder = os.path.join(temp_dir, items[0])
                
                # Create output folder
                output_folder = os.path.join(temp_dir, "canvas_output")
                os.makedirs(output_folder, exist_ok=True)
                
                # Process images
                processed_count = 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                all_files = []
                for root, dirs, files in os.walk(input_folder):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        if ext in ALLOWED_IMAGE_EXTENSIONS:
                            all_files.append(os.path.join(root, file))
                
                for i, file_path in enumerate(all_files):
                    status_text.text(f"Processing {os.path.basename(file_path)}... ({i+1}/{len(all_files)})")
                    progress_bar.progress((i + 1) / len(all_files))
                    
                    try:
                        with PILImage.open(file_path) as img:
                            width, height = img.size
                            
                            if width < canvas_size and height < canvas_size:
                                # Create transparent canvas
                                canvas = PILImage.new('RGBA', (canvas_size, canvas_size), (255, 255, 255, 0))
                                
                                # Convert image to RGBA if needed
                                if img.mode != 'RGBA':
                                    img = img.convert('RGBA')
                                
                                # Center image on canvas
                                paste_x = (canvas_size - width) // 2
                                paste_y = (canvas_size - height) // 2
                                canvas.paste(img, (paste_x, paste_y), img)
                                
                                # Save processed image
                                new_filename = os.path.splitext(os.path.basename(file_path))[0] + ".png"
                                output_path = os.path.join(output_folder, new_filename)
                                canvas.save(output_path)
                                processed_count += 1
                    except Exception as e:
                        st.warning(f"Error processing {os.path.basename(file_path)}: {str(e)}")
                
                status_text.empty()
                progress_bar.empty()
                
                if processed_count > 0:
                    st.success(f"✅ Successfully processed {processed_count} images")
                    
                    # Create zip file with processed images
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for root, dirs, files in os.walk(output_folder):
                            for file in files:
                                file_path = os.path.join(root, file)
                                rel_path = os.path.relpath(file_path, output_folder)
                                zip_file.write(file_path, rel_path)
                    
                    zip_buffer.seek(0)
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Processed Images (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="centered_images.zip",
                        mime="application/zip"
                    )
                    
                    # Show preview
                    with st.expander("🖼️ Preview Processed Images"):
                        sample_files = []
                        for root, dirs, files in os.walk(output_folder):
                            for file in files[:3]:  # Get first 3 files from each folder
                                if len(sample_files) < 6:  # Max 6 samples
                                    sample_files.append(os.path.join(root, file))
                        
                        cols = st.columns(3)
                        for i, file_path in enumerate(sample_files):
                            with cols[i % 3]:
                                img = PILImage.open(file_path)
                                st.image(img, caption=os.path.basename(file_path), use_column_width=True)
                        
                        if processed_count > 6:
                            st.info(f"Showing sample images. Total processed: {processed_count}")
                else:
                    st.warning("No images were processed (all images were already larger than the canvas size).")
            
            except Exception as e:
                st.error(f"An error occurred: {str(e)}")
    
    else:
        st.info("👆 Please upload a zip file with images to process.")
        
        # Instructions
        with st.expander("📖 Instructions"):
            st.markdown("""
            **How to use this tool:**
            
            1. Upload a zip file containing images
            2. Set the canvas size (default is 500x500 pixels)
            3. Click "Process Images"
            4. Download the processed images
            
            **Features:**
            - Centers images smaller than canvas size on transparent background
            - Images larger than canvas size are skipped
            - Outputs PNG files to preserve transparency
            - Preserves folder structure
            
            **Note:**
            - Only images smaller than the specified canvas size will be processed
            - Original aspect ratio is maintained
            """)
