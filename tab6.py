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


def resize_with_transparent_canvas_tool():
    st.header("Resize with Transparent Canvas")
    st.markdown("Resize images to 400px on the largest side while maintaining aspect ratio on transparent background.")
    
    # File upload with unique key
    uploaded_file = st.file_uploader(
        "Choose a ZIP file with images",
        type=['zip'],
        help="Upload a zip file containing images to resize.",
        key="resize_transparent_upload"
    )
    
    if uploaded_file:
        # Show the fixed resize setting
        st.info("Images will be resized to 400px on the largest side (width or height), maintaining aspect ratio.")
        
        # Process button with unique key
        if st.button("Process Images", type="primary", key="process_resize_transparent_images"):
            try:
                # Extract zip file
                with st.spinner("Extracting zip file..."):
                    temp_dir = extract_zip_to_temp(uploaded_file)
                    input_folder = temp_dir
                    items = os.listdir(temp_dir)
                    if len(items) == 1 and os.path.isdir(os.path.join(temp_dir, items[0])):
                        input_folder = os.path.join(temp_dir, items[0])
                
                # Create output folder
                output_folder = os.path.join(temp_dir, "resized_transparent_output")
                os.makedirs(output_folder, exist_ok=True)
                
                # Process images
                processed_count = 0
                error_count = 0
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                all_files = []
                for root, dirs, files in os.walk(input_folder):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        if ext in ALLOWED_IMAGE_EXTENSIONS:
                            all_files.append(os.path.join(root, file))
                
                total_files = len(all_files)
                
                for i, file_path in enumerate(all_files):
                    status_text.text(f"Processing {i+1}/{total_files}: {os.path.basename(file_path)}...")
                    progress_bar.progress((i + 1) / total_files)
                    
                    try:
                        relative_path = os.path.relpath(file_path, input_folder)
                        new_path = os.path.join(output_folder, os.path.splitext(relative_path)[0] + ".png")
                        os.makedirs(os.path.dirname(new_path), exist_ok=True)
                        
                        with PILImage.open(file_path) as img:
                            img = img.convert("RGBA")
                            orig_w, orig_h = img.size
                            
                            # Calculate scaling factor based on largest side
                            max_dimension = 400
                            if orig_w >= orig_h:
                                # Width is larger or equal
                                scale_factor = max_dimension / orig_w
                                new_w = max_dimension
                                new_h = int(orig_h * scale_factor)
                            else:
                                # Height is larger
                                scale_factor = max_dimension / orig_h
                                new_h = max_dimension
                                new_w = int(orig_w * scale_factor)
                            
                            # Resize image while maintaining aspect ratio
                            img_resized = img.resize((new_w, new_h), PILImage.LANCZOS)
                            
                            # Create transparent canvas with the exact resized dimensions
                            canvas = PILImage.new("RGBA", (new_w, new_h), (0, 0, 0, 0))
                            
                            # The resized image fits perfectly on the canvas (no centering needed)
                            canvas.paste(img_resized, (0, 0), mask=img_resized)
                            
                            canvas.save(new_path)
                            processed_count += 1
                    except Exception as e:
                        error_count += 1
                        st.warning(f"Error processing {os.path.basename(file_path)}: {str(e)}")
                
                status_text.empty()
                progress_bar.empty()
                
                # Results summary
                st.success(f"""
                Processing complete!
                - ✅ Processed: {processed_count} images
                - ❌ Errors: {error_count} images
                - All images resized to 400px on largest side
                """)
                
                if processed_count > 0:
                    # Create zip file with processed images
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for root, dirs, files in os.walk(output_folder):
                            for file in files:
                                file_path = os.path.join(root, file)
                                rel_path = os.path.relpath(file_path, output_folder)
                                zip_file.write(file_path, rel_path)
                    
                    zip_buffer.seek(0)
                    
                    # Download button with unique key
                    st.download_button(
                        label="📥 Download Resized Images (ZIP)",
                        data=zip_buffer.getvalue(),
                        file_name="resized_400px_images.zip",
                        mime="application/zip",
                        key="download_resized_transparent"
                    )
                    
                    # Show preview
                    with st.expander("🖼️ Preview Resized Images"):
                        sample_files = []
                        for root, dirs, files in os.walk(output_folder):
                            for file in files[:3]:  # Get first 3 files from each folder
                                if len(sample_files) < 6:  # Max 6 samples
                                    sample_files.append(os.path.join(root, file))
                        
                        if sample_files:
                            cols = st.columns(3)
                            for i, file_path in enumerate(sample_files):
                                with cols[i % 3]:
                                    img = PILImage.open(file_path)
                                    st.image(img, caption=f"{os.path.basename(file_path)} ({img.size[0]}x{img.size[1]})", use_column_width=True)
                        
                        if processed_count > 6:
                            st.info(f"Showing sample images. Total processed: {processed_count}")
            
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
    
    else:
        st.info("👆 Please upload a zip file with images to resize.")
        
        # Instructions
        with st.expander("📖 Instructions"):
            st.markdown("""
            **How to use this tool:**
            
            1. Upload a zip file containing images
            2. All images will be resized to 400px on their largest side (width or height)
            3. Aspect ratio is maintained - smaller dimension will be scaled proportionally
            4. Click "Process Images"
            5. Download the resized images with transparent backgrounds
            
            **Examples:**
            - 1000x800 image → 400x320 image (width was larger)
            - 600x1200 image → 200x400 image (height was larger)  
            - 400x400 image → 400x400 image (already correct size)
            
            **Features:**
            - Maintains original aspect ratio
            - Outputs PNG files to preserve transparency
            - Preserves folder structure
            - Fixed 400px maximum on largest dimension
            """)
