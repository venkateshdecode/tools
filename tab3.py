import io
import os
import tempfile
import zipfile
import cv2
import numpy as np
import streamlit as st
from PIL import Image as PILImage
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

ALLOWED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
ALLOWED_TEXT_EXTENSIONS = {'.txt', '.md', '.csv'}
ALLOWED_VIDEO_EXTENSIONS = {'.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.webm', '.m4v', '.3gp', '.ogv'}
ALLOWED_AUDIO_EXTENSIONS = {'.mp3', '.wav', '.aac', '.flac', '.ogg', '.m4a', '.wma'}
ALL_ALLOWED_EXTENSIONS = ALLOWED_IMAGE_EXTENSIONS.union(ALLOWED_TEXT_EXTENSIONS).union(ALLOWED_VIDEO_EXTENSIONS).union(ALLOWED_AUDIO_EXTENSIONS)

def white_to_transparent_tool():
    st.header("White Background to Transparent")
    st.markdown("Upload images to convert white backgrounds to transparent.")
    
    # Options
    col1, col2 = st.columns(2)
    with col1:
        crop_image = st.checkbox("Crop white background", value=False)
    with col2:
        overwrite = st.checkbox("Overwrite original files", value=False)
    
    # File upload
    uploaded_files = st.file_uploader(
        "Choose image files",
        type=list(ALLOWED_IMAGE_EXTENSIONS),
        accept_multiple_files=True,
        help="Upload images to convert white backgrounds to transparent.",
        key="transparent_upload"
    )
    
    if uploaded_files:
        if st.button("Process Images", type="primary"):
            temp_dir = tempfile.mkdtemp()
            processed_files = []
            failed_files = []
            
            # Save uploaded files to temp dir
            for uploaded_file in uploaded_files:
                file_path = os.path.join(temp_dir, uploaded_file.name)
                with open(file_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
            
            # Process images
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, uploaded_file in enumerate(uploaded_files):
                status_text.text(f"Processing {uploaded_file.name}... ({i+1}/{len(uploaded_files)})")
                progress_bar.progress((i + 1) / len(uploaded_files))
                
                try:
                    img = PILImage.open(os.path.join(temp_dir, uploaded_file.name)).convert("RGBA")
                    
                    # Convert to numpy array for OpenCV processing
                    img_array = np.array(img)
                    
                    if crop_image:
                        # Crop image
                        gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
                        th, threshed = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)
                        
                        kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (11,11))
                        morphed = cv2.morphologyEx(threshed, cv2.MORPH_CLOSE, kernel)
                        
                        cnts, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                        if cnts:  # Only crop if contours are found
                            cnt = sorted(cnts, key=cv2.contourArea)[-1]
                            x,y,w,h = cv2.boundingRect(cnt)
                            img_array = img_array[y:y+h, x:x+w]
                    
                    # Make white background transparent
                    gray = cv2.cvtColor(img_array, cv2.COLOR_BGR2GRAY)
                    th, threshed = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY_INV)
                    
                    kernel = cv2.getStructuringElement(cv2.MORPH_ELLIPSE, (11,11))
                    morphed = cv2.morphologyEx(threshed, cv2.MORPH_CLOSE, kernel)
                    
                    roi, _ = cv2.findContours(morphed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
                    
                    mask = np.zeros(img_array.shape, img_array.dtype)
                    if roi:  # Only proceed if contours are found
                        cv2.fillPoly(mask, roi, (255,)*img_array.shape[2])
                        masked_image = cv2.bitwise_and(img_array, mask)
                        
                        # Convert back to PIL Image
                        result_img = PILImage.fromarray(masked_image, mode="RGBA")
                        
                        # Save the processed image
                        new_filename = os.path.splitext(uploaded_file.name)[0] + ".png"
                        processed_path = os.path.join(temp_dir, 
                                                       new_filename)
                        result_img.save(processed_path)
                        processed_files.append(processed_path)
                    else:
                        failed_files.append(uploaded_file.name)
                except Exception as e:
                    failed_files.append(uploaded_file.name)
                    st.warning(f"Failed to process {uploaded_file.name}: {str(e)}")
            
            status_text.empty()
            progress_bar.empty()
            
            if processed_files:
                st.success(f"✅ Successfully processed {len(processed_files)} images")
                
                # Create zip file with processed images
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for file_path in processed_files:
                        zip_file.write(file_path, os.path.basename(file_path))
                
                zip_buffer.seek(0)
                
                # Download button
                st.download_button(
                    label="📥 Download Processed Images (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="transparent_images.zip",
                    mime="application/zip"
                )
                
                # Show preview
                with st.expander("🖼️ Preview Processed Images"):
                    cols = st.columns(3)
                    for i, file_path in enumerate(processed_files[:6]):  # Show max 6 images
                        with cols[i % 3]:
                            img = PILImage.open(file_path)
                            st.image(img, caption=os.path.basename(file_path), use_column_width=True)
                    
                    if len(processed_files) > 6:
                        st.info(f"Showing first 6 images. Total processed: {len(processed_files)}")
            
            if failed_files:
                st.warning(f"⚠️ Failed to process {len(failed_files)} files: {', '.join(failed_files)}")
    
    else:
        st.info("👆 Please upload image files to get started.")
        
        # Instructions
        with st.expander("📖 Instructions"):
            st.markdown("""
            **How to use this tool:**
            
            1. Upload one or more image files
            2. Choose processing options:
               - Crop white background: Removes excess white space around the image
               - Overwrite original files: Not applicable in this web version (always creates new files)
            3. Click "Process Images"
            4. Download the processed images with transparent backgrounds
            
            **Note:**
            - Output images will always be in PNG format to support transparency
            - Only white backgrounds will be made transparent (threshold of 240/255)
            - Complex images may require manual adjustment for best results
            """)