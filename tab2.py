import io
import os
import shutil
import zipfile
from pathlib import Path
import pandas as pd
import streamlit as st
from helper import extract_zip_to_temp

ALLOWED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
ALLOWED_TEXT_EXTENSIONS = {'.txt', '.md', '.csv'}
ALLOWED_VIDEO_EXTENSIONS = {'.mp4', '.avi', '.mov', '.wmv', '.flv', '.mkv', '.webm', '.m4v', '.3gp', '.ogv'}
ALLOWED_AUDIO_EXTENSIONS = {'.mp3', '.wav', '.aac', '.flac', '.ogg', '.m4a', '.wma'}
ALL_ALLOWED_EXTENSIONS = ALLOWED_IMAGE_EXTENSIONS.union(ALLOWED_TEXT_EXTENSIONS).union(ALLOWED_VIDEO_EXTENSIONS).union(ALLOWED_AUDIO_EXTENSIONS)

def file_renamer_tool():
    st.header("Brand File Renamer")
    st.markdown("Upload a zip file containing brand folders to automatically rename files with brand prefixes.")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a ZIP file",
        type=['zip'],
        # help="Upload a zip file containing your brand folders with assets.",
        key="renamer_zip_upload"
    )
    
    if uploaded_file is not None:
        try:
            # Extract zip file
            with st.spinner("Extracting zip file..."):
                temp_dir = extract_zip_to_temp(uploaded_file)
            
            # Find the actual input folder (in case zip has a root folder)
            input_folder = Path(temp_dir)
            items = os.listdir(temp_dir)
            if len(items) == 1 and os.path.isdir(os.path.join(temp_dir, items[0])):
                input_folder = Path(temp_dir) / items[0]
            
            # Create output folder
            output_folder = input_folder.parent / (input_folder.name + "_renamed")
            output_folder.mkdir(exist_ok=True)
            
            # Process files
            total_files = 0
            processed_files = []
            file_type_counts = {"images": 0, "text": 0, "videos": 0}
            
            for subfolder in input_folder.iterdir():
                if subfolder.is_dir():
                    markenname = subfolder.name
                    new_subfolder = output_folder / markenname
                    new_subfolder.mkdir(parents=True, exist_ok=True)
                    
                    for file in subfolder.iterdir():
                        if file.is_file() and file.suffix.lower() in ALL_ALLOWED_EXTENSIONS:
                            new_name = f"{markenname}_{file.name}"
                            target_path = new_subfolder / new_name
                            shutil.copy2(file, target_path)
                            processed_files.append((file.name, new_name, markenname))
                            total_files += 1
                            
                            # Count file types
                            if file.suffix.lower() in ALLOWED_IMAGE_EXTENSIONS:
                                file_type_counts["images"] += 1
                            elif file.suffix.lower() in ALLOWED_TEXT_EXTENSIONS:
                                file_type_counts["text"] += 1
                            elif file.suffix.lower() in ALLOWED_VIDEO_EXTENSIONS:
                                file_type_counts["videos"] += 1
            
            if total_files > 0:
                st.success(f"Successfully renamed and copied {total_files} files")
                
                # Show file type breakdown
                st.info(f"File types processed: {file_type_counts['images']} images, {file_type_counts['text']} text files, {file_type_counts['videos']} videos")
                
                # Show preview
                with st.expander("📋 Preview renamed files"):
                    preview_df = pd.DataFrame(
                        processed_files,  
                        columns=["Original Name", "New Name", "Brand"]
                    )
                    st.dataframe(preview_df)
                    
                    if total_files > 20:
                        st.info(f"Total processed: {total_files}")
                
                # Create zip file with renamed files
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for root, dirs, files in os.walk(output_folder):
                        for file in files:
                            file_path = Path(root) / file
                            rel_path = os.path.relpath(file_path, output_folder)
                            zip_file.write(file_path, rel_path)
                
                zip_buffer.seek(0)
                
                # Download button
                st.download_button(
                    label="📥 Download Renamed Files (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="brand_files_renamed.zip",
                    mime="application/zip"
                )
            else:
                st.warning("No files were found to rename in the uploaded zip file.")
        
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    
    else:
        st.info("👆 Please upload a zip file to get started.")
        
        # Instructions
        with st.expander("📖 Instructions"):
            st.markdown("""
            **How to use this tool:**
            
            1. Create folders named after your brands
            2. Place brand assets inside each brand folder
            3. Zip the entire structure
            4. Upload the zip file
            
            **The tool will:**
            - Create a new folder structure matching the original
            - Rename all files with the brand prefix (e.g., "Brand1_filename.jpg")
            - Provide a zip file with all renamed files
            
            **Supported file types:**
            - **Images:** JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP
            - **Text files:** TXT, MD, CSV
            - **Videos:** MP4, AVI, MOV, WMV, FLV, MKV, WEBM, M4V, 3GP, OGV
            """)
