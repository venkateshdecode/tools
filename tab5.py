import io
import os
import shutil
import zipfile
from pathlib import Path
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from helper import extract_zip_to_temp, analyze_files_by_filename, generate_filename_based_pdf_report_with_extensions, process_files, generate_excel_report



def brand_renamer_tool():
    st.header("Advanced Brand Processor")
    st.markdown("Automatically rename and organize brand assets with numbering and generate reports.")
   
    # Initialize session state for persistent results
    if 'processing_complete' not in st.session_state:
        st.session_state.processing_complete = False
    if 'processed_data' not in st.session_state:
        st.session_state.processed_data = {}
   
    uploaded_file = st.file_uploader(
        "Choose a ZIP file",
        type=['zip'],
    )
   
    # Reset processing state when new file is uploaded
    if uploaded_file and 'uploaded_file_name' in st.session_state:
        if st.session_state.uploaded_file_name != uploaded_file.name:
            st.session_state.processing_complete = False
            st.session_state.processed_data = {}
   
    if uploaded_file:
        # Store uploaded file name to detect changes
        st.session_state.uploaded_file_name = uploaded_file.name
       
        if not st.session_state.processing_complete:
            with st.spinner("Extracting and analyzing files..."):
                temp_dir = extract_zip_to_temp(uploaded_file)
                
                # Determine the actual input folder
                items = [item for item in os.listdir(temp_dir) if not item.startswith('.')]
                
                # ✅ ROOT CAUSE FIX: If flat structure, create a dedicated input subfolder
                if len(items) == 1 and os.path.isdir(os.path.join(temp_dir, items[0])):
                    # Single root folder - use it as input
                    input_folder = Path(temp_dir) / items[0]
                else:
                    # Flat structure - move everything to an input subfolder
                    input_folder = Path(temp_dir) / "input"
                    input_folder.mkdir(exist_ok=True)
                    
                    # Move all items to input folder
                    for item in items:
                        src = Path(temp_dir) / item
                        dst = input_folder / item
                        shutil.move(str(src), str(dst))
                
                # Now create output folder at temp_dir level (guaranteed to be separate)
                output_folder = Path(temp_dir) / "output"
                output_folder.mkdir(exist_ok=True)
                        
                marken_set, _, _ = analyze_files_by_filename(input_folder)
               
                if not marken_set:
                    st.error("No brands found in the uploaded files.")
                    return
               
                st.success(f"Detected brands: {', '.join(sorted(marken_set))}")
               
                # Brand numbering selection
                st.subheader("Brand Numbering")
                erste_marke = st.selectbox(
                    "Which brand should be number 01?",
                    options=sorted(marken_set),
                    index=0
                )
               
                marken_index = {erste_marke: "01"}
                aktuelle_nummer = 2
                for marke in sorted(marken_set):
                    if marke != erste_marke:
                        marken_index[marke] = f"{aktuelle_nummer:02d}"
                        aktuelle_nummer += 1
               
                if st.button("Process Files and Generate Reports", type="primary"):
                    with st.spinner("Processing files..."):
                        # Process and rename files
                        renamed_files, file_to_factorgroup = process_files(
                            input_folder,
                            marken_index,
                            output_folder
                        )
                       
                        # Rename files to preserve original extensions instead of forcing .png
                        for file_path in output_folder.iterdir():
                            if file_path.is_file():
                                # Get original extension (if it exists)
                                original_ext = file_path.suffix.lower()
                                
                                # If file has no extension, try to detect from content or keep as is
                                if not original_ext:
                                    # Try to find original extension from renamed_files mapping
                                    original_name = None
                                    for folder_data in renamed_files.values():
                                        for brand_data in folder_data.values():
                                            for orig_path, new_name in brand_data:
                                                if Path(new_name).name == file_path.name:
                                                    original_name = Path(orig_path).name
                                                    break
                                    
                                    if original_name:
                                        orig_ext = Path(original_name).suffix.lower()
                                        if orig_ext:
                                            new_file_path = file_path.with_suffix(orig_ext)
                                            try:
                                                file_path.rename(new_file_path)
                                            except Exception as e:
                                                print(f"Could not rename {file_path}: {e}")
                       
                        # Generate PDF with extensions - supports all file types
                        pdf_buffer, error = generate_filename_based_pdf_report_with_extensions(
                            input_folder,
                            erste_marke,
                            marken_index  # Pass marken_index for ID generation
                        )
                        if error:
                            st.error(f"PDF generation failed: {error}")
                       
                        excel_path = generate_excel_report(output_folder, marken_index, file_to_factorgroup)
                       
                        # Create individual file buffers
                        zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                            for file in output_folder.iterdir():
                                if file.is_file():
                                    zip_file.write(file, file.name)
                        zip_buffer.seek(0)
                       
                        # Create combined download with all files
                        combined_zip_buffer = io.BytesIO()
                        with zipfile.ZipFile(combined_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as combined_zip:
                            # Add processed files
                            for file in output_folder.iterdir():
                                if file.is_file():
                                    combined_zip.write(file, f"processed_files/{file.name}")
                           
                            # Add PDF report if available
                            if pdf_buffer:
                                combined_zip.writestr("reports/Brand_Assets_Report.pdf", pdf_buffer.getvalue())
                           
                            # Add Excel report
                            with open(excel_path, 'rb') as excel_file:
                                combined_zip.writestr("reports/Brand_Assets_Overview.xlsx", excel_file.read())
                       
                        combined_zip_buffer.seek(0)
                       
                        # Store processed data in session state
                        st.session_state.processed_data = {
                            'pdf_buffer': pdf_buffer,
                            'excel_path': excel_path,
                            'zip_buffer': zip_buffer,
                            'combined_zip_buffer': combined_zip_buffer,
                            'temp_dir': temp_dir,
                            'renamed_files': renamed_files
                        }
                        st.session_state.processing_complete = True
                       
                        # Force rerun to show download buttons
                        st.rerun()
       
        # Show download options if processing is complete
        if st.session_state.processing_complete:
            st.success("Processing complete!")
           
            # Display a preview of processed names with original extensions
            with st.expander("Preview Processed Names"):
                if 'renamed_files' in st.session_state.processed_data:
                    st.write("Sample processed filenames (with original extensions preserved):")
                    # Handle the case where renamed_files might be a defaultdict
                    renamed_files_data = st.session_state.processed_data['renamed_files']
                    
                    # Extract sample file names safely
                    sample_files = []
                    if hasattr(renamed_files_data, 'items'):
                        # It's a dictionary-like object
                        for folder_name, brand_data in list(renamed_files_data.items())[:3]:  # First 3 folders
                            for brand_name, files in list(brand_data.items())[:2]:  # First 2 brands per folder
                                for file_path, new_name in files[:2]:  # First 2 files per brand
                                    if isinstance(new_name, str):
                                        # Preserve original extension
                                        original_ext = Path(file_path).suffix
                                        pdf_name = Path(new_name).stem + original_ext
                                        sample_files.append((Path(file_path).name, pdf_name))
                                    if len(sample_files) >= 10:
                                        break
                                if len(sample_files) >= 10:
                                    break
                            if len(sample_files) >= 10:
                                break
                    
                    # Display the sample files
                    for old_name, pdf_name in sample_files:
                        st.write(f"{old_name} → {pdf_name}")
           
            # Option to start over
            if st.button("🔄 Process New File", help="Clear results and upload a new file"):
                st.session_state.processing_complete = False
                st.session_state.processed_data = {}
                if 'temp_dir' in st.session_state.processed_data:
                    try:
                        shutil.rmtree(st.session_state.processed_data['temp_dir'])
                    except:
                        pass
                st.rerun()
           
            st.markdown("---")
           
            # Combined download option (recommended)
            st.subheader("📦 Complete Package Download")
            st.markdown("**Recommended:** Download everything in one convenient package")
           
            if st.session_state.processed_data.get('combined_zip_buffer'):
                st.download_button(
                    label="🎁 Download Complete Package (All Files + Reports)",
                    data=st.session_state.processed_data['combined_zip_buffer'].getvalue(),
                    file_name="Brand_Assets_Complete_Package.zip",
                    mime="application/zip",
                    type="primary"
                )
           
            st.markdown("---")
           
            # Individual download options
            st.subheader("📁 Individual Downloads")
            st.markdown("Or download items separately:")
           
            col1, col2, col3 = st.columns(3)
           
            with col1:
                if st.session_state.processed_data.get('pdf_buffer'):
                    st.download_button(
                        label="📄 PDF Report",
                        data=st.session_state.processed_data['pdf_buffer'].getvalue(),
                        file_name="Brand_Assets_Report.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.info("PDF report not available")
           
            with col2:
                if st.session_state.processed_data.get('excel_path'):
                    try:
                        with open(st.session_state.processed_data['excel_path'], 'rb') as excel_file:
                            st.download_button(
                                label="📊 Excel Report",
                                data=excel_file.read(),
                                file_name="Brand_Assets_Overview.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except:
                        st.info("Excel report not available")
           
            with col3:
                if st.session_state.processed_data.get('zip_buffer'):
                    st.download_button(
                        label="📦 Processed Files",
                        data=st.session_state.processed_data['zip_buffer'].getvalue(),
                        file_name="Brand_Assets_Processed.zip",
                        mime="application/zip"
                    )
           
            st.markdown("---")
            st.info("💡 **Tip:** Downloads will remain available until you upload a new file or click 'Process New File'")
   
    else:
        # Reset session state when no file is uploaded
        if st.session_state.processing_complete:
            st.session_state.processing_complete = False
            st.session_state.processed_data = {}
       
        st.info("👆 Please upload a zip file to get started.")
       
        with st.expander("📖 Instructions"):
            st.markdown("""
            Upload a ZIP file containing brand asset folders. The tool will:
            1. Detect all brands automatically
            2. Let you choose which brand gets number 01
            3. Process and rename all files with proper numbering (preserving all file extensions)
            4. Generate a PDF report showing all assets with original filenames and extensions
            5. Create an Excel overview of the processed files (all file types)
            6. Package everything for easy download
            
            **Supported file types:** All extensions are supported including:
            - Images: .png, .jpg, .jpeg, .gif, .bmp, .svg, .webp
            - Videos: .mp4, .avi, .mov, .mkv, .webm
            - Audio: .mp3, .wav, .ogg, .flac, .m4a
            - Documents: .pdf, .doc, .docx, .txt, .xlsx
            - And any other file type!
            """)
