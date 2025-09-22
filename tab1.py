import io
import os
import tempfile
import zipfile
from pathlib import Path
import streamlit as st
from reportlab.lib.pagesizes import A4

from helper import get_excel_worksheets_simple, extract_images_from_xlsx_by_sheet


def excel_extractor_tool():
    st.header("Excel Image Extractor")
    st.markdown("Upload Excel files (.xlsx) to extract embedded images from specific sheets.")
    
    # File upload for Excel files
    uploaded_files = st.file_uploader(
        "Choose Excel file(s)",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Upload one or more .xlsx files to extract embedded images.",
        key="excel_upload_fixed"
    )
    
    if uploaded_files:
        # For each file, let user select sheets
        sheet_selections = {}
        
        for uploaded_file in uploaded_files:
            # Temporarily save file to read sheet names
            temp_dir = tempfile.mkdtemp()
            temp_path = os.path.join(temp_dir, uploaded_file.name)
            
            with open(temp_path, 'wb') as f:
                f.write(uploaded_file.getbuffer())
            
            try:
                # Get available sheets using the simple function
                available_sheets = get_excel_worksheets_simple(temp_path)
                
                if available_sheets:
                    st.subheader(f"Select sheets from: {uploaded_file.name}")
                    
                    # Add "All sheets" option
                    all_options = ["All worksheets"] + available_sheets
                    
                    selected = st.multiselect(
                        f"Choose sheets to extract images from:",
                        all_options,
                        default=["All worksheets"],
                        help="Select specific sheets or 'All worksheets' to extract from all sheets",
                        key=f"sheets_fixed_{uploaded_file.name}"
                    )
                    
                    if "All worksheets" in selected:
                        sheet_selections[uploaded_file.name] = available_sheets
                    else:
                        sheet_selections[uploaded_file.name] = [s for s in selected if s != "All worksheets"]
                else:
                    st.warning(f"Could not read sheets from {uploaded_file.name}")
                    sheet_selections[uploaded_file.name] = []
                    
            except Exception as e:
                st.error(f"Error reading {uploaded_file.name}: {str(e)}")
                sheet_selections[uploaded_file.name] = []
            
            # Clean up
            try:
                os.remove(temp_path)
                os.rmdir(temp_dir)
            except:
                pass
        
        if st.button("Extract Images", type="primary", key="extract_btn_fixed"):
            temp_dir = tempfile.mkdtemp()
            all_extracted_files = []
            total_images = 0
            extraction_results = []
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, uploaded_file in enumerate(uploaded_files):
                file_name = uploaded_file.name
                selected_sheets = sheet_selections.get(file_name, [])
                
                if not selected_sheets:
                    st.warning(f"No sheets selected for {file_name}, skipping...")
                    continue
                
                status_text.text(f"Processing {file_name}...")
                progress_bar.progress((i + 1) / len(uploaded_files))
                
                # Save uploaded file temporarily
                temp_xlsx_path = os.path.join(temp_dir, file_name)
                with open(temp_xlsx_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
                
                # Create output directory for this file
                output_dir = os.path.join(temp_dir, f"{Path(file_name).stem}_images")
                os.makedirs(output_dir, exist_ok=True)
                
                # Extract images for each selected sheet individually
                file_total = 0
                for sheet_name in selected_sheets:
                    sheet_output_dir = os.path.join(output_dir, sheet_name.replace('/', '_').replace('\\', '_'))
                    os.makedirs(sheet_output_dir, exist_ok=True)
                    
                    count, extracted_files, error, _ = extract_images_from_xlsx_by_sheet(
                        temp_xlsx_path, 
                        sheet_output_dir, 
                        sheet_name
                    )
                    
                    if error:
                        st.warning(f"Issue with {file_name} - {sheet_name}: {error}")
                    
                    if count > 0:
                        extraction_results.append(f"✅ {file_name} - {sheet_name}: {count} images")
                        all_extracted_files.extend(extracted_files)
                        file_total += count
                    else:
                        extraction_results.append(f"⚪ {file_name} - {sheet_name}: No images found")
                
                total_images += file_total
                if file_total > 0:
                    st.success(f"Extracted {file_total} images from {file_name}")
                else:
                    st.info(f"No images found in {file_name}")
            
            status_text.empty()
            progress_bar.empty()
            
            # Show detailed extraction results
            with st.expander("Detailed Extraction Results"):
                for result in extraction_results:
                    st.write(result)
            
            if total_images > 0:
                st.success(f"Total: {total_images} images extracted from {len(uploaded_files)} file(s)")
                
                # Show extracted files summary
                with st.expander("Extracted Files Summary"):
                    for file_path in all_extracted_files[:20]:  # Show first 20
                        st.write(f"📄 {os.path.relpath(file_path, temp_dir)}")
                    if len(all_extracted_files) > 20:
                        st.info(f"Showing first 20 files. Total extracted: {len(all_extracted_files)}")
                
                # Create zip file with all extracted images
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                    for file_path in all_extracted_files:
                        # Get relative path for zip structure
                        rel_path = os.path.relpath(file_path, temp_dir)
                        zip_file.write(file_path, rel_path)
                
                zip_buffer.seek(0)
                
                # Download button for all extracted images
                st.download_button(
                    label="Download All Extracted Images (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="extracted_images_by_sheet.zip",
                    mime="application/zip"
                )
            
            else:
                st.info("No images were found in the selected sheets of the uploaded Excel file(s).")
    
    else:
        st.info("Please upload Excel file(s) to get started.")
        
        with st.expander("Instructions"):
            st.markdown("""
 
            """)