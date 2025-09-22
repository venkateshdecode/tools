import os
import zipfile
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet

from helper import extract_zip_to_temp, erkenne_marken_aus_ordnern, generate_pdf_report, get_files_by_marke, analyze_files_by_filename, generate_two_section_pdf_report

def pdf_generator_tool():
    st.header("Asset Overview​")
    st.markdown("Upload a zip file containing brand assets to generate asset overview by brand​")
    
    # Analysis method selection
    analysis_method = st.radio(
        "Choose:",
        ["Overview by Brand", "Overview by Asset Type"],
        help="Overview by Brand: assets are organized by brand folders (they do NOT include brandname_ (underscore \"_\") in filename) ​. Overview by Asset Type: assets are organized by blocks AND include brandname_ (underscore \"_\") in filename)​"
    )
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a ZIP file",
        type=['zip'],
        key="pdf_zip_upload"
    )
    
    if uploaded_file is not None:
        try:
            # Extract zip file
            with st.spinner("Extracting zip file..."):
                temp_dir = extract_zip_to_temp(uploaded_file)
            
            # Find the actual input folder (in case zip has a root folder)
            input_folder = temp_dir
            items = os.listdir(temp_dir)
            if len(items) == 1 and os.path.isdir(os.path.join(temp_dir, items[0])):
                input_folder = os.path.join(temp_dir, items[0])
            
            if analysis_method == "Overview by Brand":
                # Original folder-based analysis
                marken = erkenne_marken_aus_ordnern(input_folder)
                
                if not marken:
                    st.error("No brand folders found in the uploaded zip file.")
                    st.info("Make sure your zip file contains folders named after your brands, with assets inside each folder.")
                    return
                
                st.success(f"Found brands: {', '.join(marken)}")
                
                # Get file counts for processing
                dateien_pro_marke = get_files_by_marke(input_folder, marken)
                
                # Options
                erste_marke = st.selectbox(
                    "Which brand should appear first in the overview?",
                    options=["Alphabetical order"] + marken
                )
                erste_marke = erste_marke if erste_marke != "Alphabetical order" else None
                
                # Generate PDF button
                if st.button("Generate PDF Report", type="primary"):
                    with st.spinner("Generating PDF report..."):
                        pdf_buffer, error = generate_pdf_report(input_folder, erste_marke)
                    
                    if error:
                        st.error(f"{error}")
                    elif pdf_buffer:
                        st.success("PDF report generated successfully!")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download PDF Report",
                            data=pdf_buffer.getvalue(),
                            file_name="Brand_Assets_Report.pdf",
                            mime="application/pdf"
                        )
                        
            else:
                # Filename-based analysis (Overview by Asset Type)
                marken_set, renamed_files_by_folder_and_marke, all_files = analyze_files_by_filename(input_folder)
                
                if not marken_set:
                    st.error("No brands found in filenames.")
                    st.info("Make sure your files are named with brand identifiers at the beginning (e.g., 'Brand1_asset.jpg').")
                    return
                
                st.success(f"Found brands from filenames: {', '.join(sorted(marken_set))}")
                
                # Show analysis preview
                with st.expander("Analysis Preview"):
                    total_files = len(all_files)
                    for folder, brands in renamed_files_by_folder_and_marke.items():
                        st.write(f"**{folder}:**")
                        for brand, files in brands.items():
                            st.write(f"- {brand}: {len(files)} files")
                
                # Options
                erste_marke = st.selectbox(
                    "Which brand should be numbered as 01?",
                    options=["Alphabetical order"] + sorted(list(marken_set))
                )
                erste_marke = erste_marke if erste_marke != "Alphabetical order" else None
                
                # Generate PDF button
                if st.button("Generate PDF Report", type="primary"):
                    with st.spinner("Generating PDF report..."):
                        # Use the two-section report function for filename-based analysis
                        pdf_buffer, error = generate_two_section_pdf_report(input_folder, erste_marke)
                    
                    if error:
                        st.error(f"{error}")
                    elif pdf_buffer:
                        st.success("PDF report generated successfully! (Contains two sections: by Brand and by Asset Type)")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download PDF Report (Two Sections)",
                            data=pdf_buffer.getvalue(),
                            file_name="IcAt_Overview by Brand and Asset Type.pdf",
                            mime="application/pdf"
                        )
        
        except zipfile.BadZipFile:
            st.error("Invalid zip file. Please upload a valid zip file.")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    
    else:
        st.info("👆 Please upload a zip file to get started.")
        
        # Instructions
        with st.expander("📖 Instructions"):
            st.markdown("""
            **How to use this tool:**
            
            **Overview by Brand:**
            - Create folders named after your brands
            - Place brand assets inside each brand folder
            - Assets do NOT need brand names in their filenames
            - Zip the entire structure
            - Generates single-section PDF report organized by brand
            
            **Overview by Asset Type:**
            - Name your files with brand identifiers at the beginning followed by an underscore (e.g., 'Brand1_asset.jpg', 'Brand2_document.pdf')
            - Organize files in any folder structure (blocks/categories)
            - The tool will extract brand names from filenames automatically
            - Generates two-section PDF report: Section 1 (by Asset Type/Block) and Section 2 (by Brand)
            
            **Supported file types:**
            - **Images:** JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP
            - **Text files:** TXT, MD, CSV
            """)
