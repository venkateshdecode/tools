import os
import tempfile
import zipfile
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Image, Paragraph, Spacer, PageBreak
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from PIL import Image as PILImage
import io
from pathlib import Path
import shutil

# Configure Streamlit page
st.set_page_config(
    page_title="Brand Assets Tools",
    page_icon="🔧",
    layout="wide"
)

# Definierte Dateitypen
ALLOWED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
ALLOWED_TEXT_EXTENSIONS = {'.txt', '.md', '.csv'}
ALL_ALLOWED_EXTENSIONS = ALLOWED_IMAGE_EXTENSIONS.union(ALLOWED_TEXT_EXTENSIONS)

def extract_zip_to_temp(uploaded_file):
    """Extract uploaded zip file to temporary directory"""
    temp_dir = tempfile.mkdtemp()
    
    with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
        zip_ref.extractall(temp_dir)
    
    return temp_dir

def extract_images_from_xlsx(xlsx_file, output_dir):
    """Extract images from Excel file to specified directory"""
    count = 0
    extracted_files = []
    
    try:
        with zipfile.ZipFile(xlsx_file, 'r') as z:
            for file_info in z.infolist():
                if file_info.filename.startswith("xl/media/"):
                    filename = Path(file_info.filename).name
                    target_path = Path(output_dir) / filename
                    
                    with z.open(file_info) as source, open(target_path, 'wb') as target:
                        shutil.copyfileobj(source, target)
                    
                    extracted_files.append(str(target_path))
                    count += 1
        
        return count, extracted_files, None
    except Exception as e:
        return 0, [], str(e)

def erkenne_marken_aus_ordnern(input_folder):
    """Recognize brands from subfolders"""
    try:
        return sorted([
            f for f in os.listdir(input_folder)
            if os.path.isdir(os.path.join(input_folder, f)) and not f.startswith('.')
        ])
    except Exception as e:
        st.error(f"Error reading folders: {str(e)}")
        return []

def get_files_by_marke(input_folder, marken):
    """Get files per brand"""
    spalteninhalte = {marke: [] for marke in marken}
    
    for marke in marken:
        folder_path = os.path.join(input_folder, marke)
        if os.path.isdir(folder_path):
            try:
                dateien = [
                    os.path.join(folder_path, f)
                    for f in os.listdir(folder_path)
                    if not f.startswith(".") and os.path.splitext(f)[1].lower() in ALL_ALLOWED_EXTENSIONS
                ]
                spalteninhalte[marke] = dateien
            except Exception as e:
                st.warning(f"Error reading files from {marke}: {str(e)}")
                spalteninhalte[marke] = []
    
    return spalteninhalte

def build_marken_einzelseiten(alle_dateien_pro_marke, styles, gesamtbreite):
    """Build individual brand overview pages"""
    elements = []
    max_bildhoehe = 100

    for marke, dateien in alle_dateien_pro_marke.items():
        elements.append(Paragraph(f"Brand Overview: {marke}", styles['Heading2']))
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"<b>Total Assets:</b> {len(dateien)}", styles['Normal']))
        elements.append(Spacer(1, 12))

        if not dateien:
            elements.append(Paragraph("No assets found in this brand folder.", styles['Normal']))
            elements.append(PageBreak())
            continue

        table_data = []
        row_data = []
        spaltenbreite = gesamtbreite / 3

        for i, file_path in enumerate(dateien):
            ext = os.path.splitext(file_path)[1].lower()
            filename = os.path.basename(file_path)

            if ext in ALLOWED_IMAGE_EXTENSIONS:
                try:
                    pil_img = PILImage.open(file_path)
                    width, height = pil_img.size
                    aspect = height / width
                    new_width = spaltenbreite * 0.9
                    new_height = min(new_width * aspect, max_bildhoehe)
                    new_width = new_height / aspect

                    img = Image(file_path, width=new_width, height=new_height)
                    cell = [img, Spacer(1, 4), Paragraph(filename, styles['Normal'])]
                except Exception as e:
                    cell = Paragraph(f"📄 {filename} (Error loading image)", styles['Normal'])
            else:
                cell = Paragraph(f"📄 {filename}", styles['Normal'])

            row_data.append(cell)
            if len(row_data) == 3:
                table_data.append(row_data)
                row_data = []

        if row_data:
            while len(row_data) < 3:
                row_data.append("")
            table_data.append(row_data)

        if table_data:
            table = Table(table_data, colWidths=[spaltenbreite] * 3)
            table.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 1, colors.grey),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]))
            elements.append(table)
        
        elements.append(PageBreak())
    
    return elements

def generate_pdf_report(input_folder, erste_marke=None):
    """Generate the complete PDF report"""
    styles = getSampleStyleSheet()
    gesamtbreite = 480

    marken = erkenne_marken_aus_ordnern(input_folder)
    if not marken:
        return None, "No brand folders found in the uploaded zip file."

    # Reorder brands if specified
    if erste_marke and erste_marke in marken:
        marken = [erste_marke] + [m for m in marken if m != erste_marke]

    dateien_pro_marke = get_files_by_marke(input_folder, marken)
    total_all = sum(len(dateien_pro_marke[marke]) for marke in marken)

    alle_elements = []

    # Overview page
    alle_elements.append(Paragraph("Brand Assets Overview", styles['Title']))
    alle_elements.append(Spacer(1, 20))
    
    übersicht_text = ", ".join([f"{marke}: {len(dateien_pro_marke[marke])}" for marke in marken])
    alle_elements.append(Paragraph(f"<b>Assets per brand:</b> {übersicht_text}", styles['Normal']))
    alle_elements.append(Spacer(1, 12))
    alle_elements.append(Paragraph(f"<b>Total assets:</b> {total_all}", styles['Normal']))
    alle_elements.append(PageBreak())

    # Individual brand pages
    markenseiten = build_marken_einzelseiten(dateien_pro_marke, styles, gesamtbreite)
    alle_elements.extend(markenseiten)

    # Generate PDF
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4)
    
    try:
        doc.build(alle_elements)
        pdf_buffer.seek(0)
        return pdf_buffer, None
    except Exception as e:
        return None, f"Error generating PDF: {str(e)}"

def pdf_generator_tool():
    st.header("📊 Brand Assets PDF Generator")
    st.markdown("Upload a zip file containing brand folders with assets to generate a comprehensive PDF report.")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a ZIP file",
        type=['zip'],
        help="Upload a zip file containing folders named after your brands, with assets inside each folder.",
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
            
            # Recognize brands
            marken = erkenne_marken_aus_ordnern(input_folder)
            
            if not marken:
                st.error("❌ No brand folders found in the uploaded zip file.")
                st.info("Make sure your zip file contains folders named after your brands, with assets inside each folder.")
                return
            
            st.success(f"🔍 Found brands: {', '.join(marken)}")
            
            # Get file counts for processing
            dateien_pro_marke = get_files_by_marke(input_folder, marken)
            
            # Options
            erste_marke = st.selectbox(
                "Which brand should appear first in the overview?",
                options=["Alphabetical order"] + marken
            )
            erste_marke = erste_marke if erste_marke != "Alphabetical order" else None
            
            # Generate PDF button
            if st.button("🎯 Generate PDF Report", type="primary"):
                with st.spinner("Generating PDF report..."):
                    pdf_buffer, error = generate_pdf_report(input_folder, erste_marke)
                
                if error:
                    st.error(f"❌ {error}")
                elif pdf_buffer:
                    st.success("✅ PDF report generated successfully!")
                    
                    # Download button
                    st.download_button(
                        label="📥 Download PDF Report",
                        data=pdf_buffer.getvalue(),
                        file_name="Brand_Assets_Report.pdf",
                        mime="application/pdf"
                    )
        
        except zipfile.BadZipFile:
            st.error("❌ Invalid zip file. Please upload a valid zip file.")
        except Exception as e:
            st.error(f"❌ An error occurred: {str(e)}")
    
    else:
        st.info("👆 Please upload a zip file to get started.")
        
        # Instructions for PDF generator
        with st.expander("📖 Instructions for PDF Generator"):
            st.markdown("""
            **How to use this tool:**
            
            1. **Prepare your zip file:**
               - Create folders named after your brands
               - Place brand assets (images, text files) inside each brand folder
               - Zip the entire structure
            
            2. **Supported file types:**
               - **Images:** JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP
               - **Text files:** TXT, MD, CSV

            
            3. **Upload and generate your PDF report!**
            """)

def excel_extractor_tool():
    st.header("🖼️ Excel Image Extractor")
    st.markdown("Upload Excel files (.xlsx) to extract embedded images from them.")
    
    # File upload for Excel files
    uploaded_files = st.file_uploader(
        "Choose Excel file(s)",
        type=['xlsx'],
        accept_multiple_files=True,
        help="Upload one or more .xlsx files to extract embedded images.",
        key="excel_upload"
    )
    
    if uploaded_files:
        if st.button("🔍 Extract Images", type="primary", key="extract_btn"):
            temp_dir = tempfile.mkdtemp()
            all_extracted_files = []
            total_images = 0
            
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, uploaded_file in enumerate(uploaded_files):
                status_text.text(f"Processing {uploaded_file.name}...")
                progress_bar.progress((i + 1) / len(uploaded_files))
                
                # Save uploaded file temporarily
                temp_xlsx_path = os.path.join(temp_dir, uploaded_file.name)
                with open(temp_xlsx_path, 'wb') as f:
                    f.write(uploaded_file.getbuffer())
                
                # Create output directory for this file
                output_dir = os.path.join(temp_dir, f"{Path(uploaded_file.name).stem}_images")
                os.makedirs(output_dir, exist_ok=True)
                
                # Extract images
                count, extracted_files, error = extract_images_from_xlsx(temp_xlsx_path, output_dir)
                
                if error:
                    st.error(f"❌ Error processing {uploaded_file.name}: {error}")
                elif count > 0:
                    st.success(f"✅ Extracted {count} image(s) from {uploaded_file.name}")
                    all_extracted_files.extend(extracted_files)
                    total_images += count
                else:
                    st.warning(f"⚠️ No images found in {uploaded_file.name}")
            
            status_text.empty()
            progress_bar.empty()
            
            if total_images > 0:
                st.success(f"🎉 Total: {total_images} images extracted from {len(uploaded_files)} file(s)")
                
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
                    label="📥 Download All Extracted Images (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name="extracted_images.zip",
                    mime="application/zip"
                )
                
                # Show preview of extracted images
                with st.expander("🖼️ Preview Extracted Images"):
                    cols = st.columns(3)
                    col_idx = 0
                    
                    for file_path in all_extracted_files[:12]:  # Show max 12 images
                        try:
                            with cols[col_idx % 3]:
                                img = PILImage.open(file_path)
                                st.image(img, caption=os.path.basename(file_path), use_column_width=True)
                            col_idx += 1
                        except Exception as e:
                            st.write(f"Could not preview {os.path.basename(file_path)}")
                    
                    if len(all_extracted_files) > 12:
                        st.info(f"Showing first 12 images. Total extracted: {len(all_extracted_files)}")
            
            else:
                st.info("No images were found in the uploaded Excel file(s).")
    
    else:
        st.info("👆 Please upload Excel file(s) to get started.")
        
        # Instructions for Excel extractor
        with st.expander("📖 How to use Excel Image Extractor"):
            st.markdown("""
            **How it works:**
            
            1. Upload one or more Excel files (.xlsx format)
            2. The tool will scan for embedded images in the Excel files
            3. All found images will be extracted and made available for download
            4. Images are packaged in a ZIP file for easy download
            
            **Note:** 
            - Only .xlsx files are supported (not .xls)
            - Images must be embedded in the Excel file (not linked)
            - Common image formats (PNG, JPG, etc.) are extracted
            """)

def main():
    st.title("🔧 Brand Assets Tools")

    tab1, tab2 = st.tabs(["📊 PDF Report Generator", "🖼️ Excel Image Extractor"])

    with tab1:
        pdf_generator_tool()
    with tab2:
        excel_extractor_tool()






if __name__ == "__main__":
    main()
