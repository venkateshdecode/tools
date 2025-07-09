import sys
import warnings

try:
    import cv2
    CV2_AVAILABLE = True
except ImportError as e:
    CV2_AVAILABLE = False
    warnings.warn(f"OpenCV not available: {str(e)} - some features disabled")

import os
import re
import tempfile
import zipfile
import streamlit as st
from pathlib import Path
from collections import defaultdict
from reportlab.lib.pagesizes import A4
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Image as RLImage, 
    Paragraph, Spacer, PageBreak, KeepInFrame
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from PIL import Image as PILImage, ImageOps
import io
import shutil
import numpy as np
import pandas as pd





st.set_page_config(
    page_title="Brand Assets Tools",
    page_icon="🔧",
    layout="wide"
)

# Defined file types
ALLOWED_IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif', '.tiff', '.webp'}
ALLOWED_TEXT_EXTENSIONS = {'.txt', '.md', '.csv'}
ALL_ALLOWED_EXTENSIONS = ALLOWED_IMAGE_EXTENSIONS.union(ALLOWED_TEXT_EXTENSIONS)

# Helper functions from both files
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

def extract_brand(file_stem):
    """Extract brand name from filename using pattern matching"""
    parts = re.split(r'[ _\-]', file_stem)
    return re.sub(r'[^a-z0-9]', '', parts[0].lower()) if parts else "unknown"

def erkenne_marken_aus_ordnern(input_folder):
    """Recognize brands from subfolders (original method)"""
    try:
        return sorted([
            f for f in os.listdir(input_folder)
            if os.path.isdir(os.path.join(input_folder, f)) and not f.startswith('.')
        ])
    except Exception as e:
        st.error(f"Error reading folders: {str(e)}")
        return []

def get_files_by_marke(input_folder, marken):
    """Get files per brand from folders"""
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

def get_asset_cell(file_path, filename, col_count):
    """Create optimized asset cell for PDF with better image handling"""
    ext = Path(file_path).suffix.lower()
    styles = getSampleStyleSheet()
    
    try:
        if ext in ALLOWED_IMAGE_EXTENSIONS:
            with PILImage.open(file_path) as img:
                orig_width, orig_height = img.size
                available_width = A4[0] - 40
                max_width = available_width / col_count
                max_height = 120

                aspect_ratio = orig_width / orig_height
                width = max_width
                height = width / aspect_ratio

                if height > max_height:
                    height = max_height
                    width = height * aspect_ratio

                buffer = io.BytesIO()
                img.convert("RGB").save(buffer, format='PNG')
                buffer.seek(0)
                rl_img = RLImage(buffer, width=width, height=height)

                frame = KeepInFrame(max_width, max_height + 30, content=[
                    rl_img,
                    Paragraph(filename, styles['Normal'])
                ], hAlign='CENTER')

                return frame
        else:
            return Paragraph(f"{filename}", styles['Normal'])
    except Exception as e:
        return Paragraph(f"[Image Error]<br/>{filename}", styles['Normal'])

def build_marken_einzelseiten(alle_dateien_pro_marke, styles, gesamtbreite):
    """Build individual brand overview pages (original method)"""
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

                    img = RLImage(file_path, width=new_width, height=new_height)
                    cell = [img, Spacer(1, 4), Paragraph(filename, styles['Normal'])]
                except Exception as e:
                    cell = Paragraph(f"{filename} (Error loading image)", styles['Normal'])
            else:
                cell = Paragraph(f"{filename}", styles['Normal'])

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

def add_brand_pages(elements, marken_spalten, renamed_files_by_folder_and_marke):
    """Add brand overview pages with optimized layout"""
    styles = getSampleStyleSheet()
    for nummer, marke in marken_spalten:
        assets = []
        for folder in renamed_files_by_folder_and_marke:
            assets.extend(renamed_files_by_folder_and_marke[folder].get(marke, []))

        if not assets:
            continue

        elements.append(PageBreak())
        elements.append(Paragraph(f"<b>Brand Overview: {nummer} – {marke}</b>", styles['Heading2']))
        elements.append(Spacer(1, 6))
        elements.append(Paragraph(f"Number of assets: {len(assets)}", styles['Normal']))
        elements.append(Spacer(1, 10))

        headers = ["Asset"] * 4
        data = [headers]
        row = []
        for i, (pfad, name) in enumerate(assets):
            cell = get_asset_cell(pfad, name, 4)
            row.append(cell)
            if len(row) == 4:
                data.append(row)
                row = []
        if row:
            row.extend([""] * (4 - len(row)))
            data.append(row)

        table = Table(data, colWidths=(A4[0] - 40) / 4)
        table.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        elements.append(table)

def analyze_files_by_filename(input_folder):
    """Analyze files by extracting brand names from filenames"""
    dateien = []
    marken_set = set()
    renamed_files_by_folder_and_marke = defaultdict(lambda: defaultdict(list))

    for subfolder in sorted(Path(input_folder).iterdir()):
        if not subfolder.is_dir():
            continue

        for file in sorted(subfolder.iterdir()):
            if file.suffix.lower() not in ALL_ALLOWED_EXTENSIONS:
                continue
            marke = extract_brand(file.stem)
            marken_set.add(marke)
            dateien.append({
                "original_file": file,
                "original_folder": subfolder.name,
                "marke": marke
            })

    for eintrag in dateien:
        orig = eintrag["original_file"]
        marke = eintrag["marke"]
        folder_name = eintrag["original_folder"]
        renamed_files_by_folder_and_marke[folder_name][marke].append((orig, orig.name))

    return marken_set, renamed_files_by_folder_and_marke, dateien

def generate_pdf_report(input_folder, erste_marke=None):
    """Generate the complete PDF report (original method)"""
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

def generate_filename_based_pdf_report(input_folder, erste_marke=None):
    """Generate PDF report based on filename brand analysis"""
    marken_set, renamed_files_by_folder_and_marke, all_files = analyze_files_by_filename(input_folder)
    
    if not marken_set:
        return None, "No brands found in filenames."

    # Create brand index
    marken_index = {}
    if erste_marke and erste_marke in marken_set:
        marken_index[erste_marke] = "01"
        aktuelle_nummer = 2
        for marke in sorted(marken_set):
            if marke != erste_marke:
                marken_index[marke] = f"{aktuelle_nummer:02d}"
                aktuelle_nummer += 1
    else:
        for i, marke in enumerate(sorted(marken_set), 1):
            marken_index[marke] = f"{i:02d}"

    # Generate PDF
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, leftMargin=20, rightMargin=20, topMargin=40, bottomMargin=30)
    elements = []
    styles = getSampleStyleSheet()

    nummer_zu_marke = {v: k for k, v in marken_index.items()}
    marken_spalten = sorted(nummer_zu_marke.items())

    # Overview by folder
    for folder in sorted(renamed_files_by_folder_and_marke):
        elements.append(Paragraph(f"<b>Folder: {folder}</b>", styles['Heading2']))

        abschnitt = renamed_files_by_folder_and_marke[folder]
        total = 0
        lines = []
        for nummer, marke in marken_spalten:
            count = len(abschnitt.get(marke, []))
            total += count
            lines.append(f"{nummer} ({marke}): {count}")
        lines.append(f"Total: {total}")
        elements.append(Paragraph("<br/>".join(lines), styles['Normal']))
        elements.append(Spacer(1, 10))

        headers = [f"{nummer} ({marke})" for nummer, marke in marken_spalten]
        data = [headers]
        col_data = []
        max_rows = 0
        for _, marke in marken_spalten:
            eintraege = abschnitt.get(marke, [])
            zellen = [get_asset_cell(p, n, len(headers)) for p, n in eintraege]
            col_data.append(zellen)
            max_rows = max(max_rows, len(zellen))

        for i in range(max_rows):
            row = []
            for col in col_data:
                row.append(col[i] if i < len(col) else "")
            data.append(row)

        col_width = (A4[0] - 40) / len(headers)
        t = Table(data, colWidths=[col_width] * len(headers))
        t.setStyle(TableStyle([
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        elements.append(t)
        elements.append(PageBreak())

    # Total overview
    elements.append(Paragraph("<b>Total Overview</b>", styles['Heading2']))
    global_counts = defaultdict(int)
    for folder_data in renamed_files_by_folder_and_marke.values():
        for marke, daten in folder_data.items():
            global_counts[marke] += len(daten)
    
    gesamt = 0
    summary = []
    for nummer, marke in marken_spalten:
        count = global_counts[marke]
        summary.append(f"{marke}: {count} assets")
        gesamt += count
    summary.append(f"Total number of all assets: {gesamt}")
    elements.append(Paragraph("<br/>".join(summary), styles['Normal']))

    # Add brand pages
    add_brand_pages(elements, marken_spalten, renamed_files_by_folder_and_marke)

    try:
        doc.build(elements)
        pdf_buffer.seek(0)
        return pdf_buffer, None
    except Exception as e:
        return None, f"Error generating PDF: {str(e)}"

def pdf_generator_tool():
    st.header("Brand Assets PDF Generator")
    st.markdown("Upload a zip file containing brand assets to generate a comprehensive PDF report.")
    
    # Analysis method selection
    analysis_method = st.radio(
        "Choose analysis method:",
        ["Folder-based analysis", "Filename-based analysis"],
        help="Folder-based: analyzes brands based on folder structure. Filename-based: extracts brand names from filenames."
    )
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a ZIP file",
        type=['zip'],
        help="Upload a zip file containing your brand assets.",
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
            
            if analysis_method == "Folder-based analysis":
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
                # Filename-based analysis
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
                        pdf_buffer, error = generate_filename_based_pdf_report(input_folder, erste_marke)
                    
                    if error:
                        st.error(f"{error}")
                    elif pdf_buffer:
                        st.success("PDF report generated successfully!")
                        
                        # Download button
                        st.download_button(
                            label="📥 Download PDF Report",
                            data=pdf_buffer.getvalue(),
                            file_name="IcAt_Overview_Branding.pdf",
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
            
            **Folder-based analysis:**
            - Create folders named after your brands
            - Place brand assets inside each brand folder
            - Zip the entire structure
            
            **Filename-based analysis:**
            - Name your files with brand identifiers at the beginning (e.g., 'Brand1_asset.jpg', 'Brand2_document.pdf')
            - Organize files in any folder structure
            - The tool will extract brand names from filenames automatically
            
            **Supported file types:**
            - **Images:** JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP
            - **Text files:** TXT, MD, CSV
            """)

def excel_extractor_tool():
    st.header("Excel Image Extractor")
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
        if st.button("Extract Images", type="primary", key="extract_btn"):
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
                    st.error(f"Error processing {uploaded_file.name}: {error}")
                elif count > 0:
                    st.success(f"Extracted {count} image(s) from {uploaded_file.name}")
                    all_extracted_files.extend(extracted_files)
                    total_images += count
                else:
                    st.warning(f"No images found in {uploaded_file.name}")
            
            status_text.empty()
            progress_bar.empty()
            
            if total_images > 0:
                st.success(f"Total: {total_images} images extracted from {len(uploaded_files)} file(s)")
                
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
                    file_name="extracted_images.zip",
                    mime="application/zip"
                )
                
                # Show preview of extracted images
                with st.expander("Preview Extracted Images"):
                    cols = st.columns(3)
                    col_idx = 0
                    
                    for file_path in all_extracted_files[:12]:  # Show max 12 images
                        try:
                            with cols[col_idx % 3]:
                                img = PILImage.open(file_path)
                                st.image(img, caption=os.path.basename(file_path), use_container_width =True)
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

def file_renamer_tool():
    st.header("Brand File Renamer")
    st.markdown("Upload a zip file containing brand folders to automatically rename files with brand prefixes.")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a ZIP file",
        type=['zip'],
        help="Upload a zip file containing your brand folders with assets.",
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
            
            if total_files > 0:
                st.success(f"Successfully renamed and copied {total_files} files")
                
                # Show preview
                with st.expander("📋 Preview renamed files"):
                    preview_df = pd.DataFrame(
                        processed_files[:20],  # Show first 20 files
                        columns=["Original Name", "New Name", "Brand"]
                    )
                    st.dataframe(preview_df)
                    
                    if total_files > 20:
                        st.info(f"Showing first 20 files. Total processed: {total_files}")
                
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
            - Images: JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP
            - Text files: TXT, MD, CSV
            """)

def main():
    st.title("🔧 Brand Assets Tools")

    tab1, tab2, tab3 = st.tabs(["PDF Report Generator", "Excel Image Extractor", "Brand File Renamer"])

    with tab1:
        pdf_generator_tool()
    with tab2:
        excel_extractor_tool()
    with tab3:
        file_renamer_tool()

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
                        processed_path = os.path.join(temp_dir, "processed_" + new_filename)
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
                            st.image(img, caption=os.path.basename(file_path), use_container_width=True)
                    
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

def find_smallest_dimensions_tool():
    st.header("Find Smallest Image Dimensions")
    st.markdown("Analyze a folder to find images with the smallest width and height.")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a ZIP file with images",
        type=['zip'],
        help="Upload a zip file containing images to analyze.",
        key="dimensions_zip_upload"
    )
    
    if uploaded_file:
        try:
            # Extract zip file
            with st.spinner("Extracting zip file..."):
                temp_dir = extract_zip_to_temp(uploaded_file)
            
            # Find the actual input folder (in case zip has a root folder)
            input_folder = temp_dir
            items = os.listdir(temp_dir)
            if len(items) == 1 and os.path.isdir(os.path.join(temp_dir, items[0])):
                input_folder = os.path.join(temp_dir, items[0])
            
            # Analyze images
            with st.spinner("Analyzing images..."):
                min_width = None
                min_height = None
                file_min_width = ""
                file_min_height = ""
                height_at_min_width = None
                width_at_min_height = None
                total_images = 0
                
                for root, dirs, files in os.walk(input_folder):
                    for file in files:
                        ext = os.path.splitext(file)[1].lower()
                        if ext in ALLOWED_IMAGE_EXTENSIONS:
                            path = os.path.join(root, file)
                            try:
                                with PILImage.open(path) as img:
                                    width, height = img.size
                                    total_images += 1
                                    
                                    if (min_width is None) or (width < min_width):
                                        min_width = width
                                        height_at_min_width = height
                                        file_min_width = path
                                    
                                    if (min_height is None) or (height < min_height):
                                        min_height = height
                                        width_at_min_height = width
                                        file_min_height = path
                            except Exception as e:
                                st.warning(f"Error processing {file}: {str(e)}")
                
                if min_width is not None and min_height is not None:
                    st.success(f"Analyzed {total_images} images")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.subheader("Smallest Width")
                        st.write(f"Dimensions: {min_width} × {height_at_min_width} px")
                        st.image(file_min_width, caption=os.path.basename(file_min_width), use_container_width =True)
                    
                    with col2:
                        st.subheader("Smallest Height")
                        st.write(f"Dimensions: {width_at_min_height} × {min_height} px")
                        st.image(file_min_height, caption=os.path.basename(file_min_height), use_container_width =True)
                else:
                    st.warning("No valid image files found in the uploaded zip.")
        
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    
    else:
        st.info("👆 Please upload a zip file with images to analyze.")
        
        # Instructions
        with st.expander("📖 Instructions"):
            st.markdown("""
            **How to use this tool:**
            
            1. Upload a zip file containing images
            2. The tool will analyze all images in the folder and subfolders
            3. It will find and display:
               - The image with the smallest width
               - The image with the smallest height
            
            **Supported image formats:**
            - JPG, JPEG, PNG, BMP, GIF, TIFF, WEBP
            """)

def resize_with_transparent_canvas_tool():
    st.header("Resize with Transparent Canvas")
    st.markdown("Resize images to a target dimension while maintaining aspect ratio on transparent background.")
    
    # File upload with unique key
    uploaded_file = st.file_uploader(
        "Choose a ZIP file with images",
        type=['zip'],
        help="Upload a zip file containing images to resize.",
        key="resize_transparent_upload"
    )
    
    if uploaded_file:
        # Options
        col1, col2 = st.columns(2)
        with col1:
            mode = st.radio(
                "Resize mode:",
                ["By height", "By width"],
                help="Resize all images to match either height or width while maintaining aspect ratio",
                key="resize_mode_selector"
            )
        with col2:
            # Set different limits based on mode
            if mode == "By height":
                target_value = st.number_input(
                    "Target height (pixels):",
                    min_value=10,       # Minimum 10px height
                    max_value=4000,     # Maximum 4000px height
                    value=1000,         # Default 1000px
                    step=10,
                    key="target_height_input"
                )
            else:  # By width
                target_value = st.number_input(
                    "Target width (pixels):",
                    min_value=10,       # Minimum 10px width
                    max_value=4000,     # Maximum 4000px width
                    value=1000,         # Default 1000px
                    step=10,
                    key="target_width_input"
                )
        
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
                            
                            if mode == "By height":
                                # Calculate new dimensions based on target height
                                factor = target_value / orig_h
                                new_h = target_value
                                new_w = int(orig_w * factor)
                            else:  # By width
                                # Calculate new dimensions based on target width
                                factor = target_value / orig_w
                                new_w = target_value
                                new_h = int(orig_h * factor)
                            
                            # Resize image while maintaining aspect ratio
                            img_resized = img.resize((new_w, new_h), PILImage.LANCZOS)
                            
                            # Create transparent canvas with target dimensions
                            if mode == "By height":
                                # For height mode, canvas matches the resized width and target height
                                canvas = PILImage.new("RGBA", (new_w, new_h), (0, 0, 0, 0))
                            else:
                                # For width mode, canvas matches the target width and resized height
                                canvas = PILImage.new("RGBA", (new_w, new_h), (0, 0, 0, 0))
                            
                            # Center the resized image on the canvas
                            x_offset = (canvas.width - img_resized.width) // 2
                            y_offset = (canvas.height - img_resized.height) // 2
                            canvas.paste(img_resized, (x_offset, y_offset), mask=img_resized)
                            
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
                        file_name="resized_transparent_images.zip",
                        mime="application/zip",
                        key="download_resized_transparent"
                    )
                    
                    # Show preview (removed key from expander)
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
                                    st.image(img, caption=os.path.basename(file_path), use_container_width=True)
                        
                        if processed_count > 6:
                            st.info(f"Showing sample images. Total processed: {processed_count}")
            
            except Exception as e:
                st.error(f"An error occurred during processing: {str(e)}")
    
    else:
        st.info("👆 Please upload a zip file with images to resize.")
        
        # Instructions (removed key from expander)
        with st.expander("📖 Instructions"):
            st.markdown("""
            **How to use this tool:**
            
            1. Upload a zip file containing images
            2. Choose resize mode:
               - By height: All images will be resized to match the target height
               - By width: All images will be resized to match the target width
            3. Set the target dimension (10-4000 pixels)
            4. Click "Process Images"
            5. Download the resized images with transparent backgrounds
            
            **Features:**
            - Maintains original aspect ratio
            - Centers images on transparent canvas
            - Outputs PNG files to preserve transparency
            - Preserves folder structure
            
            **Limits:**
            - Minimum dimension: 10 pixels
            - Maximum dimension: 4000 pixels
            - For batch processing of many images, please be patient
            """)
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
            max_value=2000,
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
                                st.image(img, caption=os.path.basename(file_path), use_container_width=True)
                        
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

def main():
    st.title("🔧 Brand Assets Tools")

    tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
        "PDF Report Generator", 
        "Excel Image Extractor", 
        "Brand File Renamer",
        "White to Transparent",
        "Find Smallest Dimensions",
        "Resize with Transparent Canvas",
        "Center on Canvas"
    ])

    with tab1:
        pdf_generator_tool()
    with tab2:
        excel_extractor_tool()
    with tab3:
        file_renamer_tool()
    with tab4:
        white_to_transparent_tool()
    with tab5:
        find_smallest_dimensions_tool()
    with tab6:
        resize_with_transparent_canvas_tool()
    with tab7:
        center_on_canvas_tool()

if __name__ == "__main__":
    main()
