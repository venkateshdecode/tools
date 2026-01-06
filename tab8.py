import io
import os
import re
import shutil
import zipfile
from pathlib import Path
from collections import defaultdict
from datetime import datetime
import streamlit as st
from helper import (
    extract_zip_to_temp,
    generate_excel_report,
    generate_filename_based_pdf_report_with_extensions,
    ALL_ALLOWED_EXTENSIONS,
    is_valid_file
)


def parse_filename(filename):
    """
    Parse filename following Tab 5 naming convention:
    {markennummer}B{blocknummer}{marke}{count_str}{cleaned}{original_ext}

    Returns: (markennummer, blocknummer, marke, count_str, cleaned, ext) or None if invalid
    """
    stem = Path(filename).stem
    ext = Path(filename).suffix.lower()

    # Pattern: 2 digits + B + 2 digits + brand name (letters only) + 2 digits + rest
    match = re.match(r'^(\d{2})B(\d{2})([a-zA-Z]+)(\d{2})(.*)$', stem)

    if match:
        markennummer, blocknummer, marke, count_str, cleaned = match.groups()
        return {
            'markennummer': markennummer,
            'blocknummer': blocknummer,
            'marke': marke.lower(),
            'count_str': count_str,
            'cleaned': cleaned,
            'ext': ext,
            'full_filename': filename
        }

    # Fallback: Try more flexible pattern if brand contains numbers
    match = re.match(r'^(\d{2})B(\d{2})([a-z0-9]+?)(\d{2})(.*)$', stem, re.IGNORECASE)
    if match:
        markennummer, blocknummer, marke, count_str, cleaned = match.groups()
        return {
            'markennummer': markennummer,
            'blocknummer': blocknummer,
            'marke': marke.lower(),
            'count_str': count_str,
            'cleaned': cleaned,
            'ext': ext,
            'full_filename': filename
        }

    return None


def get_base_filename_without_ext(filename):
    """Get base filename without extension for duplicate detection"""
    return Path(filename).stem


def reconstruct_folder_structure(files_folder):
    """
    Reconstruct the folder structure from parsed filenames.
    Groups files by blocknummer (folder) and marke (brand).
    Handles duplicate base filenames with different extensions.
    """
    marken_set = set()
    marken_index = {}
    renamed_files_by_folder_and_marke = defaultdict(lambda: defaultdict(list))
    file_to_factorgroup = {}
    
    # Track unique base filenames for counting
    unique_base_files = defaultdict(set)  # folder -> set of base filenames

    # Collect all valid files
    parsed_files = []
    skipped_files = []
    total_files = 0

    for file_path in sorted(Path(files_folder).iterdir()):
        if not file_path.is_file():
            continue

        total_files += 1

        if not is_valid_file(file_path):
            skipped_files.append(f"{file_path.name} (invalid/system file)")
            continue

        # Skip Excel and PDF files from previous runs
        if file_path.suffix.lower() in ['.xlsx', '.pdf']:
            if any(keyword in file_path.name for keyword in ['Overview', 'Report', 'IcAt']):
                skipped_files.append(f"{file_path.name} (report file)")
                continue

        # Parse filename
        parsed = parse_filename(file_path.name)
        if parsed:
            parsed['file_path'] = file_path
            parsed_files.append(parsed)
        else:
            skipped_files.append(f"{file_path.name} (doesn't match naming convention)")

    # Build structure
    for parsed in parsed_files:
        marke = parsed['marke']
        markennummer = parsed['markennummer']
        blocknummer = parsed['blocknummer']
        file_path = parsed['file_path']

        # Add to marken_set
        marken_set.add(marke)

        # Add to marken_index (use first occurrence)
        if marke not in marken_index:
            marken_index[marke] = markennummer

        # Create folder name
        folder_name = f"{blocknummer}Folder"

        # Track unique base filename for this folder
        base_filename = get_base_filename_without_ext(file_path.name)
        unique_base_files[folder_name].add(base_filename)

        # Add to structure
        renamed_files_by_folder_and_marke[folder_name][marke].append(
            (file_path, file_path.name)
        )

        # Map file to factorgroup
        file_to_factorgroup[file_path.name] = folder_name

    # Prepare debug info
    debug_info = {
        'total_files': total_files,
        'parsed_files': len(parsed_files),
        'skipped_files': skipped_files,
        'unique_counts': {folder: len(bases) for folder, bases in unique_base_files.items()}
    }

    return marken_set, renamed_files_by_folder_and_marke, marken_index, file_to_factorgroup, debug_info


def find_existing_reports(files_folder):
    """Find existing PDF and Excel reports in the folder"""
    pdf_file = None
    excel_file = None
    
    for file in Path(files_folder).iterdir():
        if not file.is_file():
            continue
            
        # Look for Excel file
        if file.suffix.lower() == '.xlsx' and 'IcAt' in file.name and 'Overview' in file.name:
            excel_file = file
            
        # Look for PDF file
        if file.suffix.lower() == '.pdf' and ('Brand' in file.name or 'Report' in file.name or 'IcAt' in file.name):
            pdf_file = file
    
    return pdf_file, excel_file


def regenerate_reports_from_folder(files_folder):
    """
    Regenerate PDF and Excel reports from a folder containing Tab 5 output files.
    Uses the EXISTING Excel file to get the correct factorgroup mappings!
    Preserves manual folder structure changes from the Excel file.
    """
    import pandas as pd

    # Find existing reports (flexible naming)
    old_pdf_file, existing_excel = find_existing_reports(files_folder)
    
    if not existing_excel or not existing_excel.exists():
        return None, None, "Error: Excel overview file not found in the folder. Please ensure the Excel file from Tab 5 is included in the ZIP."

    # Read existing Excel to get the file_to_factorgroup mapping
    try:
        df = pd.read_excel(existing_excel, sheet_name='Assets')
        print(f"✅ Excel has {len(df)} rows")

        # Build file_to_factorgroup from the Excel
        file_to_factorgroup = {}
        for idx, row in df.iterrows():
            lang_val = str(row['Language'])
            factorgroup = str(row['factorgroup'])

            # Check if Language looks like a filename (has an extension)
            if '.' in lang_val and not lang_val.startswith('['):
                filename = lang_val
                file_to_factorgroup[filename] = factorgroup

                if idx < 5:
                    print(f"  Row {idx}: '{filename}' -> {factorgroup}")
            elif idx < 5:
                print(f"  Row {idx}: Skipping (text content): '{lang_val[:30]}...'")

        print(f"✅ Loaded {len(file_to_factorgroup)} file mappings from existing Excel")
    except Exception as e:
        return None, None, f"Error reading existing Excel file: {str(e)}"

    # Reconstruct structure from filenames
    marken_set, renamed_files_by_folder_and_marke, marken_index, _, debug_info = reconstruct_folder_structure(files_folder)

    if not marken_set:
        error_msg = f"No valid files found matching the naming convention.\n\n"
        error_msg += f"Total files found: {debug_info['total_files']}\n"
        error_msg += f"Successfully parsed: {debug_info['parsed_files']}\n"
        if debug_info['skipped_files']:
            error_msg += f"\nSkipped files:\n"
            for skipped in debug_info['skipped_files'][:10]:
                error_msg += f"  - {skipped}\n"
            if len(debug_info['skipped_files']) > 10:
                error_msg += f"  ... and {len(debug_info['skipped_files']) - 10} more\n"
        return None, None, error_msg

    # Add any NEW files to file_to_factorgroup
    new_files_added = 0
    print(f"\n🔍 Scanning folder for NEW files not in existing Excel...")
    for file_path in Path(files_folder).iterdir():
        if file_path.is_file() and file_path.name not in file_to_factorgroup:
            # Skip Excel/PDF files
            if file_path.suffix.lower() in ['.xlsx', '.pdf']:
                if any(keyword in file_path.name for keyword in ['Overview', 'Report', 'IcAt']):
                    continue

            # Parse to get blocknummer for new files
            parsed = parse_filename(file_path.name)
            if parsed:
                folder_name = f"{parsed['blocknummer']}Folder"
                file_to_factorgroup[file_path.name] = folder_name
                new_files_added += 1
                print(f"  ✨ NEW: {file_path.name} -> {folder_name}")

    if new_files_added == 0:
        print(f"  ℹ️  No new files found - all files already in Excel")

    print(f"\n📊 Summary:")
    print(f"  - Files from existing Excel: {len(file_to_factorgroup) - new_files_added}")
    print(f"  - NEW files added: {new_files_added}")
    print(f"  - Total files to process: {len(file_to_factorgroup)}")

    # Create a temporary input folder structure that matches Tab 5 expectations
    temp_input = Path(files_folder) / "_temp_input_structure"
    temp_input.mkdir(exist_ok=True)

    try:
        # Create subfolder structure for PDF generation
        for folder_name in renamed_files_by_folder_and_marke:
            folder_path = temp_input / folder_name
            folder_path.mkdir(exist_ok=True)

            # Copy files to appropriate subfolders
            for marke, files in renamed_files_by_folder_and_marke[folder_name].items():
                for file_path, _ in files:
                    shutil.copy2(file_path, folder_path / file_path.name)

        # Generate PDF using existing helper function
        pdf_buffer, pdf_error = generate_filename_based_pdf_report_with_extensions(
            temp_input,
            erste_marke=None,
            marken_index=marken_index
        )

        if pdf_error:
            return None, None, f"PDF generation failed: {pdf_error}"

        # Backup the old Excel file before deleting
        backup_excel_path = Path(files_folder) / f"{existing_excel.stem}.backup"
        if existing_excel.exists():
            shutil.copy2(existing_excel, backup_excel_path)
            existing_excel.unlink()

        # Generate Excel using existing helper function
        print(f"\n📝 Generating Excel report...")
        print(f"  Passing {len(file_to_factorgroup)} file mappings to generate_excel_report")

        excel_path = generate_excel_report(
            Path(files_folder),
            marken_index,
            file_to_factorgroup
        )

        # Debug: Check how many entries were added
        try:
            df_new = pd.read_excel(excel_path, sheet_name='Assets')
            print(f"\n✅ Excel generated with {len(df_new)} rows")
        except Exception as e:
            print(f"Could not read Excel for debug: {e}")

        return pdf_buffer, excel_path, None

    finally:
        # Clean up temp folder
        if temp_input.exists():
            shutil.rmtree(temp_input)


def regenerate_reports_tool():
    """Tab 6: Regenerate PDF and Excel from Tab 5 output + manually added images"""
    st.header("Re Run Matrix")
    st.markdown("""
    Upload a ZIP containing your Tab 5 output folder (with optional manually added images).

    The tool will regenerate updated reports with the current date stamp.
    """)

    # Initialize session state
    if 'regeneration_complete' not in st.session_state:
        st.session_state.regeneration_complete = False
    if 'regenerated_data' not in st.session_state:
        st.session_state.regenerated_data = {}

    uploaded_file = st.file_uploader(
        "Choose a ZIP file containing Tab 5 output (including both PDF and Excel reports)",
        type=['zip'],
    )

    # Reset processing state when new file is uploaded
    if uploaded_file and 'uploaded_file_name_tab6' in st.session_state:
        if st.session_state.uploaded_file_name_tab6 != uploaded_file.name:
            st.session_state.regeneration_complete = False
            st.session_state.regenerated_data = {}

    if uploaded_file:
        st.session_state.uploaded_file_name_tab6 = uploaded_file.name

        if not st.session_state.regeneration_complete:
            if st.button("Regenerate Reports", type="primary"):
                with st.spinner("Extracting files and analyzing structure..."):
                    # Extract ZIP
                    temp_dir = extract_zip_to_temp(uploaded_file)

                    # Find the actual files folder
                    root_folder = Path(temp_dir)

                    # Keep descending if there's only one subfolder
                    while True:
                        items = [item for item in os.listdir(root_folder) if not item.startswith('.') and not item.startswith('_')]

                        if len(items) == 1 and os.path.isdir(os.path.join(root_folder, items[0])):
                            root_folder = root_folder / items[0]
                        else:
                            break

                    # Look for processed_files folder
                    processed_files_folder = root_folder / "processed_files"
                    if not processed_files_folder.exists():
                        processed_files_folder = root_folder

                    # Regenerate reports
                    with st.spinner("Regenerating PDF and Excel reports..."):
                        pdf_buffer, excel_path, error = regenerate_reports_from_folder(processed_files_folder)

                    if error:
                        st.error(error)
                        shutil.rmtree(temp_dir)
                        return

                    # Show Excel statistics
                    if excel_path and excel_path.exists():
                        import pandas as pd
                        try:
                            df_new = pd.read_excel(excel_path, sheet_name='Assets')

                            # Compare with backup
                            backup_files = list(processed_files_folder.glob("*.backup"))
                            if backup_files:
                                try:
                                    df_old = pd.read_excel(backup_files[0], sheet_name='Assets')
                                    old_count = len(df_old)
                                    new_count = len(df_new)
                                    added_count = new_count - old_count

                                    if added_count > 0:
                                        st.success(f"✅ Excel generated with {new_count} entries ({added_count} NEW entries added!)")
                                    else:
                                        st.success(f"✅ Excel generated with {new_count} entries (same as before)")
                                except:
                                    st.success(f"✅ Excel generated with {len(df_new)} entries")
                            else:
                                st.success(f"✅ Excel generated with {len(df_new)} entries")

                        except Exception as e:
                            st.warning(f"Excel created but couldn't read for stats: {e}")

                    # Generate timestamp for filenames
                    timestamp = datetime.now().strftime("%Y%m%d")

                    # Create individual processed files ZIP
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file in processed_files_folder.iterdir():
                            if file.is_file():
                                if file.name.endswith('.backup'):
                                    continue
                                zip_file.write(file, file.name)
                    zip_buffer.seek(0)

                    # Create combined ZIP with processed_files/ and reports/ folders
                    combined_zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(combined_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as combined_zip:
                        # Add all processed files
                        for file in processed_files_folder.iterdir():
                            if file.is_file():
                                if file.name.endswith('.backup'):
                                    continue
                                combined_zip.write(file, f"processed_files/{file.name}")

                        # Add PDF report with timestamp
                        if pdf_buffer:
                            combined_zip.writestr(f"reports/Brand_Assets_Report_update_{timestamp}.pdf", pdf_buffer.getvalue())

                        # Add Excel report with timestamp
                        if excel_path and excel_path.exists():
                            with open(excel_path, 'rb') as excel_file:
                                combined_zip.writestr(f"reports/Brand_Assets_Overview_update_{timestamp}.xlsx", excel_file.read())

                    combined_zip_buffer.seek(0)

                    # Store in session state
                    st.session_state.regenerated_data = {
                        'pdf_buffer': pdf_buffer,
                        'excel_path': excel_path,
                        'zip_buffer': zip_buffer,
                        'combined_zip_buffer': combined_zip_buffer,
                        'temp_dir': temp_dir,
                        'timestamp': timestamp
                    }
                    st.session_state.regeneration_complete = True

                    st.rerun()

        # Show download options if processing is complete
        if st.session_state.regeneration_complete:
            st.success("✅ Reports regenerated successfully!")

            # Option to start over
            if st.button("🔄 Process New File", help="Clear results and upload a new file"):
                st.session_state.regeneration_complete = False
                st.session_state.regenerated_data = {}
                if 'temp_dir' in st.session_state.regenerated_data:
                    try:
                        shutil.rmtree(st.session_state.regenerated_data['temp_dir'])
                    except:
                        pass
                st.rerun()

            st.markdown("---")

            timestamp = st.session_state.regenerated_data.get('timestamp', datetime.now().strftime("%Y%m%d"))

            # Combined download option (recommended)
            st.subheader("📦 Complete Package Download")
            st.markdown("**Recommended:** Download everything in one convenient package")

            if st.session_state.regenerated_data.get('combined_zip_buffer'):
                st.download_button(
                    label="🎁 Download Complete Package (All Files + Reports)",
                    data=st.session_state.regenerated_data['combined_zip_buffer'].getvalue(),
                    file_name=f"Brand_Assets_Complete_Package_{timestamp}.zip",
                    mime="application/zip",
                    type="primary"
                )

            st.markdown("---")

            # Individual download options
            st.subheader("📁 Individual Downloads")
            st.markdown("Or download items separately:")

            col1, col2, col3 = st.columns(3)

            with col1:
                if st.session_state.regenerated_data.get('pdf_buffer'):
                    st.download_button(
                        label="📄 PDF Report",
                        data=st.session_state.regenerated_data['pdf_buffer'].getvalue(),
                        file_name=f"Brand_Assets_Report_update_{timestamp}.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.info("PDF report not available")

            with col2:
                if st.session_state.regenerated_data.get('excel_path'):
                    try:
                        with open(st.session_state.regenerated_data['excel_path'], 'rb') as excel_file:
                            st.download_button(
                                label="📊 Excel Report",
                                data=excel_file.read(),
                                file_name=f"Brand_Assets_Overview_update_{timestamp}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except:
                        st.info("Excel report not available")

            with col3:
                if st.session_state.regenerated_data.get('zip_buffer'):
                    st.download_button(
                        label="📦 Processed Files",
                        data=st.session_state.regenerated_data['zip_buffer'].getvalue(),
                        file_name=f"Brand_Assets_Processed_{timestamp}.zip",
                        mime="application/zip"
                    )

            st.markdown("---")
            st.info("💡 **Tip:** Downloads will remain available until you upload a new file or click 'Process New File'")

    else:
        # Reset session state when no file is uploaded
        if st.session_state.regeneration_complete:
            st.session_state.regeneration_complete = False
            st.session_state.regenerated_data = {}

        st.info("👆 Please upload a zip file to get started.")

        with st.expander("📖 Instructions"):
            st.markdown("""
            **What is Tab 6 (Re Run Matrix)?**

            This tool regenerates PDF and Excel reports from a Tab 5 output folder that may contain manually added images.

            **Use Case:**
            1. You ran Tab 5 and got an output folder with images, PDF, and Excel reports
            2. You manually added new images to that folder (following the same naming convention)
            3. You want to regenerate the reports to include the new images

            **How to use:**
            1. Take your Tab 5 output folder (the one with processed_files/ and reports/ folders)
            2. **Important:** Make sure BOTH the PDF and Excel reports are included in the ZIP
            3. Add any new images you want (following the naming convention)
            4. ZIP the entire folder
            5. Upload the ZIP here
            6. Click "Regenerate Reports"
            7. Download the updated reports (with today's date in the filename)

            **Naming Convention for New Files:**
            All files must follow this pattern: `{markennummer}B{blocknummer}{marke}{count_str}{cleaned}{ext}`

            Example: `01B01brand01filename.png`
            - `01` = brand number (markennummer)
            - `B01` = block/folder number
            - `brand` = brand name
            - `01` = sequential counter
            - `filename` = cleaned name
            - `.png` = file extension

            **Important Notes:**
            - The tool preserves any manual folder structure changes you made in the Excel file
            - Files with the same name but different extensions (e.g., .mp4 and .png) are counted only once
            - Both formats will appear in the reports, but the count remains accurate
            - Report files are automatically named with the current date (YYYYMMDD format)
            - The tool will skip any existing report files when processing

            **Multiple File Formats:**
            If you have files like:
            - `01B17nivea47appvideo.mp4`
            - `01B17nivea47appvideo.png`

            Both will be included in the reports, but counted as 1 asset (not 2) in the summary statistics.
            """)