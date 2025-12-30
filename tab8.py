import io
import os
import re
import shutil
import zipfile
from pathlib import Path
from collections import defaultdict
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
    # Try to match brand name as letters only first, then extract count_str
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


def reconstruct_folder_structure(files_folder):
    """
    Reconstruct the folder structure from parsed filenames.
    Groups files by blocknummer (folder) and marke (brand).

    Returns:
    - marken_set: set of all brands
    - renamed_files_by_folder_and_marke: nested dict structure
    - marken_index: mapping of brand to markennummer
    - file_to_factorgroup: mapping of filename to folder name
    """
    marken_set = set()
    marken_index = {}
    renamed_files_by_folder_and_marke = defaultdict(lambda: defaultdict(list))
    file_to_factorgroup = {}

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
            if 'Overview' in file_path.name or 'Report' in file_path.name:
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

        # Create folder name - use just blocknummer to match Tab 5 output format
        # The folder_name will be like "00Folder", "01Folder", etc.
        folder_name = f"{blocknummer}Folder"

        # Add to structure
        renamed_files_by_folder_and_marke[folder_name][marke].append(
            (file_path, file_path.name)
        )

        # Map file to factorgroup - this is what goes into Excel's factorgroup column
        file_to_factorgroup[file_path.name] = folder_name

    # Prepare debug info
    debug_info = {
        'total_files': total_files,
        'parsed_files': len(parsed_files),
        'skipped_files': skipped_files
    }

    return marken_set, renamed_files_by_folder_and_marke, marken_index, file_to_factorgroup, debug_info


def regenerate_reports_from_folder(files_folder):
    """
    Regenerate PDF and Excel reports from a folder containing Tab 5 output files.
    Uses the EXISTING Excel file to get the correct factorgroup mappings!
    """
    import pandas as pd

    # CRITICAL: Read the existing Excel file to get factorgroup mappings
    existing_excel = Path(files_folder) / "IcAt_Overview_Final.xlsx"

    if not existing_excel.exists():
        return None, None, "Error: IcAt_Overview_Final.xlsx not found in the folder. This file is required to regenerate reports."

    # Read existing Excel to get the file_to_factorgroup mapping
    try:
        df = pd.read_excel(existing_excel, sheet_name='Assets')
        print(f"✅ Excel has {len(df)} rows")

        # Build file_to_factorgroup from the Excel
        file_to_factorgroup = {}
        for idx, row in df.iterrows():
            lang_val = str(row['Language'])
            factorgroup = str(row['factorgroup'])

            # For normal files, Language contains the full filename with extension
            # For text files, Language contains the text content
            # For B20 files, Language contains the full filename with extension

            # Check if Language looks like a filename (has an extension)
            if '.' in lang_val and not lang_val.startswith('['):  # Not an error message
                filename = lang_val
                file_to_factorgroup[filename] = factorgroup

                if idx < 5:  # Debug: show first 5 mappings
                    print(f"  Row {idx}: '{filename}' -> {factorgroup}")
            elif idx < 5:
                # This is likely a text file - skip it for now
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
            for skipped in debug_info['skipped_files'][:10]:  # Show first 10
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
                if 'Overview' in file_path.name or 'Report' in file_path.name:
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
        old_excel_path = Path(files_folder) / "IcAt_Overview_Final.xlsx"
        backup_excel_path = Path(files_folder) / "IcAt_Overview_Final.xlsx.backup"
        if old_excel_path.exists():
            # Create backup for comparison
            shutil.copy2(old_excel_path, backup_excel_path)
            # Remove old file to generate fresh
            old_excel_path.unlink()

        # Generate Excel using existing helper function
        # Note: generate_excel_report will scan ALL files in files_folder
        # including any newly added images
        print(f"\n📝 Generating Excel report...")
        print(f"  Passing {len(file_to_factorgroup)} file mappings to generate_excel_report")

        # Debug: Show a few sample mappings
        sample_count = 0
        for filename, factorgroup in file_to_factorgroup.items():
            if sample_count < 3:
                print(f"    Sample: '{filename}' -> '{factorgroup}'")
                sample_count += 1

        excel_path = generate_excel_report(
            Path(files_folder),
            marken_index,
            file_to_factorgroup
        )

        # Debug: Check how many entries were added
        import pandas as pd
        try:
            df = pd.read_excel(excel_path, sheet_name='Assets')
            print(f"\n✅ Excel generated with {len(df)} rows")

            # Check specifically for the user's new file
            user_file = '01B01hersheys06packing.png'
            if user_file in df['Language'].values:
                row = df[df['Language'] == user_file].iloc[0]
                print(f"  ✅ Found user's new file '{user_file}':")
                print(f"     factor: {row['factor']}")
                print(f"     factorgroup: {row['factorgroup']}")
            else:
                print(f"  ❌ User's new file '{user_file}' NOT FOUND in Excel!")
                print(f"  Checking if file exists in folder...")
                file_path = Path(files_folder) / user_file
                if file_path.exists():
                    print(f"    ✅ File EXISTS in folder")
                else:
                    print(f"    ❌ File DOES NOT EXIST in folder")
        except Exception as e:
            print(f"Could not read Excel for debug: {e}")

        return pdf_buffer, excel_path, None

    finally:
        # Clean up temp folder
        if temp_input.exists():
            shutil.rmtree(temp_input)


def regenerate_reports_tool():
    """Tab 8: Regenerate PDF and Excel from Tab 5 output + manually added images"""
    st.header("Re Run Matrix")
    st.markdown("Upload a ZIP containing Tab 5 output folder with optionally added images. The tool will regenerate PDF and Excel reports.")

    # Initialize session state
    if 'regeneration_complete' not in st.session_state:
        st.session_state.regeneration_complete = False
    if 'regenerated_data' not in st.session_state:
        st.session_state.regenerated_data = {}

    uploaded_file = st.file_uploader(
        "Choose a ZIP file containing Tab 5 output",
        type=['zip'],
    )

    # Reset processing state when new file is uploaded
    if uploaded_file and 'uploaded_file_name_tab8' in st.session_state:
        if st.session_state.uploaded_file_name_tab8 != uploaded_file.name:
            st.session_state.regeneration_complete = False
            st.session_state.regenerated_data = {}

    if uploaded_file:
        st.session_state.uploaded_file_name_tab8 = uploaded_file.name

        if not st.session_state.regeneration_complete:
            if st.button("Regenerate Reports", type="primary"):
                with st.spinner("Extracting files and analyzing structure..."):
                    # Extract ZIP
                    temp_dir = extract_zip_to_temp(uploaded_file)

                    # Find the actual files folder - look for processed_files
                    root_folder = Path(temp_dir)

                    # Keep descending if there's only one subfolder (to handle extra nesting)
                    while True:
                        items = [item for item in os.listdir(root_folder) if not item.startswith('.') and not item.startswith('_')]

                        # If only one item and it's a directory, descend into it
                        if len(items) == 1 and os.path.isdir(os.path.join(root_folder, items[0])):
                            root_folder = root_folder / items[0]
                        else:
                            # Found the actual root level
                            break

                    # Look for processed_files folder
                    processed_files_folder = root_folder / "processed_files"
                    if not processed_files_folder.exists():
                        # Maybe files are directly in root
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

                            # Compare with old Excel
                            old_excel = processed_files_folder / "IcAt_Overview_Final.xlsx.backup"
                            if not old_excel.exists():
                                old_excel = processed_files_folder / "IcAt_Overview_Final.xlsx"

                            try:
                                df_old = pd.read_excel(old_excel, sheet_name='Assets')
                                old_count = len(df_old)
                                new_count = len(df_new)
                                added_count = new_count - old_count

                                if added_count > 0:
                                    st.success(f"✅ Excel generated with {new_count} entries ({added_count} NEW entries added!)")
                                else:
                                    st.success(f"✅ Excel generated with {new_count} entries (same as before)")
                            except:
                                st.success(f"✅ Excel generated with {len(df_new)} entries")

                        except Exception as e:
                            st.warning(f"Excel created but couldn't read for stats: {e}")

                    # Create individual processed files ZIP (just like tab5)
                    zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                        for file in processed_files_folder.iterdir():
                            if file.is_file():
                                # Skip old backups
                                if file.name.endswith('.backup'):
                                    continue
                                zip_file.write(file, file.name)
                    zip_buffer.seek(0)

                    # Create combined ZIP with processed_files/ and reports/ folders (just like tab5)
                    combined_zip_buffer = io.BytesIO()
                    with zipfile.ZipFile(combined_zip_buffer, 'w', zipfile.ZIP_DEFLATED) as combined_zip:
                        # Add all processed files to processed_files/ folder
                        for file in processed_files_folder.iterdir():
                            if file.is_file():
                                # Skip old backups
                                if file.name.endswith('.backup'):
                                    continue
                                combined_zip.write(file, f"processed_files/{file.name}")

                        # Add PDF report to reports/ folder
                        if pdf_buffer:
                            combined_zip.writestr("reports/Brand_Assets_Report.pdf", pdf_buffer.getvalue())

                        # Add Excel report to reports/ folder
                        if excel_path and excel_path.exists():
                            with open(excel_path, 'rb') as excel_file:
                                combined_zip.writestr("reports/Brand_Assets_Overview.xlsx", excel_file.read())

                    combined_zip_buffer.seek(0)

                    # Store in session state
                    st.session_state.regenerated_data = {
                        'pdf_buffer': pdf_buffer,
                        'excel_path': excel_path,
                        'zip_buffer': zip_buffer,
                        'combined_zip_buffer': combined_zip_buffer,
                        'temp_dir': temp_dir
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

            # Combined download option (recommended)
            st.subheader("📦 Complete Package Download")
            st.markdown("**Recommended:** Download everything in one convenient package")

            if st.session_state.regenerated_data.get('combined_zip_buffer'):
                st.download_button(
                    label="🎁 Download Complete Package (All Files + Reports)",
                    data=st.session_state.regenerated_data['combined_zip_buffer'].getvalue(),
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
                if st.session_state.regenerated_data.get('pdf_buffer'):
                    st.download_button(
                        label="📄 PDF Report",
                        data=st.session_state.regenerated_data['pdf_buffer'].getvalue(),
                        file_name="Brand_Assets_Report.pdf",
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
                                file_name="Brand_Assets_Overview.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    except:
                        st.info("Excel report not available")

            with col3:
                if st.session_state.regenerated_data.get('zip_buffer'):
                    st.download_button(
                        label="📦 Processed Files",
                        data=st.session_state.regenerated_data['zip_buffer'].getvalue(),
                        file_name="Brand_Assets_Processed.zip",
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
            **What is Tab 8?**

            Tab 8 regenerates PDF and Excel reports from a Tab 5 output folder that may contain manually added images.

            **Use Case:**
            1. You ran Tab 5 and got an output folder with images, PDF, and Excel
            2. You manually added new images to that folder (following the same naming convention)
            3. You want to regenerate the reports to include the new images

            **How to use:**
            1. Take your Tab 5 output folder (with any manually added images)
            2. ZIP the entire folder
            3. Upload the ZIP here
            4. Click "Regenerate Reports"
            5. Download the new PDF and Excel files

            **Naming Convention:**
            All files must follow this pattern: `{markennummer}B{blocknummer}{marke}{count_str}{cleaned}{ext}`

            Example: `01B01brand01filename.png`
            - `01` = brand number
            - `B01` = block/folder number
            - `brand` = brand name
            - `01` = sequential counter
            - `filename` = cleaned name
            - `.png` = file extension

            **Note:** The tool will skip any existing Excel/PDF reports in the folder and only regenerate new ones.
            """)
