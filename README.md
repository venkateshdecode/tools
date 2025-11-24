---
title: Brand Asset Management Tools
emoji: 🔧
colorFrom: blue
colorTo: green
sdk: streamlit
sdk_version: 1.32.2
app_file: app.py
pinned: false
license: mit
python_version: 3.12
---

# 🔧 Brand Asset Management Tools

A comprehensive suite of tools for managing and processing brand assets, built with Streamlit.

## Features

This application provides 7 powerful tools for brand asset management:

1. **Extract Images from Excel** - Extract embedded images from Excel worksheets (.xlsx files)
2. **Name Assets by Brand** - Rename files based on brand naming conventions
3. **Canvas White to Transparent** - Convert white backgrounds to transparent in images
4. **Asset Overview** - Generate PDF overview of your assets
5. **Name Stims by Block** - Organize and rename stimulus files by block with output file creation
6. **Resize** - Resize images while maintaining aspect ratio on transparent canvas
7. **Place on Canvas** - Center images on a custom canvas size

## Usage

Simply select the tool you need from the tabs at the top of the application and follow the instructions for each tool.

### Extract Images from Excel
- Upload one or more Excel files (.xlsx)
- Select specific worksheets or all worksheets
- Download extracted images as a ZIP file

### Name Assets by Brand
- Upload images and specify brand naming patterns
- Batch rename files according to your conventions

### Canvas White to Transparent
- Convert white backgrounds in images to transparent
- Useful for logo and asset preparation

### Asset Overview
- Generate comprehensive PDF reports of your asset collections

### Name Stims by Block
- Organize stimulus files by experimental blocks
- Create structured output files

### Resize
- Resize images to specific dimensions
- Maintain transparent backgrounds

### Place on Canvas
- Center images on custom canvas sizes
- Perfect for creating consistent asset formats

## Requirements

- Python 3.12
- Streamlit 1.32.2
- OpenCV (headless)
- Pillow
- ReportLab
- Pandas
- openpyxl
- xlrd

See [requirements.txt](requirements.txt) for complete dependencies.

## Local Development

To run this application locally:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## License

MIT License

## Author

Brand Asset Management Tools for research and marketing teams.