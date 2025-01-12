# SCORM Generator

This program converts PPSX presentation files to the SCORM format, which can be used in Learning Management Systems (LMS) like Moodle. The program converts PPSX files to PPTX, exports slides to images, generates HTML files for each slide, and creates a SCORM manifest (`imsmanifest.xml`). All these components are then packaged into a ZIP archive as a ready-to-use SCORM package.

## Features

1. **Convert PPSX to PPTX**  
   The program uses the `win32com.client` library to open a PPSX file in PowerPoint and save it as PPTX.

2. **Export Slides to Images**  
   Slides from the PPTX file are exported as PNG images.

3. **Generate HTML Files**  
   For each slide, an HTML file is created with navigation between slides and a thumbnail sidebar.

4. **Generate SCORM Manifest**  
   The program creates a SCORM manifest file (`imsmanifest.xml`) defining the course structure and relationships between its elements.

5. **Create SCORM Package**  
   All files (images, HTML, manifest) are packed into a ZIP archive as a SCORM package.

## Requirements

- Python 3.x
- Libraries:
  - `os`
  - `zipfile`
  - `time`
  - `pptx`
  - `win32com.client`
- Microsoft PowerPoint installed (required for PPSX to PPTX conversion).

## How to Use

1. Place the PPSX file in the appropriate directory and update the `ppsx_file` variable in the code to point to this file.
2. Specify the output path for the PPTX file by updating the `pptx_file` variable.
3. Run the program. The program will automatically:
   - Convert the PPSX file to PPTX.
   - Export slides to the `output_scorm` folder as PNG images.
   - Generate HTML files for each slide.
   - Create a SCORM manifest file (`imsmanifest.xml`).
   - Package all files into a ZIP archive as a SCORM package.
4. Upload the generated SCORM package to your LMS, such as Moodle.

## Output Folder Structure

After the program completes, the output folder (`output_scorm`) will contain:

- PNG files: images representing slides.
- HTML files: pages for each slide with navigation.
- `imsmanifest.xml`: SCORM manifest file.
- `scorm_package.zip`: ready-to-upload SCORM package.

## Key Functions

### `close_powerpoint_instances()`

Closes all running PowerPoint instances to avoid conflicts.

### `convert_ppsx_to_pptx(ppsx_file, pptx_file)`

Converts a PPSX file to PPTX using `win32com.client`.

### `export_slides_to_images(pptx_file, output_folder)`

Exports slides from a PPTX file as PNG images.

### `convert_pptx_to_scorm(pptx_file, output_folder)`

Generates the entire SCORM package:
- Calls `export_slides_to_images`.
- Creates HTML files for each slide.
- Generates the SCORM manifest file.
- Packages everything into a ZIP archive.

### `generate_html_files(slide_images, output_folder)`

Creates HTML files for each slide with navigation and thumbnails.

### `generate_scorm_manifest(slide_images)`

Generates the SCORM manifest file (`imsmanifest.xml`).

## Notes

- Ensure PowerPoint is properly installed on your computer.
- This program requires Windows to run because it uses `win32com.client`.

## Author

This program was created to convert presentations into SCORM packages for use in LMS platforms like Moodle. If you have any questions, feel free to reach out! "https://github.com/NorbertKowalski80"

