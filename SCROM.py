import os
import zipfile
import time
from pptx import Presentation
import win32com.client


def close_powerpoint_instances():
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Quit()
    except Exception as e:
        pass  # Ignore if no PowerPoint instances are running


def convert_ppsx_to_pptx(ppsx_file, pptx_file):
    close_powerpoint_instances()
    time.sleep(1)  # Give some time to ensure PowerPoint instances are closed
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(ppsx_file, WithWindow=False)
    presentation.SaveAs(pptx_file, 24)  # 24 oznacza format pptx
    presentation.Close()
    powerpoint.Quit()


def export_slides_to_images(pptx_file, output_folder):
    close_powerpoint_instances()
    time.sleep(1)

    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    powerpoint.Visible = 1
    presentation = powerpoint.Presentations.Open(pptx_file, WithWindow=False)

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    for i, slide in enumerate(presentation.Slides):
        image_path = os.path.abspath(os.path.join(output_folder, f"slide_{i + 1}.png"))
        slide.Export(image_path, "PNG")

    presentation.Close()
    powerpoint.Quit()


def convert_pptx_to_scorm(pptx_file, output_folder):
    export_slides_to_images(pptx_file, output_folder)
    slide_images = [os.path.abspath(os.path.join(output_folder, f"slide_{i + 1}.png")) for i in
                    range(len(Presentation(pptx_file).slides))]
    generate_html_files(slide_images, output_folder)
    scorm_manifest = generate_scorm_manifest(slide_images)
    manifest_path = os.path.join(output_folder, 'imsmanifest.xml')
    with open(manifest_path, 'w') as f:
        f.write(scorm_manifest)

    scorm_zip = os.path.join(output_folder, 'scorm_package.zip')
    with zipfile.ZipFile(scorm_zip, 'w') as zipf:
        for image in slide_images:
            zipf.write(image, os.path.basename(image))
        zipf.write(manifest_path, 'imsmanifest.xml')
        for i in range(len(slide_images)):
            html_file = os.path.join(output_folder, f"slide_{i + 1}.html")
            zipf.write(html_file, os.path.basename(html_file))

    print(f"SCORM package created at {scorm_zip}")


def generate_html_files(slide_images, output_folder):
    for i, image in enumerate(slide_images):
        thumbnails = "".join([
                                 f'<div style="text-align: center; margin-bottom: 10px;"><a href="slide_{j + 1}.html"><img src="{os.path.basename(img)}" alt="Slide {j + 1}" style="width:100%; background-color: #ccc; {"border: 4px solid red;" if i == j else "border: 2px solid transparent;"}" id="thumb_{j + 1}"></a><br><span style="color: white;">{j + 1}</span></div>'
                                 for j, img in enumerate(slide_images)])
        next_slide = f"slide_{i + 2}.html" if i + 1 < len(slide_images) else "slide_1.html"
        prev_slide = f"slide_{i}.html" if i > 0 else f"slide_{len(slide_images)}.html"
        html_content = f'''
        <!DOCTYPE html>
        <html>
        <head>
            <title>Slide {i + 1}</title>
            <style>
                body {{
                    display: flex;
                    height: 100vh;
                    margin: 0;
                    background-color: #1a1a1a;
                    color: white;
                    font-family: Arial, sans-serif;
                }}
                .content {{
                    display: flex;
                    justify-content: center;
                    align-items: center;
                    flex-direction: column;
                    flex-grow: 1;
                }}
                img {{
                    max-width: 90%;
                    max-height: 80vh;
                }}
                .slide-description {{
                    margin-top: 10px;
                    font-size: 18px;
                }}
                .nav-buttons {{
                    margin-top: 20px;
                }}
                .nav-buttons button {{
                    padding: 10px 20px;
                    font-size: 16px;
                    margin: 0 10px;
                    cursor: pointer;
                    background-color: #444;
                    color: white;
                    border: none;
                    border-radius: 5px;
                    transition: background-color 0.3s ease;
                }}
                .nav-buttons button:hover {{
                    background-color: #1abc9c;
                }}
                .sidebar {{
                    width: 200px;
                    background-color: #333;
                    padding: 10px;
                    overflow-y: auto;
                    height: 100vh;
                }}
                .sidebar img {{
                    transition: border 0.3s ease, background-color 0.3s ease;
                    background-color: #ccc;
                }}
                .sidebar img:hover {{
                    border: 2px solid #1abc9c;
                }}
                .sidebar img.active {{
                    border: 4px solid red;
                    background-color: #666;
                }}
            </style>
            <script>
                window.onload = function() {{
                    var currentThumb = document.getElementById("thumb_{i + 1}");
                    if (currentThumb) {{
                        currentThumb.scrollIntoView({{behavior: "smooth", block: "center"}});
                        currentThumb.classList.add("active");
                    }}
                }};
                function navigateTo(url) {{
                    var currentThumb = document.querySelector(".sidebar img.active");
                    if (currentThumb) {{
                        currentThumb.classList.remove("active");
                    }}
                    window.location.href = url;
                }}
            </script>
        </head>
        <body>
            <div class="sidebar">
                {thumbnails}
            </div>
            <div class="content">
                <img src="{os.path.basename(image)}" alt="Slide {i + 1}">
                <div class="slide-description">Strona {i + 1} z {len(slide_images)}</div>
                <div class="nav-buttons">
                    <button onclick="navigateTo('{prev_slide}')">Wstecz</button>
                    <button onclick="navigateTo('{next_slide}')">Dalej</button>
                </div>
            </div>
        </body>
        </html>
        '''
        html_path = os.path.join(output_folder, f"slide_{i + 1}.html")
        with open(html_path, 'w') as f:
            f.write(html_content)


def generate_scorm_manifest(slide_images):
    manifest_template = '''<?xml version="1.0" encoding="UTF-8"?>
<manifest identifier="com.example.scorm" version="1.2">
  <metadata>
    <schema>ADL SCORM</schema>
    <schemaversion>1.2</schemaversion>
  </metadata>
  <organizations default="ORG">
    <organization identifier="ORG">
      <title>Sample SCORM Course</title>
      {items}
    </organization>
  </organizations>
  <resources>
    {resources}
  </resources>
</manifest>'''

    item_template = '<item identifier="ITEM-{index}" identifierref="RES-{index}"><title>Slide {index}</title></item>'
    resource_template = '<resource identifier="RES-{index}" type="webcontent" href="slide_{index}.html"><file href="slide_{index}.html"/><file href="slide_{index}.png"/></resource>'

    items = '\n'.join([item_template.format(index=i + 1) for i in range(len(slide_images))])
    resources = '\n'.join([resource_template.format(index=i + 1) for i in range(len(slide_images))])

    return manifest_template.format(items=items, resources=resources)


# Ścieżka do pliku PPSX
ppsx_file = r'C:\Users\kowal\OneDrive\Pulpit\CUI Prezentacja\eLearning_CUI_WrocławEDUEOFVer.ppsx'
# Ścieżka do zapisu pliku PPTX
pptx_file = r'C:\Users\kowal\OneDrive\Pulpit\CUI Prezentacja\eLearning_CUI_WrocławEDUEOFVer.pptx'

# Konwersja pliku PPSX na PPTX
convert_ppsx_to_pptx(ppsx_file, pptx_file)

# Generowanie pakietu SCORM z pliku PPTX
output_folder = 'output_scorm'
convert_pptx_to_scorm(pptx_file, output_folder)
