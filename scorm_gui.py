import os
import zipfile
import time
from pptx import Presentation
import win32com.client
import tkinter as tk
from tkinter import filedialog, messagebox, Label, Button, Canvas
from PIL import Image, ImageTk
from convert_ppsx_to_pptx import convert_ppsx_to_pptx, close_powerpoint_instances  # Import funkcji z drugiego pliku


# Reszta kodu bez zmian...

def export_slides_to_images(pptx_file, output_folder):
    close_powerpoint_instances()
    time.sleep(5)
    try:
        if not os.path.exists(pptx_file):
            raise FileNotFoundError(f"File {pptx_file} not found.")
        if not os.access(pptx_file, os.R_OK):
            raise PermissionError(f"File {pptx_file} is not readable.")
        if not os.access(output_folder, os.W_OK):
            raise PermissionError(f"Folder {output_folder} is not writable.")

        print(f"Starting PowerPoint application...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        print(f"Opening PPTX file: {pptx_file}")
        presentation = powerpoint.Presentations.Open(pptx_file, WithWindow=False)

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        for i, slide in enumerate(presentation.Slides):
            image_path = os.path.abspath(os.path.join(output_folder, f"slide_{i + 1}.png"))
            slide.Export(image_path, "PNG")
            print(f"Exported slide {i + 1} to {image_path}")

        presentation.Close()
        powerpoint.Quit()
    except Exception as e:
        print(f"Error exporting slides: {e}")


def convert_pptx_to_scorm(pptx_file, output_folder):
    try:
        export_slides_to_images(pptx_file, output_folder)
        slide_images = [os.path.abspath(os.path.join(output_folder, f"slide_{i + 1}.png")) for i in
                        range(len(Presentation(pptx_file).slides))]
        generate_html_files(slide_images, output_folder)
        scorm_manifest = generate_scorm_manifest(slide_images)
        manifest_path = os.path.join(output_folder, 'imsmanifest.xml')
        with open(manifest_path, 'w') as f:
            f.write(scorm_manifest)
        print(f"Generated SCORM manifest at {manifest_path}")

        scorm_zip = os.path.join(output_folder, 'scorm_package.zip')
        with zipfile.ZipFile(scorm_zip, 'w') as zipf:
            for image in slide_images:
                if not os.path.exists(image):
                    raise FileNotFoundError(f"Image file {image} not found.")
                zipf.write(image, os.path.basename(image))
            zipf.write(manifest_path, 'imsmanifest.xml')
            for i in range(len(slide_images)):
                html_file = os.path.join(output_folder, f"slide_{i + 1}.html")
                if not os.path.exists(html_file):
                    raise FileNotFoundError(f"HTML file {html_file} not found.")
                zipf.write(html_file, os.path.basename(html_file))
        print(f"Created SCORM package at {scorm_zip}")

        messagebox.showinfo("Sukces", f"Pakiet SCORM został utworzony w {scorm_zip}")
    except Exception as e:
        messagebox.showerror("Błąd", f"Błąd podczas konwersji PPTX na SCORM: {e}")
        print(f"Error converting PPTX to SCORM: {e}")


def generate_html_files(slide_images, output_folder):
    for i, image in enumerate(slide_images):
        thumbnails = "".join([
                                 f'<a href="slide_{j + 1}.html"><img src="{os.path.basename(img)}" alt="Slide {j + 1}" style="width:100%; background-color: #ccc; {"border: 4px solid red;" if i == j else "border: 2px solid transparent;"}" id="thumb_{j + 1}"></a>'
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
            </script>
        </head>
        <body>
            <div class="sidebar">
                {thumbnails}
            </div>
            <div class="content">
                <img src="{os.path.basename(image)}" alt="Slide {i + 1}">
                <div class="nav-buttons">
                    <button onclick="location.href='{prev_slide}'">Wstecz</button>
                    <button onclick="location.href='{next_slide}'">Dalej</button>
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


def select_ppsx_file():
    ppsx_file = filedialog.askopenfilename(title="Wybierz plik PPSX", filetypes=[("PPSX files", "*.ppsx")])
    if ppsx_file:
        ppsx_label.config(text=ppsx_file)
    return ppsx_file


def select_pptx_file():
    pptx_file = filedialog.askopenfilename(title="Wybierz plik PPTX", filetypes=[("PPTX files", "*.pptx")])
    if pptx_file:
        pptx_label.config(text=pptx_file)
    return pptx_file


def select_output_folder():
    output_folder = filedialog.askdirectory(title="Wybierz folder wyjściowy")
    if output_folder:
        output_label.config(text=output_folder)
    return output_folder


def start_conversion():
    ppsx_file = ppsx_label.cget("text")
    pptx_file = pptx_label.cget("text")
    output_folder = output_label.cget("text")

    if not ppsx_file and not pptx_file:
        messagebox.showerror("Błąd", "Nie wybrano pliku PPSX ani PPTX.")
        return

    if not output_folder:
        messagebox.showerror("Błąd", "Nie wybrano folderu wyjściowego.")
        return

    if ppsx_file:
        pptx_file = ppsx_file.replace(".ppsx", ".pptx")
        convert_ppsx_to_pptx(ppsx_file, pptx_file)

    if pptx_file and os.path.exists(pptx_file):
        convert_pptx_to_scorm(pptx_file, output_folder)
    else:
        messagebox.showerror("Błąd", f"Plik PPTX nie istnieje: {pptx_file}")


def main():
    root = tk.Tk()
    root.title("PPSX to SCORM Converter")
    root.geometry("800x600")

    # Load background image
    bg_image_path = "background.jpg"
    if not os.path.exists(bg_image_path):
        messagebox.showerror("Błąd", "Plik background.jpg nie został znaleziony.")
        return

    bg_image = Image.open(bg_image_path)
    bg_image = bg_image.resize((800, 600), Image.LANCZOS)
    bg_photo = ImageTk.PhotoImage(bg_image)

    # Create a canvas
    canvas = Canvas(root, width=800, height=600)
    canvas.pack(fill="both", expand=True)
    canvas.create_image(0, 0, image=bg_photo, anchor="nw")

    # Add labels and buttons
    canvas.create_text(400, 50, text="PPSX to SCORM Converter", font=("Helvetica", 24), fill="white")
    canvas.create_text(150, 150, text="Wybierz plik PPSX:", font=("Helvetica", 14), fill="white")
    canvas.create_text(150, 200, text="Wybierz plik PPTX:", font=("Helvetica", 14), fill="white")
    canvas.create_text(150, 300, text="Wybierz folder do zapisu:", font=("Helvetica", 14), fill="white")
    canvas.create_text(400, 550, text="Autor programu: Norbert Kowalski CUI Wrocław", font=("Helvetica", 12),
                       fill="white")

    global ppsx_label, pptx_label, output_label
    ppsx_label = Label(root, text="", bg="#1a1a1a", fg="white", font=("Helvetica", 12))
    pptx_label = Label(root, text="", bg="#1a1a1a", fg="white", font=("Helvetica", 12))
    output_label = Label(root, text="", bg="#1a1a1a", fg="white", font=("Helvetica", 12))

    ppsx_button = Button(root, text="Wybierz plik PPSX", command=select_ppsx_file)
    pptx_button = Button(root, text="Wybierz plik PPTX", command=select_pptx_file)
    output_button = Button(root, text="Wybierz folder", command=select_output_folder)
    convert_button = Button(root, text="Konwertuj", command=start_conversion)

    canvas.create_window(300, 150, anchor="w", window=ppsx_button)
    canvas.create_window(300, 200, anchor="w", window=pptx_button)
    canvas.create_window(300, 300, anchor="w", window=output_button)
    canvas.create_window(300, 400, anchor="w", window=convert_button)
    canvas.create_window(500, 150, anchor="w", window=ppsx_label)
    canvas.create_window(500, 200, anchor="w", window=pptx_label)
    canvas.create_window(500, 300, anchor="w", window=output_label)

    root.mainloop()


if __name__ == "__main__":
    main()
