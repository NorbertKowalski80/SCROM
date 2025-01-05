import os
import time
import win32com.client

def close_powerpoint_instances():
    try:
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Quit()
        print("Closed existing PowerPoint instances.")
    except Exception as e:
        print(f"Error closing PowerPoint instances: {e}")

def convert_ppsx_to_pptx(ppsx_file, pptx_file):
    close_powerpoint_instances()
    time.sleep(5)  # Give more time to ensure PowerPoint instances are closed
    try:
        if not os.path.exists(ppsx_file):
            raise FileNotFoundError(f"File {ppsx_file} not found.")
        if not os.access(ppsx_file, os.R_OK):
            raise PermissionError(f"File {ppsx_file} is not readable.")

        print(f"Starting PowerPoint application...")
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1
        print(f"Opening PPSX file: {ppsx_file}")
        presentation = powerpoint.Presentations.Open(ppsx_file, WithWindow=False)
        presentation.SaveAs(pptx_file, 24)  # 24 oznacza format pptx
        presentation.Close()
        powerpoint.Quit()
        print(f"Converted {ppsx_file} to {pptx_file}")
    except Exception as e:
        print(f"Error converting PPSX to PPTX: {e}")

if __name__ == "__main__":
    # Przykładowe użycie
    ppsx_file = r'C:\Users\kowal\OneDrive\Pulpit\CUI Prezentacja\eLearning_CUI_WrocławEDUEOFVer.ppsx'
    pptx_file = r'C:\Users\kowal\OneDrive\Pulpit\CUI Prezentacja\eLearning_CUI_WrocławEDUEOFVer.pptx'
    convert_ppsx_to_pptx(ppsx_file, pptx_file)
