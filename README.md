# PowerShell-DOCX-PDF-Converter
A PowerShell based DOCX to PDF converter based on using LibreOffice instead of Microsoft Office. This uses either a installed or portable version of LibreOffice to convert your documents from docx to pdf. The script is made to bulk (batch) convert from directory A to directory B. Feel free to edit the script to your needs if you require otherwise.

### Usage
```
.\Convert-Documents.ps1 -InputDirectory C:\Users\example\Desktop\In -OutputDirectory C:\Users\example\Desktop\Out\ -ConvertTo pdf -LibreOfficeExe <LibreOffice (soffice.exe) location>

.\Convert-Documents.ps1 -InputDirectory C:\Users\example\Desktop\In -OutputDirectory C:\Users\example\Desktop\Out\ -ConvertTo pdf -LibreOfficeExe <LibreOffice (soffice.exe) location>
```

### LibreOffice
As I said you can either use the portable version or the installed version on your computer. 

Please visit https://libreoffice.org/ for the download of this amazing product. 

If you wish to directly use the portable version then please go here https://www.libreoffice.org/download/portable-versions/

### Learning
I am still a beginner in PowerShell atleast I feel that way and I am learning it through my work as I need to get things done so please be kind to me when you see any weird mistakes or things in the script. I hope it helps someone somewhere :D !
