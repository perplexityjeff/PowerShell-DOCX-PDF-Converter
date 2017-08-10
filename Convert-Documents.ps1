<#
.SYNOPSIS
A PowerShell and LibreOffice based converter mainly made to convert DOCX to PDF

.DESCRIPTION
A PowerShell script that uses either a portable or installed version of LibreOffice and the command line interface of it to convert documents in bulk (batch)

.EXAMPLE
.\Convert-Documents.ps1 -InputDirectory C:\Users\example\Desktop\In -OutputDirectory C:\Users\example\Desktop\Out\ -ConvertTo pdf -LibreOfficeExe <LibreOffice (soffice.exe) location>

.EXAMPLE
.\Convert-Documents.ps1 -InputDirectory C:\Users\example\Desktop\In -OutputDirectory C:\Users\example\Desktop\Out\ -ConvertFrom docx -ConvertTo pdf -LibreOfficeExe <LibreOffice (soffice.exe) location>

.EXAMPLE
.\Convert-Documents.ps1 -InputDirectory C:\Users\example\Desktop\In -OutputDirectory C:\Users\example\Desktop\Out\ -ConvertFrom docx -ConvertTo pdf -LibreOfficeExe <LibreOffice (soffice.exe) location>

.LINK
https://github.com/perplexityjeff/PowerShell-DOCX-PDF-Converter
#>

[cmdletbinding()]
param(
[Parameter(Mandatory=$true)]$InputDirectory,
[Parameter(Mandatory=$true)]$OutputDirectory,
[Parameter(Mandatory=$true)]$LibreOfficeExe, 
$ConvertTo = "pdf",
$ConvertFrom = "docx"
)

Write-Verbose "Checking if LibreOffice EXE Location exists"
If (!(Test-Path $LibreOfficeExe))
{
    Write-Error "LibreOffice EXE location was not found, please specify another location"
}

Write-Verbose "Checking if OutputDirectory exists"
if (!(Test-Path $OutputDirectory))
{
    New-Item -ItemType Directory $Out
}

Write-Verbose "Checking if InputDirectory exists"
if (!(Test-Path $InputDirectory))
{
    Write-Error "Input directory was not found, please speciy another location"
}

$Files = Get-ChildItem $InputDirectory -Filter "*.$ConvertFrom"

foreach($File in $Files)
{
    if ($File.Exists)
    {
        $Argument = '--headless --convert-to ' + $ConvertTo + ' --outdir "' + $OutputDirectory + '" "' + $File.FullName + '"'
        Write-Verbose "Starting convert using Arguments: $Argument"
        Start-Process $LibreOfficeExe -ArgumentList $Argument -Wait
        Write-Verbose "$File has been converted"
    }
}

Write-Verbose "Script has been completed"

