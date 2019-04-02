<# 
 .Synopsis
  Creates a PDF document from a txt payslip embedded in an email.

 .Parameter path
  The folder path containing the email messages (*.msg)

 .Description
  This script is useful when a payslip is embedded as an email attachment.
  There are a few assumptions made:
  1. The payslip is in a txt format
  2. The payslip is attached to an outlook email
  3. The email(s) have been exported from outlook and are in the .msg format
  4. Outlook and Word are both installed on the computer
  5. The payslip contains a line starting with PAY PERIOD followed by the date

  If each of these assumptions hold, the script will extract the attachment, 
  rename it and convert it into a PDF document.

  NOTE: This script should not be run as administrator

 .Example
   .\Get-PayslipFromEmail.ps1 -path "C:\temp"
#>

param([String]$folderPath)

<# 
 .Synopsis
  Extracts attachments from an outlook message

 .Parameter path
  The folder path containing the email messages (*.msg)

 .Description
  Extracts all attachments from email messages and deposits them in the same folder
#>
function Get-MsgAttachment
{
    Param
    (
        [String]$Path
    )

    # Load application
    Write-Host "Extracting attachments from emails..." -ForegroundColor Yellow
    Write-Host "Source directory $Path"
    $outlook = New-Object -ComObject Outlook.Application
    
    $messageItems = Get-ChildItem -Path $Path -Filter "*.msg"
    $messageItems | ForEach-Object {
        # Work out file names
        $msgFn = $_.FullName
        $dir = $_.DirectoryName

        # Extract message body
        Write-Host "    Extracting attachments from $_..."
        $msg = $outlook.CreateItemFromTemplate($msgFn)
        $msg.Attachments | ForEach-Object {
            # Work out attachment file name
            $attFn = "$dir\$($_.FileName)"

            # Do not try to overwrite existing files
            if (Test-Path -literalPath $attFn) {
                Write-Host "    Skipping $($_.FileName) (file already exists)..."
                return
            }

            # Save attachment
            Write-Host "    Saving $($_.FileName)..."
            $_.SaveAsFile($attFn)

            # Output to pipeline
            Get-ChildItem -LiteralPath $attFn
        }
    }

    Write-Host "All attachments extracted" -ForegroundColor Green
}

<# 
 .Synopsis
  Converts a month string to a number

 .Parameter path
  The short form name of the month
#>
function ConvertMonth {
    Param
    (
        [String]$month
    )
    
    switch ( $month )
    {
        "Jan" { $result = "01" }
        "Feb" { $result = "02" }
        "Mar" { $result = "03" }
        "Apr" { $result = "04" }
        "May" { $result = "05" }
        "Jun" { $result = "06" }
        "Jul" { $result = "07" }
        "Aug" { $result = "08" }
        "Sep" { $result = "09" }
        "Oct" { $result = "10" }
        "Nov" { $result = "11" }
        "Dec" { $result = "12" }
        default { $result = "Unknown-$month" }
    }

    return $result
}

<# 
 .Synopsis
  Reads the content of the text file and renames it based on the contained date

 .Parameter FilePath
  The path to the text file containing the date
#>
function Set-CorrectPayslipName {
    Param
    (
        [String]$FilePath
    )
    
    $regex = "PAY PERIOD\s*(\d\d)\s*([A-z]*)\s*(\d\d\d\d)\s*TO\s* (\d\d) ([A-z]*) (\d\d\d\d)"
    foreach($line in Get-Content $filePath) {
        if($line -match $regex){
            #$result = $line -match $regex
    
            $secondMonthNo = ConvertMonth $Matches[5]        
            $fileName = "$($Matches[6])-$secondMonthNo-$($Matches[4]) - saab au payslip.txt"
            
            try {
                Rename-Item -Path $FilePath -NewName $fileName
            }
            catch {
                $fileName = "$($Matches[6])-$secondMonthNo-$($Matches[4]) - saab au payslip (extra).txt"
                Rename-Item -Path $FilePath -NewName $fileName
            }

            $folder = Split-Path -Path $FilePath
            $result = Join-Path $folder $fileName
            return $result
         }
    }  

}

<# 
 .Synopsis
  Creates a pdf document from a text payslip file

 .Parameter path
  The folder path containing the payslips messages (*.txt)

 .Description
  Renames the text files based on the date and then creates a PDF document from them.
#>
function New-PdfFromPayslip {
    Param
    (
        [String]$Path
    )

    Write-Host "Creating PDF documents from txt payslips" -ForegroundColor Yellow
    
    $textFiles = Get-ChildItem -Path $Path -Filter "*.txt"
    $textFiles | ForEach-Object {
        # Work out file names
        $fileName = $_.FullName
        $newName = Set-CorrectPayslipName $fileName

        # File paths
        $txtPath = $newName
        $pdfPath = [io.path]::ChangeExtension($newName, "pdf")
        
        Write-Host "    Creating pdf for $pdfPath"
        
        # Required Word Variables
        $wdExportFormatPDF = 17
        $wdDoNotSaveChanges = 0

        # Create a hidden Word window
        $word = New-Object -ComObject word.application
        $word.visible = $false

        # Open the document
        $doc = $word.Documents.Open($txtPath)

        # Set the page orientation to portrait
        $doc.PageSetup.Orientation = 0

        # Export the PDF file and close without saving a Word document
        $doc.ExportAsFixedFormat($pdfPath,$wdExportFormatPDF)
        $doc.close([ref]$wdDoNotSaveChanges)
        $word.Quit() 
    }

    Write-Host "Conversion Complete" -ForegroundColor Green
}

if (-not (Test-Path $folderPath)) {
    Write-Host "Please enter a valid path..." -ForegroundColor Red
    return
}

Get-MsgAttachment -Path $folderPath
New-PdfFromPayslip -Path $folderPath