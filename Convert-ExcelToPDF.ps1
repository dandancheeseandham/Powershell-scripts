function Convert-ExcelToPDF {
    param (
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )

    # Check if the folder exists
    if (-not (Test-Path $FolderPath)) {
        Write-Error "The folder '$FolderPath' does not exist."
        return
    }

    # Load the required Excel COM assembly
    Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"

    # Create an Excel application object
    $excelApp = New-Object -ComObject "Excel.Application"

    # Set Excel application visibility to False (hidden)
    $excelApp.Visible = $false

    # Define the Microsoft Print to PDF printer
    $printerName = "Microsoft Print to PDF"

    # Define the Paper Size (A3)
    $paperSize = [Microsoft.Office.Interop.Excel.XlPaperSize]::xlPaperA3

    # Process each Excel file in the folder
    Get-ChildItem -Path $FolderPath -Filter "*.xls*" | ForEach-Object {
        $excelFilePath = $_.FullName
        $pdfFilePath = [System.IO.Path]::ChangeExtension($excelFilePath, "pdf")

        try {
            # Open the Excel file
            $workbook = $excelApp.Workbooks.Open($excelFilePath)

            # Set the paper size for each worksheet to A3
            foreach ($worksheet in $workbook.Worksheets) {
                $worksheet.PageSetup.PaperSize = $paperSize
            }

            # Save the workbook as PDF using the Microsoft Print to PDF printer
            $workbook.ExportAsFixedFormat([Microsoft.Office.Interop.Excel.XlFixedFormatType]::xlTypePDF, $pdfFilePath)
            Write-Host "Successfully converted '$excelFilePath' to '$pdfFilePath'."
        }
        catch {
            Write-Error "An error occurred while converting '$excelFilePath': $_"
        }
        finally {
            # Close the workbook and release the COM object
            if ($workbook) {
                $workbook.Close($false)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook)
            }
        }
    }

    # Quit the Excel application and release the COM object
    $excelApp.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelApp)
}

Convert-ExcelToPDF -FolderPath "C:\path\to\your\folder"