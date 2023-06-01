# Author 
#
#
#
#


param(
    [string]$BaseFileName = 'Docs\baseFile.docx',
    [string]$ChangedFileName = 'Docs\changedFile.docx',
    [string]$NewFileName = 'Docs\Result\comparison.docx'
)

$ErrorActionPreference = 'Stop'

function resolve($relativePath) {
    (Resolve-Path $relativePath).Path
}

$BaseFileName = resolve $BaseFileName

$ChangedFileName = resolve $ChangedFileName

$NewFileName = (Get-Item .).FullName + "\" +$NewFileName

if (-not(Test-Path -Path $NewFileName -PathType Leaf)){
    
        try{
        New-Item -ItemType File -Path $NewFileName -Force -ErrorAction Stop
        #Write-Output ("New File is Created!: "+ $NewFileName)
        }
        catch{
            Write-Output ("Exception: " + $_.Exception.Message)
        }
}
else {
        #Write-Output ("File already exists!: "+ $NewFileName)
}

Write-Output ("Comparing: "+ $BaseFileName +" AND " +$ChangedFileName)

# Remove the readonly attribute because Word is unable to compare readonly files:
$baseFile = Get-ChildItem $BaseFileName
if ($baseFile.IsReadOnly) {
    $baseFile.IsReadOnly = $false
}

# Constants
$wdDoNotSaveChanges = 0
$wdCompareTargetNew = 2

try {
    $word = New-Object -ComObject Word.Application

    $word.Visible = $false

    $document = $word.Documents.Open($BaseFileName, $false, $false)

    $document.Compare( $changedFileName , [ref]"Comparison", [ref]$wdCompareTargetNew, [ref]$true, [ref]$true)

    $word.ActiveDocument.SaveAs2([ref]$NewFileName)

    $word.ActiveDocument.Close([ref]$wdDoNotSaveChanges)
    

    # Now close the document so only compare results window persists:
    $document.Close([ref]$wdDoNotSaveChanges)

    # Close MS Word
    $word.quit()

    Write-Output ("Comparison is saved in  : "+ $NewFileName)
    
} catch {

    $document.Close([ref]$wdDoNotSaveChanges)

    Write-Output ("Exception: " + $_.Exception.Message)

}
