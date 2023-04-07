# Load assembly to create FileBrowser object
Add-Type -AssemblyName System.Windows.Forms

#Info text
Write-Output "
Wähle das Dokument aus, welches du bearbeiten möchtest..

_______________________________________________________________________________
"

# Instantiate FileBrowser object
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    # Default directory which can be seen at startup
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
}

# Runs the FileBrowser, returns OK
$FileBrowser.ShowDialog() | Out-Null

# Stops program if the file is not a .docx Document
if($FileBrowser.SafeFileName -notlike "*.docx"){
    Write-Output "Das war keine .docx Datei. 

Programm wird beendet in..
"

    # Closing process
    Write-Output "10 Sekunden"
    Start-Sleep -Seconds 5
    Write-Output "5 Sekunden"
    Start-Sleep -Seconds 2
    Write-Output "3.."
    Start-Sleep -Seconds 1
    Write-Output "2."
    Start-Sleep -Seconds 1
    Write-Output "1"
    Start-Sleep -Seconds 1
    exit
}

# Open Word and document
$word = NEW-Object –comobject Word.Application
$word.Visible = $false
$document = $word.documents.open($FileBrowser.FileName)

# Count pages
$pages = $document.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)

# Info text
Write-Output "Dein Dokument besteht aus $pages Seiten.

Du kannst nun angeben, welche Seiten du entfernen möchtest.

Trenne mit einem Komma und Benutze einen Bindestrich um Seiten zu verbinden
(Leerzeichen werden ignoriert)

z.B. 2,9-13,17,24-31

-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -
"

# Get pages from host
$pagesStr = Read-Host
$pagesStr = $pagesStr -replace " "
$pagesArr = $pagesStr -split ","

# Delete Pages
for ($i=$pagesArr.Length-1; $i -ge 0; $i--)
{
    if($pagesArr[$i] -match "-"){
        $rangeArr = $pagesArr[$i] -split "-"
        for ($j=[uint16]$rangeArr[1]; $j -ge [uint16]$rangeArr[0]; $j--){
            # Page Deleter
            $word.Selection.GoTo(1, 1, $j) | Out-Null
            $document.Bookmarks("\Page").Range.Delete() | Out-Null
        }
        
    }
    else{
        # Page Deleter
        $word.Selection.GoTo(1, 1, [uint16]$pagesArr[$i]) | Out-Null
        $document.Bookmarks("\Page").Range.Delete() | Out-Null
    }
}

# Save document as new (_Edited)
$docName = $FileBrowser.FileName.Split(".")
$newSave = $docName[0]+"_Edited.docx"
$document.SaveAs2($newSave)

# Close MS Word
$word.Quit()

# Info text
Write-Output "
_______________________________________________________________________________

Dein Dokument wurde fertiggestellt und befindet sich hier." $newSave"
"

# Closing process
Read-Host -Prompt "Drücke ENTER, um zu beenden.."