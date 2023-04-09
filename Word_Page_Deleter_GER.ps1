# imports and definitions
Add-Type -AssemblyName System.Windows.Forms
[Array]$tempIntArr = $null
[Array]$pagesIntArr = $null
$isExcisting = $false

# select .docx file
function ChooseFile{
    echo "______________________________________________________________________`n"
    Read-Host -Prompt "Drücke ENTER um ein Dokument zum Bearbeiten auszuwählen.."
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop')
    }
    $FileBrowser.ShowDialog() | Out-Null
    if($FileBrowser.SafeFileName -notlike "*.docx"){
        ErrorMsg -int 1
        return
    }
    Set-Variable "FileBrowser" -Value $FileBrowser -Scope Global
}

# error message and stop program
function ErrorMsg {
    param ($int)
    echo "______________________________________________________________________`n"
    Switch ($int){
        1 {
            echo "Das ist keine .docx Datei.`nBitte wähle eine andere Datei aus.`n"
            Read-Host -Prompt "Drücke ENTER um fortzufahren.."
            ChooseFile
            break
        }
        2 {echo "Der eingegebene Text enthält eine fehlerhafte Syntax.`nEs wurden Buchstaben erfasst.`nBitte versuche es erneut.`n"}
        3 {echo "Der eingegebene Text enthält eine fehlerhafte Syntax.`nFalsche Kommasetzung.`nBitte versuche es erneut.`n"}
        4 {echo "Der eingegebene Text enthält eine fehlerhafte Syntax.`nEin Interval weißt einen höheren Start- als Endwert auf.`nBitte versuche es erneut.`n"}
        5 {echo "Der eingegebene Text enthält eine fehlerhafte Syntax.`nEs dürfen nur Zahlen eingegeben werden die kleiner als $pages sind.`nBitte versuche es erneut.`n"}
        6 {echo "Der eingegebene Text enthält eine fehlerhafte Syntax.`nBindestriche falsch gesetzt.`nBitte versuche es erneut.`n"}
        default {echo "Ein unbekannter Fehler ist aufgetreten.`n"}
    }
    if($int -ne 1){
        $document.Close(0)
        $word.Quit()
        Read-Host -Prompt "Drücke ENTER um neuzustarten.."
        powershell "& $PSCommandPath"
        exit
    }
}

ChooseFile

# open document in word, count pages and give info
$word = NEW-Object –comobject Word.Application
$word.Visible = $false
$document = $word.documents.open($FileBrowser.FileName)
$pages = $document.ComputeStatistics([Microsoft.Office.Interop.Word.WdStatistic]::wdStatisticPages)
echo "______________________________________________________________________`n"
echo "Dein Dokument besteht aus $pages Seiten.`n`nDu kannst nun angeben, welche Seiten du entfernen möchtest.`n`nTrenne mit einem Komma und Benutze einen Bindestrich für Intervalle.`nLeerzeichen und Mehrfachnennungen werden ignoriert.`n`nz.B. 2,9-13,17,24-31`n`n-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -`n"


# get pages and check syntax
$inputStr = Read-Host
$tempStrArr = $inputStr -replace " " -split ","
if($inputStr -match "[a-z]"){ErrorMsg -int 2}
for($i = 0; $i -lt $tempStrArr.Length; $i++){
    if(!($tempStrArr[$i])){
        ErrorMsg -int 3
        break
    }
}

# sort input into integer array
for ($i=$tempStrArr.Length-1; $i -ge 0; $i--){
    if($tempStrArr[$i] -like "*-*-*"){
        ErrorMsg -int 6
        break
    }
    if($tempStrArr[$i] -match "-"){
        $rangeArr = $tempStrArr[$i] -split "-"
        if([Int]$rangeArr[1] -lt [Int]$rangeArr[0]) {
            ErrorMsg -int 4
            break
        }elseif([Int]$rangeArr[1] -gt $pages){
            ErrorMsg -int 5
            break
        }
        for ($j=[Int]$rangeArr[1]; $j -ge [Int]$rangeArr[0]; $j--){
            $tempIntArr+=$j
        }
    }elseif([Int]$tempStrArr[$i] -gt $pages){
        ErrorMsg -int 5
        break
    }elseif([Int]$tempStrArr[$i] -in 1..$pages){
        $tempIntArr+=[Int]$tempStrArr[$i]
    }else{
        ErrorMsg
        break
    }
}
$tempIntArr = $tempIntArr | sort -descending

# sort out multiple mentions of pages
$pagesIntArr += $tempIntArr[0]
for($i = 1; $i -lt $tempIntArr.Length; $i++){
    $temp = $pagesIntArr.Length
    for($j = 0; $j -lt $temp; $j++){
        if($pagesIntArr[$j] -eq $tempIntArr[$i]){
            $isExcisting = $true
        }
    }
    if($isExcisting -eq $false){
        $pagesIntArr += $tempIntArr[$i]
    }
    $isExcisting = $false
}

# delete pages
for ($i=0; $i -lt $pagesIntArr.Length; $i++){
    $word.Selection.GoTo(1, 1, $pagesIntArr[$i]) | Out-Null
    $document.Bookmarks("\Page").Range.Delete() | Out-Null
}

# save document and close MS Word
$path = $document.Path
$newSafe = $path+"\"+$FileBrowser.SafeFileName.Split(".")[0]+"_Edited.docx"
$document.SaveAs2($newSafe)
echo "`n______________________________________________________________________`n"
echo "Dein Dokument wurde fertiggestellt und befindet sich hier." $newSafe
$word.Quit()

# closing program
echo "`n______________________________________________________________________`n"
Read-Host -Prompt "Drücke ENTER um das Programm zu schließen`nDer Pfad deines neuen Dokuments wird im Anschluss geöffnet."
ii -Path $path
exit
