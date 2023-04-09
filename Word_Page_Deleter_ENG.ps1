# imports and definitions
Add-Type -AssemblyName System.Windows.Forms
[Array]$tempIntArr = $null
[Array]$pagesIntArr = $null
$isExcisting = $false

# select .docx file
function ChooseFile{
    echo "___________________________________________________________________`n"
    Read-Host -Prompt "Press ENTER to select the document you want to edit.."
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
    echo "___________________________________________________________________`n"
    Switch ($int){
        1 {
            echo "This was not a .docx file.`nPlease choose another file.`n"
            Read-Host -Prompt "Press ENTER to continue.."
            ChooseFile
            break
        }
        2 {echo "Your input text contains incorrect syntax.`nLetters were captured.`nPlease try again.`n"}
        3 {echo "Your input text contains incorrect syntax.`nComma(s) in a wrong place.`nPlease try again.`n"}
        4 {echo "Your input text contains incorrect syntax.`nAn interval has a higher start than end value.`nPlease try again.`n"}
        5 {echo "Your input text contains incorrect syntax.`nOnly numbers less than $pages may be entered.`nPlease try again.`n"}
        6 {echo "Your input text contains incorrect syntax.`nDash(es) in a wrong place.`nPlease try again.`n"}
        default {echo "An unknown Error occurred."}
    }
    if($int -ne 1){
        $document.Close(0)
        $word.Quit()
        Read-Host -Prompt "Press ENTER to restart.."
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
echo "___________________________________________________________________`n"
echo "Your document has $pages pages.`n`nYou can now type in which pages you want to delete.`n`nSeperate with a comma and use a dash to connect multiple pages.`nSpaces and multiple mentions will be ignored.`n`ne.g. 2,9-13,17,24-31`n`n-  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -`n"


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
echo "`n___________________________________________________________________`n"
echo "Your document has been completed and is located here." $newSafe
$word.Quit()

# closing program
echo "`n___________________________________________________________________`n"
Read-Host -Prompt "Press ENTER to exit the program`nThe path of your new document will be opened afterwards"
ii -Path $path
exit
