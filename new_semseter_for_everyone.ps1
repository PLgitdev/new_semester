#Application for a new semester

$WordApplication = New-Object -ComObject "word.application"
$WordApplication.Visible = $false


#classes
$classA = "your first class here"
mkdir $HOME\Desktop\$classA
$classB = "your second class here"
mkdir $HOME\Desktop\$classB

# these loops are for an 8 week quarter which are considered semesters at my school (-lt 9) you can edit this to make it longer to a traditional semester

#week loops

for($i= 1; $i -lt 9;$i++){
$class1 = ($classA +"week"+ $i)
$hwFolder1 =("hwweek" + $i)
$mFolder1 = ("mweek" + $i)
$docName  = "dis" +$i +".docx"
mkdir $HOME\Desktop\$classA\$class1\$hwFolder1 
mkdir $HOME\Desktop\$classA\$class1\$mFolder1

$document = $WordApplication.Documents.Add()
$path = ("$HOME\Desktop\"+$classA+"\"+$class1+"\"+$hwFolder1+"\"+$docName)
$document.SaveAs([ref]($path),[ref]$SaveFormat::wdFormatDocument)}

mkdir $HOME\Desktop\$classA\syl

#second class loop

for($i=1; $i -lt 9;$i++){
$class2 = ($classB +"week"+ $i)
$hwFolder2 =("hwweek" + $i)
$mFolder2 = ("mweek" + $i)
$docName  = "dis" +$i +".docx"
mkdir $HOME\Desktop\$classB\$class2\$hwFolder2 
mkdir $HOME\Desktop\$classB\$class2\$mFolder2
$document = $WordApplication.Documents.Add()
$path = ("$HOME\Desktop\"+$classB+"\"+$class2+"\"+$hwFolder2+"\"+$docName)
$document.SaveAs([ref]($path),[ref]$SaveFormat::wdFormatDocument)}

mkdir $HOME\Desktop\$classB\syl

#close it up

$WordApplication.Quit()

#get rid of ComObject and free up memory
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$WordApplication)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
Remove-Variable WordApplication 
