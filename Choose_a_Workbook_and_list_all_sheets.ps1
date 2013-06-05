# Thanks Ati for the file explorer call and save the results to a text file
Function Get-FileName()
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null


 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} #end function Get-FileName



$initialDirectory = ".\"
$filePath = Get-FileName
$filePath | out-file "C:\Data\Powershell_Library\filePath.txt" -Encoding ASCII

#Thanks to Bryan C O'Connell for the listing worksheets cmdlets logic  http://bryanoconnell.blogspot.com/p/licenses.html
$Excel = New-Object -ComObject "Excel.Application" 
$Excel.Visible = $false                               #Runs Excel in the background. 
$Excel.DisplayAlerts = $false                         #Supress alert messages. 
$Workbook = $Excel.Workbooks.open($filepath) 
#Cycle through the Workbook and list each Worksheet
if ($Workbook.Worksheets.Count -gt 0)
 { 
    $Worksheet = $Workbook.Worksheets.item(1) 
    foreach($Worksheet in $Workbook.Worksheets) { 
	    $worksheetname =  $Worksheet.Name | out-file "C:\Data\Powershell_Library\worksheetlist.txt" -Encoding ASCII -append
         } 
} 

$Workbook.Close() 
$Excel.Quit()
