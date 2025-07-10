# Get every PDF and append or prepend a string of your choosing to it.
$window_height = [Console]::WindowHeight
[Console]::SetCursorPosition(0, ($window_height - 1))
$current_file_count = 0
$documents = Get-ChildItem -include "*.pdf" -name
$total_file_count = $documents.length
$ap_str = Read-Host -Prompt "Enter string you would like to add to each .pdf file (e.g. -MEM-AHRI_MEMBERDOCS)`n"
$ap_flag = "A"
while($true){
	$aop = Read-Host -Prompt "Would you like to (P)repend or (A)ppend the string? (Default: Append)"
	if(($aop -eq "A") -or ($aop -eq "")){
		break;
	}
	elseif($aop -eq "P"){
		ap_flag = "P"
		break
	}
	else{
		Write-Output "Please enter either A or P."
	}
}
Get-ChildItem -Path $pwd -Filter "*.pdf" | ForEach-Object {
    	write-progress -Id 0 -Activity "Renaming Files" -Status "$current_file_count / $total_file_count" -CurrentOperation "Editing $_"
	if($ap_flag = "A"){
		$new_name = $_.BaseName + $ap_str + $_.Extension
	}
	elseif($ap_flag = "P"){
		$new_name = $ap_str + $_.BaseName + $_.Extension
	}
    	Rename-Item -Path $_.FullName -NewName $new_name
	$current_file_count++
}
write-progress -Id 0 -Activity "Moving files" -Status "$current_file_count / $total_file_count" -Completed
$current_file_count = 0
$total_file_count = 0
Write-Output "`nPress any key to continue...";
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
