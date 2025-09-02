# Get every instance of a filetype and append or prepend a string of your choosing to it.

$window_height = [Console]::WindowHeight
[Console]::SetCursorPosition(0, ($window_height - 1))
$current_file_count = 0
$user_input = Read-Host -Prompt "Enter filetype to modify (Default: .pdf)"
if($user_input -eq ""){
	$filetype = "*.pdf"
}
else{
	$cleaned_input = $user_input -replace '[^0-9A-Za-z_]' # Remove all non alphanumerical characters
	$filetype = "*." + $cleaned_input # Create a usable file extension for Get-ChildItem
}

$documents = Get-ChildItem -include $filetype -name

# If no documents were selected, we can give the user a shot to correct themselves
# If they try again and there are still no files, simply exit the program

if(!$documents){
	Write-Host -Foregroundcolor "yellow" "WARN: No files were selected with type $filetype. Did you enter it correctly?"
	$second_try = Read-Host -Prompt "Enter filetype to modify (Default: no change)"
	if($second_try -ne ""){
		$cleaned_input = $second_try -replace '[^0-9A-Za-z_]'
		$filetype = "*." + $cleaned_input
	}
	$documents = Get-ChildItem -include $filetype -name
	if(!$documents){
		Write-Host -Foregroundcolor "yellow" "WARN: No files were selected with type $filetype."
		Write-Output "`nPress any key to exit script...";
		$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
		exit
	}
}


$total_file_count = $documents.length # This is for Write-Progress' sake

# Give the user the option of either prepending or appending the string

$ap_str = Read-Host -Prompt "Enter string you would like to add to each $filetype file (e.g. -MEM-AHRI_MEMBERDOCS)`n"
while($true){
	$aop = Read-Host -Prompt "Would you like to (P)repend or (A)ppend the string? "
	if(($aop -eq "A")){
		$ap_flag = "A"
		break;
	}
	elseif($aop -eq "P"){
		$ap_flag = "P"
		break;
	}
	else{
		Write-Output "Please enter either A or P."
	}
}

# Edit the name of each file with the provided extension based on the user's option to prepend or append

Get-ChildItem -Path $pwd -Filter $filetype | ForEach-Object {
    	write-progress -Id 0 -Activity "Renaming Files" -Status "$current_file_count / $total_file_count" -CurrentOperation "Editing $_"
	if($ap_flag -eq "A"){
		$new_name = $_.BaseName + $ap_str + $_.Extension
	}
	elseif($ap_flag -eq "P"){
		$new_name = $ap_str + $_.BaseName + $_.Extension
	}
    	Rename-Item -Path $_.FullName -NewName $new_name
	$current_file_count++
}

# cleanup and exit

write-progress -Id 0 -Activity "Moving files" -Status "$current_file_count / $total_file_count" -Completed
$current_file_count = 0
$total_file_count = 0
Write-Output "`nPress any key to continue...";
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
