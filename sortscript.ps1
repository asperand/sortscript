# Sorting procedure:
#
# Name files in this format: <YEAR/SUBFOLDER> - <TOP-LEVEL CATEGORY> - <NAME OF DOCUMENT>
#
#
# e.g. 1999-ACC-AHRI990c.pdf
#
#
#
# YEAR is normally in format [YYYY].
#
# <YEAR/SUBFOLDER> IS FLEXIBLE. It will create a subfolder with whatever you give it.
#
# e.g. Undated-ACC-BuildingContract.pdf
#
# However, <TOP-LEVEL CATEGORY> is NOT flexible. You will need to add the logic to the switch statement.
#
#
#
# **** Run this file using Right Click -> Run with Powershell ****

### BEGIN SCRIPT

# Setting our console cursor position to bottom to make space for write-progress

$window_height = [Console]::WindowHeight
[Console]::SetCursorPosition(0, ($window_height - 1))

# These are for better output when the script finishes.

$incorrect_file_count = 0
$unsorted_file_count = 0
$dupe_file_count = 0
$current_file_count = 0
$warning_count = 0

# Hard-coded home directory to ensure our files are going in the right place

$default_dir = "Z:\Shared\AHRI General Share (AllShare)\ArchivingProject" # CHANGE THIS IF THE LOCATION OF THE FOLDER CHANGES
$staging_dir = $default_dir + "\Staging" # No need to change this unless you want to rename the default structure.

# only select PDF documents
$documents = Get-ChildItem -include "*.pdf" -name
$total_file_count = $documents.length

#
# Some extra warnings to ensure our user is using the script correctly.
#

# Give our user a warning if no files were found.
if($total_file_count -eq 0){
	Write-Host -BackgroundColor "darkyellow" "`nWARN:"
	Write-Host -ForegroundColor "yellow" "`nNo files were selected by the script. `nIf the Staging folder is not empty, ensure your home directory is set correctly.`nCurrent home directory: $default_dir"
	$warning_count++
}

# Give our user a warning if the current working directory is wrong.
if($pwd.path -ne $staging_dir){
	Write-Host -BackgroundColor "darkyellow" "`nWARN:"
	Write-Host -ForegroundColor "yellow" "`nCurrent working directory is different than expected. Some functions may break!`nPlease move this script to $staging_dir`nor change the home directory."
	$warning_count++
}

#
# Our main loop for going through every pdf file in our list
#

foreach($file in $documents){

	# You can feel free to change the delimiter to whatever character you prefer here if you want to change the naming convention

	$file_info=$file -split "-"

	#
	# Some error checking on the name of the file before we try to sort it.
	# Incorrectly spelled categories, subfolders, or filenames are fine,
	# but we want to ensure the filenames are following the 2 dash format and things like:
	# 1999-acc-.pdf, -acc-.pdf, --.pdf, 1999accfile.pdf, are filtered out and skipped.
	#

	if(($file_info.length -ne 3) -or 			# Error of too little or too many dashes
		($file_info[0] -eq $default_dir + $staging_dir + "\") -or 	# Error of missing subfolder name or wrong use of dashes
		($file_info[2] -eq ".pdf") -or 			# Error of missing filename or wrong use of dashes
		($file_info.contains("")) 			# Error of an empty string anywhere in the array

	){ # Print our error message
		Write-Host -NoNewLine -BackgroundColor DarkRed "`nERROR:"
		Write-Host -NoNewLine -BackgroundColor Black " Skipping $file due to incorrect naming convention."
		$incorrect_file_count++
		continue
	}

	#
	# This switch block is where you can add new categories to the file reader.
	# If you follow the same format, it should work with no problems as the script
	# WILL ensure that a directory with name $category exists or is created before continuing.
	#

	switch($file_info[1].toUpper()){ # assign correct directory name based on category written
		"ACC" {
			$category="Accounting"
			break
		}
		"AG" {
			$category="Agendas"
			break
		}
		"ETC" {
			$category="Etc Documents"
			break
		}
		"MEM" {
			$category="Membership"
			break
		}
		"MIN" {
			$category="Minutes"
			break
		}
    		"NATE" {
			$category="NATE Documents"
			break
		}
    		"REP" {
			$category="Reports"
			break
		}
    		"LEG" {
			$category="Legal"
			break
		}
    		"RM" {
			$category="Rulemaking"
			break
		}
		"SD" {
			$category="Staff Documents"
			break
		}
		"HR" {
			$category="HR"
			break
		}
		default { # If the given category isn't found
			$category= $staging_dir + "\Unsorted" # Feel free to change the end string if adjusting the file structure.
			Write-Host -NoNewLine -BackgroundColor DarkRed "`nERROR:"
			Write-Host -NoNewLine -BackgroundColor Black " Incorrect category $file_info[1]. Sending to .\Unsorted`n"
			$unsorted_file_count++
			break
		}
   	}
	$categorydir = $default_dir + "\" + $category + "\"
   	new-item -itemtype Directory -force -path $categorydir | Out-Null # ensure category directory exists
    	$target_path=$default_dir + "\" + $category + "\" + $file_info[0]
   	new-item -itemtype Directory -force -path $target_path | Out-Null # ensure year directory exists
   	$target_file=$target_path+"\"+$file_info[2]
	$current_file_count++
   	write-progress -Id 0 -Activity "Moving files" -Status "$current_file_count / $total_file_count" -CurrentOperation "Moving $file"
	mv $file $target_file -ErrorAction 'silentlycontinue' | Out-Null

	#
	# This block fires in the case that the original file move is not successful.
	# While there could be many reasons, the most likely culprit is a duplicate file name.
	# A random number is appended to the filename. The use of WHILE allows this to loop in the case that there
	# is ANOTHER file in the Duplicates folder with the same exact name and random number.
	# This will break if there are 10,000 duplicate items with numbers 0-9999. (But that should never happen)
	#

	# This funky bit of code is needed for counting dupe files correctly, but the overall functionality stays the same
	$dupe_flag = $?
	if(!$dupe_flag){
		$dupe_file_count++
	}
	while(!$dupe_flag){
		$rand = get-random -maximum 9999
		$dupe_file = $file | get-childitem
		$dupe_path = ".\Duplicates" + "\" + $dupe_file.basename + "_" + $rand.toString() + $dupe_file.extension
		Write-Host -NoNewLine -BackgroundColor DarkRed "`nERROR:"
		Write-Host -NoNewLine -BackgroundColor Black " File move failed, likely because filename already exists. Moving $file to $dupe_path.`n"
		mv $file $dupe_path -ErrorAction 'silentlycontinue' | Out-Null
		$dupe_flag = $?
	}
}
# All done! If there are no errors, we can celebrate! Otherwise, list out our problems.
write-progress -Id 0 -Activity "Moving files" -Status "$current_file_count / $total_file_count" -Completed
# Extra lines to fix the problem of file counts carrying over from session to session.
$total_file_count = 0 
Remove-Variable -Name documents
if(($incorrect_file_count -eq 0) -and ($unsorted_file_count -eq 0) -and ($dupe_file_count -eq 0) -and ($warning_count -eq 0)){
	Write-Host -ForegroundColor green "`nScript finished with zero errors!!!"
}
else{ # This could use a bit of cleanup.
	Write-Host -ForegroundColor yellow "`n`nScript finished with:"
	if($incorrect_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	Write-Host "$incorrect_file_count skipped file(s)"
	if($unsorted_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	Write-Host "$unsorted_file_count file(s) sent to \Unsorted\"
	if($dupe_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	Write-Host "$dupe_file_count duplicate file(s) sent to \Duplicates\"
	if($warning_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}
	else{$Host.UI.RawUI.ForegroundColor = "DarkYellow"}
	Write-Host "$warning_count system warning(s) thrown"
}
$Host.UI.RawUI.ForegroundColor = "White"
Write-Output "`nPress any key to continue...";
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
