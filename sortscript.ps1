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

# Setting our console cursor position to bottom
$windowHeight = [Console]::WindowHeight
[Console]::SetCursorPosition(0, ($windowHeight - 1))

# These are for better output when the script finishes.
$incorrect_file_count = 0
$unsorted_file_count = 0
$dupe_file_count = 0
$current_file_count = 0

# Hard-coded home directory to ensure our files are going in the right place
# CHANGE THIS IF THE LOCATION OF THE FOLDER CHANGES

$homedir = "Z:\Shared\AHRI General Share (AllShare)\ArchivingProject"

# only PDF documents

$documents = dir -include "*.pdf" -name
$total_file_count = $documents.length

# Our main loop for going through every pdf file in our list

foreach($file in $documents){
	
	# You can feel free to change the delimiter to whatever character you prefer here if you want to change the naming convention

	$fileinfo=$file -split "-"
	
	#
	# Some error checking on the name of the file before we try to sort it.
	# Incorrectly spelled categories, subfolders, or filenames are fine,
	# but we want to ensure the filenames are following the 2 dash format and things like:
	# 1999-acc-.pdf, -acc-.pdf, --.pdf, 1999accfile.pdf, are filtered out and skipped.
	#

	if(($fileinfo.length -ne 3) -or 			# Error of too little or too many dashes
		($fileinfo[0] -eq $homedir + "\Staging\") -or 	# Error of missing subfolder name or wrong use of dashes
		($fileinfo[2] -eq ".pdf") -or 			# Error of missing filename or wrong use of dashes
		($fileinfo.contains("")) 			# Error of an empty string anywhere in the array

	){ # Print our error message
		write-host -NoNewLine -BackgroundColor DarkRed "`nERROR:"
		write-host -NoNewLine -BackgroundColor Black " Skipping $file due to incorrect naming convention."
		$incorrect_file_count++
		continue
	}

	#
	# This switch block is where you can add new categories to the file reader.
	# If you follow the same format, it should work with no problems as the script
	# WILL ensure that a directory with name $category exists or is created before continuing.
	#

	switch($fileinfo[1].toUpper()){ # assign correct directory name based on category written
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
		default { # If the given category isn't found or is incorrectly spelled
			$category="Staging\Unsorted"
			write-host -NoNewLine -BackgroundColor DarkRed "`nERROR:" 
			write-host -NoNewLine -BackgroundColor Black " Incorrect category $fileinfo[1]. Sending to .\Unsorted`n"
			$unsorted_file_count++
			break
		}
   	}
	$categorydir = $homedir + "\" + $category + "\"
   	ni -itemtype Directory -force -path $categorydir | Out-Null # ensure category directory exists
    	$targetpath=$homedir + "\" + $category + "\" + $fileinfo[0]
   	ni -itemtype Directory -force -path $targetpath | Out-Null # ensure year directory exists
   	$targetfile=$targetpath+"\"+$fileinfo[2]
	$current_file_count++
   	write-progress -Id 0 -Activity "Moving files" -Status "$current_file_count / $total_file_count" -CurrentOperation "Moving $file"
	mv $file $targetfile -ErrorAction 'silentlycontinue' | Out-Null
	
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
		$dupefile = $file | get-childitem
		$dupepath = ".\Duplicates" + "\" + $dupefile.basename + "_" + $rand.toString() + $dupefile.extension
		write-host -NoNewLine -BackgroundColor DarkRed "`nERROR:"
		write-host -NoNewLine -BackgroundColor Black " File move failed, likely because filename already exists. Moving $file to $dupepath.`n"
		mv $file $dupepath -ErrorAction 'silentlycontinue' | Out-Null
		$dupe_flag = $?
	}
}
# All done! If there are no errors, we can celebrate! Otherwise, list out our problems.
write-progress -Id 0 -Activity "Moving files" -Status "$current_file_count / $total_file_count" -Completed
if(($incorrect_file_count -eq 0) -and ($unsorted_file_count -eq 0) -and ($dupe_file_count -eq 0)){
	write-host -ForegroundColor green "`nScript finished with zero errors!!!"
}
else{ # This could use a bit of cleanup.
	write-host -ForegroundColor yellow "`nScript finished with:"
	if($incorrect_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}	
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	write-host "$incorrect_file_count skipped files"
	if($unsorted_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}	
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	write-host "$unsorted_file_count files sent to \Unsorted\"
	if($dupe_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}	
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	write-host "$dupe_file_count duplicate files sent to \Duplicates\"
	$Host.UI.RawUI.ForegroundColor = "White"
}
$Host.UI.RawUI.ForegroundColor = "White"
Write-Host -NoNewLine 'Press any key to close window...';
$key_pressed = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
