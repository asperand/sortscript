# Sorting procedure:
# 
# Name files in this format: <YEAR/SUBFOLDER> - <TOP-LEVEL CATEGORY> - <NAME OF DOCUMENT>
#
#
# e.g. 1999-ACC-990c.pdf
#
#
#
# YEAR is normally in format [YYYY].
#
# <YEAR/SUBFOLDER> IS FLEXIBLE. It will create a subfolder with whatever you give it. 
#
# e.g. Undated-ACC-BuildingContract.pdf
#
# However, <TOP-LEVEL CATEGORY> is NOT flexible. You will need to add the logic to the switch statement
# found at line 57.
#
#
# **** Run this file using Right Click -> Run with Powershell ****

### BEGIN SCRIPT

# These are for better output when the script finishes.
$incorrect_file_count = 0
$unsorted_file_count = 0
$dupe_file_count = 0

$homedir = "C:\"
$documents = dir -include "*.pdf" -name # only PDF documents
foreach($file in $documents){
	$fileinfo=$file -split "-"
	
	#
	# Some error checking on the name of the file before we try to sort it.
	# Incorrectly spelled categories, subfolders, or filenames are fine,
	# but we want to ensure the filenames are following the 2 dash format and things like:
	# 1999-acc-.pdf, -acc-.pdf, --.pdf, 1999accfile.pdf, are filtered out and skipped.
	#

	if(($fileinfo.length -ne 3) -or
		($fileinfo[0] -eq $homedir + "\Staging\") -or # Missing subfolder name or wrong use of dashes
		($fileinfo[2] -eq ".pdf") -or # Missing filename or wrong use of dashes
		($fileinfo.contains("")) # Empty string anywhere in the array

	){
		$Host.UI.RawUI.BackgroundColor = "DarkRed"
		echo -n "`nERROR:"
		$Host.UI.RawUI.BackgroundColor = "Black" 
		echo "Skipping $file due to incorrect naming convention."
		$incorrect_file_count++
		continue
	}
	#
	# This switch block is where you can add new categories to the file reader.
	# If you follow the same format, it should work with no problems as the script
	# WILL ensure that a directory with name $category exists  or is created before continuing.
	#

	switch($fileinfo[1]){ # assign correct directory name based on category written
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
			$category="Org Documents"
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
		"SD" {
			$category="Staff Documents"
			break
		}
		"HR" {
			$category="HR"
			break
		}
		default {
			$category="Staging\Unsorted"
			$Host.UI.RawUI.BackgroundColor = "DarkRed"
			echo "`nERROR:" 
			$Host.UI.RawUI.BackgroundColor = "Black"
			echo "Incorrect category $fileinfo[1]. Sending to .\Unsorted"
			$unsorted_file_count++
			break
		}
   	}
	$categorydir = $homedir + "\" + $category + "\"
   	ni -itemtype Directory -force -path $categorydir | Out-Null # ensure category directory exists
    	$targetpath=$homedir + "\" + $category + "\" + $fileinfo[0]
   	ni -itemtype Directory -force -path $targetpath | Out-Null # ensure year directory exists
   	$targetfile=$targetpath+"\"+$fileinfo[2]
	$Host.UI.RawUI.BackgroundColor = "DarkGreen"
   	echo "`nACTION:"
	$Host.UI.RawUI.BackgroundColor = "Black" 
	echo "Moving file $file to $targetpath"
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
		$Host.UI.RawUI.BackgroundColor = "DarkRed"
		echo "`nERROR:"
		$Host.UI.RawUI.BackgroundColor = "Black"
		echo "File move failed, likely because filename already exists. Moving $file to $dupepath."
		mv $file $dupepath -ErrorAction 'silentlycontinue' | Out-Null
		$dupe_flag = $?
	}

}
if(($incorrect_file_count -eq 0) -and ($unsorted_file_count -eq 0) -and ($dupe_file_count -eq 0)){
	$Host.UI.RawUI.ForegroundColor = "Green"
	echo "Script finished with zero errors!!!"
	$Host.UI.RawUI.BackgroundColor = "Black"
}
else{
	$Host.UI.RawUI.ForegroundColor = "Yellow"
	echo "`nScript finished with:"
	if($incorrect_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}	
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	echo "$incorrect_file_count skipped files"
	if($unsorted_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}	
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	echo "$unsorted_file_count files sent to \Unsorted\"
	if($dupe_file_count -eq 0){$Host.UI.RawUI.ForegroundColor = "Green"}	
	else{$Host.UI.RawUI.ForegroundColor = "Red"}
	echo "$dupe_file_count duplicate files sent to \Duplicates\`n"
	$Host.UI.RawUI.ForegroundColor = "White"
}
$Host.UI.RawUI.BackgroundColor = "Black"
Write-Host -NoNewLine 'Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
