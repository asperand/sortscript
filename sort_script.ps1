# Sorting procedure:
#
# Name files in this format: <YEAR/SUBFOLDER> - <TOP-LEVEL CATEGORY> - <NAME OF DOCUMENT>
#
#
# e.g. 1999-ACC-990_AND_1099
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

# Setting our console cursor position to bottom to make space for write-progress

$window_height = [Console]::WindowHeight
[Console]::SetCursorPosition(0, ($window_height - 1))

# These are for better output when the script finishes.

$incorrect_file_count = 0
$unsorted_file_count = 0
$dupe_file_count = 0
$current_file_count = 0
$warning_count = 0

# Get our tags and folder names from our tags.cfg file and place them into a hashmap
$tags_map = @{}
$tags_file = get-content -erroraction 'silentlycontinue' tags.cfg | Out-String
if(!$?){
	$tags_map["DEFAULT_TAG"] = "Default"
	Write-Host -BackgroundColor "darkyellow" "`nWARN:"
	Write-Host -ForegroundColor "yellow" "`nCouldn't open tags.cfg. Are you sure it exists?`n"
	$warning_count++
}
else{
	$tags_result = $tags_file -match '{([^{]*?)}'
	if($tags_result -eq $false){
		$tags_map["DEFAULT_TAG"] = "Default"
		Write-Host -BackgroundColor "darkyellow" "`nWARN:"
		Write-Host -ForegroundColor "yellow" "`nCouldn't pull any tags from tags.cfg. Are you sure the syntax is correct?`n"
		$warning_count++
	}
	else{
		$raw_tags = $matches[1]
		$raw_tags_array = $raw_tags -split "`n"
		foreach($pair in $raw_tags_array){
			if($pair -eq ""){ # Skip empties
				continue
			}
			$split_items = $pair -split ":"
			$tags_map[$split_items[0]] = $split_items[1]
		}
	}
}

# Set our directory and file extension info

$cfg_file = get-content -erroraction 'silentlycontinue' sortscript.cfg
if(!$?){ # If the file doesn't exist
	$default_dir = $pwd.path
	$default_filetype = "*.pdf"
	Write-Host -BackgroundColor "darkyellow" "`nWARN:"
	Write-Host -ForegroundColor "yellow" "`nsortscript.cfg file was not found. Setting default path to current working directory.`n"
	$warning_count++
}
else{ # file was found, let's pull the content if we can
	$path_line = select-string .\sortscript.cfg -pattern "TARGET_FOLDER"
	$found_quotes = $path_line -match '"(.*?)"' # only text in quotes
	if($found_quotes -eq $false){ # Coverage for no string match
		$default_dir = $pwd.path
		Write-Host -BackgroundColor "darkyellow" "`nWARN:"
		Write-Host -ForegroundColor "yellow" "`nError processing the sortscript.cfg file. Are you sure your target dir is correct?`n"
		$warning_count++
	}
	$default_dir = $matches[1]
	if($default_dir -eq ""){
		$default_dir = $pwd.path
		Write-Host -BackgroundColor "darkyellow" "`nWARN:"
		Write-Host -ForegroundColor "yellow" "`nFound empty target in sortscript.cfg. Setting to default.`n"
		$warning_count++
	}
	$default_dir_exists = test-path $default_dir
	if($default_dir_exists -eq $false){ # Coverage for provided dir not existing
		$default_dir = $pwd.path
		Write-Host -BackgroundColor "darkyellow" "`nWARN:"
		Write-Host -ForegroundColor "yellow" "`nProvided default target directory does not exist. Are you sure it's correctly set?`n"
		$warning_count++
	}
	$extension_line = select-string .\sortscript.cfg -pattern "DEFAULT_FILE_TYPE"
	$found_quotes = $extension_line -match '"(.*?)"' # only text in quotes
	if($found_quotes -eq $false){ # Coverage for no string match
		$default_filetype = "*.pdf"
		Write-Host -BackgroundColor "darkyellow" "`nWARN:"
		Write-Host -ForegroundColor "yellow" "`nError processing the sortscript.cfg file. Are you sure your default filetype is correct?`n"
		$warning_count++
	}
	$default_filetype = $matches[1]
	$cleaned_input = $default_filetype -replace '[^0-9A-Za-z_]' # Remove all non alphanumerical characters
	if($cleaned_input -eq ""){
		$default_filetype ="*.pdf"
		Write-Host -BackgroundColor "darkyellow" "`nWARN:"
		Write-Host -ForegroundColor "yellow" "`nError processing the sortscript.cfg file. Are you sure your default filetype is correct?`n"
		$warning_count++
	}
	else{
		$default_filetype = "*." + $cleaned_input
	}
}

$staging_dir = $pwd.path
$staging_name = split-path -path $pwd -leaf

# Allow user to set the file extension
# Only fires if the flag is set in cfg
$user_input = Read-Host -Prompt "Enter filetype to move (Default: $default_filetype)"
if($user_input -eq ""){
	$filetype = $default_filetype
}
else{
	$cleaned_input = $user_input -replace '[^0-9A-Za-z_]' # Remove all non alphanumerical characters
	$filetype = "*." + $cleaned_input # Create a usable file extension for Get-ChildItem
}
$documents = Get-ChildItem -include $filetype -name
if(!$documents){
	Write-Host -Foregroundcolor "yellow" "WARN: No files were selected with type $filetype. Did you enter it correctly?"
	$second_try = Read-Host -Prompt "Enter filetype to move (Default: no change)"
	if($second_try -ne ""){
		$cleaned_input = $second_try -replace '[^0-9A-Za-z_]'
		$filetype = "*." + $cleaned_input
	}
	$documents = Get-ChildItem -include $filetype -name
	if(!$documents){
		Write-Host -Foregroundcolor "yellow" "WARN: You are about to select no files with type $filetype."
		while($true){
			[Console]::SetCursorPosition(0, ($window_height - 1))
			$exit_op = Read-Host -Prompt "Would you like to continue with no selected files? (y/n, default: Y)"
			if(($exit_op -eq "Y") -or ($exit_op -eq "")){
				break;
			}
			elseif($exit_op -eq "n"){
				exit
			}
			else{
				Write-Host -NoNewLine -ForegroundColor Red "Please enter either Y, N, or press enter."
				Start-Sleep -seconds 2
				Clear-Host
			}
		}
	}
}
Clear-Host # Clean up the user input text
[Console]::SetCursorPosition(0, ($window_height - 1)) # Reset console position to bottom


$total_file_count = $documents.length
$extension_checker = "." + $cleaned_input # This is a messy way of doing this, but it's necessary for our error checking.

#
# Some extra warnings to ensure our user is using the script correctly.
#

# Give our user a warning if no files were found.
if($total_file_count -eq 0){
	Write-Host -BackgroundColor "darkyellow" "`nWARN:"
	Write-Host -ForegroundColor "yellow" "`nNo files were selected by the script. You may be required to change the default extension in sortscript.cfg.`n"
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
		($file_info[2] -eq $extension_checker) -or 			# Error of missing filename or wrong use of dashes
		($file_info.contains("")) 			# Error of an empty string anywhere in the array

	){ # Print our error message
		Write-Host -NoNewLine -BackgroundColor DarkRed "`nSKIP:"
		Write-Host -NoNewLine -BackgroundColor Black " Skipping $file due to incorrect naming convention."
		$incorrect_file_count++
		continue
	}

	# Find our category based on the tag we found in the filename

	if($tags_map.containskey($file_info[1])){
		$category = $tags_map[$file_info[1]]
		$category = $category -replace '[^0-9A-Za-z_]' # TODO: We should be a bit more conservative about what we're filtering.
		# Ideally, we should just remove illegal path characters and control characters from the string
	}
	else{ # If the given category isn't found
		$category= "Unsorted"
		Write-Host -NoNewLine -BackgroundColor DarkRed "`nNOT FOUND:"
		Write-Host -NoNewLine -BackgroundColor Black " Incorrect category " 
		Write-Host -NoNewLine $file_info[1]
		Write-Host -NoNewLine ". Sending to .\Unsorted`n"
		$unsorted_file_count++
   	}

	if($category -eq "Unsorted"){
		$categorydir = $staging_dir + "\" + $category + "\"
    		$target_path = $staging_dir + "\" + $category + "\" + $file_info[0]
	}
	else{
		$categorydir = $default_dir + "\" + $category + "\"
    		$target_path=$default_dir + "\" + $category + "\" + $file_info[0]
	}

	new-item -itemtype Directory -force -path $categorydir | Out-Null # ensure category directory exists
   	new-item -itemtype Directory -force -path $target_path | Out-Null # ensure subdirectory exists
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
		$dupe_folder = $staging_dir + "\Duplicates"
		new-item -itemtype Directory -force -path $dupe_folder | Out-Null # ensure category directory exists
		$dupe_path = $dupe_folder + "\" + $dupe_file.basename + "_" + $rand.toString() + $dupe_file.extension
		Write-Host -NoNewLine -BackgroundColor DarkRed "`nNO PATH:"
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
	Write-Host "$warning_count system warning(s) thrown`n"
}
$Host.UI.RawUI.ForegroundColor = "White"
if($incorrect_file_count -gt 5){
	Write-Host -BackgroundColor DarkCyan "`nTIP:"
	Write-Host "If your files are being skipped, ensure you are following the 3-dash naming convention."
	Write-Host "e.g. SUBFOLDER-TAG-FILENAME.extension"
}
if($unsorted_file_count -gt 4){
	Write-Host -BackgroundColor DarkCyan "`nTIP:"
	Write-Host "If you are using a custom tag, make sure it's added in the tags.cfg file."
}
if($dupe_file_count -gt 2){
	Write-Host -BackgroundColor DarkCyan "`nTIP:"
	Write-Host "If your files are getting flagged as dupes, it may be an issue with permissions rather than duplicates."
	Write-Host "Ensure you have permissions to write to the target directory set in sortscript.cfg."
}

Write-Output "`nPress any key to continue...";
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
