A simple file-sorter with some built-in customization for tags and target folder. Currently, the name parsing and resulting folder structure is hard-coded, meaning that a category folder will be created at the target directory with a provided subdirectory which the file will be placed in. i.e. SUBFOLDER-TAG-FILENAME. SUBFOLDER is dynamic, meaning that you can feel free to create any name you'd like. Tags can be changed through the *tags.cfg* folder if one follows the formatting. This script has robust error checking and a decent error handling system with plenty of descriptive warnings. It WILL break if you mess around with it too much, but it will work how it's intended 99% of the time even if you make slight mistakes.

Also provided is a Append/Prepend file name editor for mass editing of names.

Eventually, I'd love to do a re-write and re-organization to use functions so the code is a bit cleaner.
