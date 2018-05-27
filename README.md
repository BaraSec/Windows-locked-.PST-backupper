# Windows-locked-.PST-backupper

A script to backup your files, compress them, including locked outlook .PST files while they are being used.


Has 2 modes, an interactive mode, and a quiet one. In the quiet mode it reads the destination and source location from the "profile.txt" file. File format is: dest. folder, new line, and then the source files/folders. Having "outlook" in a line by itself instructs the script to locate the %HOMEPATH% of the user, and the typical outlook-files folder and copies it. Logs everything in a detailed manner to a file called "logs.txt" in the same directory as the script.

Requires administrator privileges in case locked .PST files need to be copied.  
