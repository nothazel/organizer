# Desktop File Organizer

This script organizes files on the folders into folders based on their file extensions. It is written in VBScript and can be run on Windows operating systems.

## How it works

When the script is run, it prompts the user for consent to organize the files on their desktop. If the user consents, the script creates an instance of the `FileSystemObject` and gets the absolute path of the desktop. It then creates an empty dictionary to map file extensions to custom folder names. You can add your own custom mappings to this dictionary.

The script then iterates through all the files on the desktop and checks if there is a custom folder name for each file's extension. If there is, it sets the target folder to be the custom folder name. Otherwise, it sets the target folder to be the file's extension.

If the target folder does not exist, the script creates it. The script then checks if a file with the same name already exists in the target folder. If it does, it prompts the user to remove the duplicate file. If the user wants to remove the duplicate, it deletes the duplicated file.

Before moving a file, the script checks if it has already moved 100 files. If it has, it prints out a list of files that have been moved so far and prompts the user if they want to continue moving files. If the user does not want to continue, the script exits. Otherwise, it resets its internal counter and continues moving files.

After moving a file, the script increments its internal counter.

After iterating through all files on desktop, script prints out final list of files that have been moved and folders that have been created along with their respective counts.

## Usage

To use this script, install the `Organizer.vbs` Then, double-click on the script file to run it. The script will prompt you for consent before organizing your files.

You can modify this script as needed by adding your own custom mappings or changing its behavior.

## License

This script is provided as-is without any warranty or support. You are free to use and modify it for your own personal use. Please do not claim it as your own.
