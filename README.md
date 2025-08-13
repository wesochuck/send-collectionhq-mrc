# Send MRC Files

This script is used to process and send MARC (`.mrc`) files, created by the SimplyReports MARC export process, to a CollectionHQ FTP server. It is designed to be run as a scheduled task, once a week.

## Prerequisites

This script requires the `WinSCPnet.dll` assembly to be present in the same directory as the script. This is used for transferring files via FTP.

## How it works

The `send-mrc.ps1` script performs the following actions:

1.  **Deletes old files**: It removes `.mrc` files older than a configured number of hours (default is 120 hours) from the specified directory. It also deletes any non-`.mrc` files.
2.  **Renames files**: It renames the remaining `.mrc` files based on rules defined in `settings.json`. The renaming helps identify which library the file belongs to.
3.  **Uploads files**: It uploads the renamed `.mrc` files to the CollectionHQ FTP server.

## Configuration

The script's behavior is customized through the `settings.json` file. This file must be in the same directory as the script.

### `settings.json` Structure

*   `basepath`: The base directory that contains the folders to be processed.
*   `folders`: An array of folder names (relative to `basepath`) to search for `.mrc` files.
*   `hours`: The maximum age of files (in hours) to keep. Files older than this will be deleted.
*   `ext`: The file extension to process (e.g., `.mrc`).
*   `dayofweektorun`: The specific day of the week the script is allowed to run (e.g., "Tuesday").
*   `transcriptpath`: The full path for the PowerShell transcript log file.
*   `ftp`: An object containing FTP server details.
    *   `url`: The FTP server hostname.
    *   `user`: The FTP username.
    *   `pass`: The FTP password.
    *   `remotepath`: The directory on the FTP server to upload files to.
*   `libraryMappings`: An array of objects defining rules for renaming files based on library.
    *   `hour`: The hour of the day (in 24-hour format) the file was created.
    *   `basename`: The name of the file without its extension.
    *   `libraryName`: The prefix to use for the new file name.
*   `defaultLibraryName`: The prefix to use for files that don't match any rules in `libraryMappings`.

## Running the script

The script can be run from a PowerShell terminal. It is designed to be run as a scheduled task.

```powershell
./send-mrc.ps1
```
## Logging

The script creates a log file in the `log` subdirectory, which records the script's actions, including which files were processed and uploaded.
