# Send MRC Files

This script is used to process and send MARC (`.mrc`) files, created by the SimplyReports MARC export process, to a CollectionHQ FTP server. It is designed to be run as a scheduled task.

## Prerequisites

*   **.NET Framework 4.5 or newer**: This script requires the .NET Framework version 4.5 or higher to be installed on the system.
*   **SimplyReports MARC Export**: The `.mrc` files must be generated using the **CollectionHQ MARC export profile** in SimplyReports.
*   **WinSCP**: This script requires the `WinSCPnet.dll` assembly to be present in the same directory as the script. This is used for transferring files via FTP.

## How it works

The `send-mrc.ps1` script performs the following actions:

1.  **Checks Run Day**: By default, the script will only run on the day of the week specified in `settings.json`. This can be overridden with the `-ForceRun` parameter.
2.  **Deletes old files**: It removes `.mrc` files older than a configured number of hours (default is 120 hours) from the specified directory. It also deletes any non-`.mrc` files.
3.  **Renames files**: It renames the remaining `.mrc` files based on rules defined in `settings.json`. The renaming helps identify which library the file belongs to. The script can use either the hour the file was created or the file's base name to determine the library.
4.  **Uploads files**: It uploads the renamed `.mrc` files to the CollectionHQ FTP server. If the upload fails, it will retry a configurable number of times.
5.  **Logs activity**: The script creates a detailed log file in the `log` subdirectory, which records the script's actions, including which files were processed and uploaded.

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
    *   `retries`: (Optional) The number of times to retry an FTP upload if it fails. Defaults to 2.
*   `libraryMappings`: An array of objects defining rules for renaming files based on library.
    *   `hour`: The hour of the day (in 24-hour format, e.g., "01", "14") the file was created.
    *   `basename`: The name of the file without its extension (e.g., "dcdl").
    *   `libraryName`: The prefix to use for the new file name.
*   `defaultLibraryName`: The prefix to use for files that don't match any rules in `libraryMappings`.

## Running the script

The script can be run from a PowerShell terminal.

```powershell
./send-mrc.ps1
```

To run the script on a day other than the one specified in `settings.json`, use the `-ForceRun` parameter:

```powershell
./send-mrc.ps1 -ForceRun
```

## Best Practices

*   **Scheduling Reports**: If you are running reports for multiple libraries, it is recommended to schedule them to run at different hours. This allows the script to more easily distinguish between them and apply the correct library name when renaming the files. While you can use the `basename` setting to differentiate files created in the same hour, using different hours is a more robust method.

## Logging

The script creates a log file in the `log` subdirectory, which records the script's actions, including which files were processed and uploaded.
