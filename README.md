# DOCX to PDF Converter Application

## Overview

This Python application converts Microsoft Word documents (.docx) in a selected folder into PDF format. It uses the `docx2pdf` library for conversion and provides a user interface through `tkinter`. This application is packaged as an executable file (`.exe`), allowing easy use without requiring Python installation.

## Features

- **Folder Selection**: Choose a folder to convert all `.docx` files to PDF.
- **Conversion Progress Tracking**: Displays real-time updates during conversion.
- **Microsoft Word Process Handling**: Checks and offers to terminate running instances of Microsoft Word (`WINWORD.exe`) for smoother conversions.
- **Custom Styling**: Enhanced UI elements and styles with `ttk` module of `tkinter`.
- **Error Handling**: Integrated `StderrRedirector` for capturing and displaying error messages.
- **Termination Countdown**: Countdown displayed post-conversion before automatic closure.

## How It Works

1. **Interface Initialization**: Launching the application initiates the `tkinter` GUI. The main window is initially hidden.

2. **Folder Selection and Validation**: User selects a folder for conversion. The application checks for and can terminate running Microsoft Word processes.

3. **Conversion Process**: Runs `convert_documents` in a separate thread using `docx2pdf`. `StderrRedirector` captures and displays progress.

4. **Status Updates**: The GUI is continuously updated with conversion progress or error messages.

5. **Completion and Termination**: Post-conversion, a countdown leads to the application's automatic closure.

## Usage

1. **Prerequisites**: No prerequisites for the `.exe` version. For Python script, requires Python with `tkinter`, `docx2pdf`, `threading`, `queue`, `sys`, `re`, `time`, and `psutil`.
2. **Starting the Application**: Run the `.exe` file to launch the application.
3. **Folder Selection**: Choose a folder containing `.docx` files for conversion.
4. **Monitor Conversion**: Observe the progress in the status window.
5. **Completion**: The application closes automatically after conversion.

## Notes

- Designed primarily for Windows, considering the specific check for `WINWORD.exe`.
- Advisable to close Microsoft Word before running the application to prevent conflicts.
    - The program will check for WINWORD.exe and prompt to close if necessary

