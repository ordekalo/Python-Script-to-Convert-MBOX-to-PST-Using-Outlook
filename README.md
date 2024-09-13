# MBOX to PST Converter Using Outlook COM Interface

This Python script allows you to convert MBOX files to PST format by using the Microsoft Outlook COM interface. It reads an MBOX file, extracts the emails, and imports them into a new PST file, with a progress bar showing the conversion progress.

## Features

- Converts emails from an MBOX file into a PST file.
- Utilizes the Microsoft Outlook COM interface to perform the conversion.
- Automatically creates an "Inbox" folder in the PST file if it doesn't exist.
- Handles both plain text and HTML email bodies.
- Displays a progress bar during conversion using `tqdm`.

## Requirements

- **Windows Operating System**: This script works only on Windows, as it uses the COM interface for Microsoft Outlook.
- **Microsoft Outlook**: You must have Outlook installed and configured on your machine.
- **Python 3.x**
- **pywin32** package: This is required for interacting with Outlook.
- **tqdm** package: This package is used to display the progress bar.

## Installation

### 1. Install Python

Make sure you have Python 3.x installed on your system. If not, you can download it from [Python's official website](https://www.python.org/downloads/).

### 2. Install the required packages

To interact with Outlook and display the progress bar, you'll need the following packages:

```bash
pip install -r requirements.txt
