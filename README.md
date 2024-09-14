# MBOX to PST Converter Using Outlook

This Python script converts MBOX files into PST format using Microsoft Outlook's COM interface. It is designed to handle large MBOX files efficiently with batch processing, retries, and attachment support.

## Features

- Converts MBOX emails to PST using the Microsoft Outlook COM interface.
- Supports attachments without needing to save them on disk.
- Handles email retries for robust error handling.
- Batch processing to avoid memory overload and improve performance.
- Resumable functionality: skips already processed emails using email hashing.
- Logs detailed information and errors to help with debugging.

## Prerequisites

- **Windows OS**: The script uses the Outlook COM interface, which requires Windows.
- **Microsoft Outlook**: You must have Outlook installed and configured.
- **Python 3.x**: Make sure you have Python 3.x installed.

## Installation

1. Clone the repository or download the script.
2. Install the required dependencies using the `requirements.txt` file:

   ```bash
   pip install -r requirements.txt
