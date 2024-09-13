# MBOX to PST Converter Using Outlook COM Interface

This Python script allows you to convert MBOX files to PST format by using the Microsoft Outlook COM interface. It reads an MBOX file, extracts the emails, and imports them into a new PST file.

## Features

- Converts emails from an MBOX file into a PST file.
- Utilizes the Microsoft Outlook COM interface to perform the conversion.
- Automatically creates an "Inbox" folder in the PST file if it doesn't exist.
- Handles both plain text and HTML email bodies.

## Requirements

- **Windows Operating System**: This script works only on Windows, as it uses the COM interface for Microsoft Outlook.
- **Microsoft Outlook**: You must have Outlook installed and configured on your machine.
- **Python 3.x**
- **pywin32** package: This is required for interacting with Outlook.

## Installation

### 1. Install Python

Make sure you have Python 3.x installed on your system. If not, you can download it from [Python's official website](https://www.python.org/downloads/).

### 2. Install the `pywin32` package

To interact with Outlook, you'll need the `pywin32` package, which provides access to the COM interface. You can install it via pip:

```bash
pip install pywin32
