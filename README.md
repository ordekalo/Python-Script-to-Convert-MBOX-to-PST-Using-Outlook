
# MBOX to PST Converter Using Outlook

This Python script converts MBOX files into PST format using the Microsoft Outlook COM interface. It supports processing large MBOX files efficiently, including attachments, retry logic, and batch processing to avoid memory overload. The script is designed to be robust, with detailed error logging, progress tracking, and directory management to ensure smooth operation.

## Features

- **MBOX to PST Conversion**: Converts emails from an MBOX file to a PST file.
- **Attachment Handling**: Directly adds email attachments without saving them to disk.
- **Retry Logic**: Automatically retries email processing in case of errors, ensuring robustness.
- **Batch Processing**: Emails are processed in batches to avoid memory overload and improve performance.
- **Graceful Error Handling**: Logs all errors and skips already processed emails using unique email hashing.
- **Directory Management**: Automatically creates directories for PST files if they don't exist.
- **Progress Logging**: Detailed logging of processing steps, including success and error information.

## Requirements

- **Operating System**: Windows
- **Microsoft Outlook**: You must have Outlook installed and configured on your system.
- **Python 3.x**: The script requires Python 3.x.

## Installation

### 1. Clone or Download the Repository

Clone the repository using Git or download it as a ZIP and extract it to your local machine:

```bash
git clone https://github.com/yourusername/mbox_to_pst_converter.git
```

### 2. Install Python Dependencies

Make sure you have Python 3.x installed. Install the required Python packages using `pip` and the provided `requirements.txt` file:

```bash
pip install -r requirements.txt
```

### 3. Required Dependencies

You can install the dependencies manually if necessary:

```bash
pip install tqdm pywin32 timeout-decorator
```

These packages are required for:
- **`tqdm`**: Displaying progress bars for long-running operations.
- **`pywin32`**: Interfacing with Microsoft Outlook using the Windows COM API.
- **`timeout-decorator`**: Adding timeout capabilities to specific operations.

## Usage

### Command-Line Usage

You can run the script by providing the path to the MBOX file, the output folder for attachments, and an optional path for the PST file. If no PST file path is provided, it defaults to `emails.pst` in the current working directory.

```bash
python convert.py <mbox_file> <output_folder> --pst_file <pst_file_path>
```

### Arguments:

1. **`mbox_file`**: The path to the MBOX file that contains the emails you want to convert to PST.
2. **`output_folder`**: The directory where attachments will be temporarily saved (if needed).
3. **`--pst_file`**: (Optional) The path where the PST file will be saved. If not provided, it defaults to `emails.pst`.

### Example Command:

```bash
python convert.py file.mbox D:\output --pst_file D:\output\emails.pst
```

In this example:
- `file.mbox` is the MBOX file to be converted.
- `D:\output` is the folder where attachments (if any) will be saved temporarily.
- `D:\output\emails.pst` is the PST file that will be generated.

## Detailed Logging

The script logs all progress and errors to a file called `import_log.txt`. You can monitor this file to track the progress of the conversion and check for any issues.

### Logging Information Includes:
- The number of emails processed.
- The creation of the PST file and its directory.
- Any errors encountered during the process (e.g., email processing errors, directory creation errors).
- Retry attempts if processing an email fails.

### Log Location

The log file is created in the same directory as the script:
```text
import_log.txt
```

## Features in Detail

### 1. **Error Handling and Retry Logic**:
   - If an email fails to be processed (due to a temporary error), the script retries the operation up to 3 times. If the operation fails after 3 retries, the email is skipped, and an error is logged.

### 2. **Attachment Handling**:
   - Attachments are saved directly to the Outlook mail item without being written to disk first. This optimizes performance and prevents cluttering the file system.

### 3. **Batch Processing**:
   - The script processes emails in batches of 500 (configurable) to avoid memory overload when dealing with large MBOX files. Each batch of emails is processed concurrently using Python’s `concurrent.futures` library.

### 4. **Graceful Shutdown & Resuming**:
   - The script uses a unique hash to identify each email. If the script is interrupted or fails, it will skip emails that have already been processed when restarted.

### 5. **Directory Creation**:
   - If the provided path for the PST file doesn’t exist, the script automatically creates the necessary directories.

### 6. **Progress Tracking**:
   - The `tqdm` library is used to display a progress bar while processing emails, showing how many emails have been processed and how many are left.

## Customization

### 1. **Batch Size**:
   - You can change the `batch_size` parameter inside the script if you want to process a different number of emails per batch (default is 500).

### 2. **Retry Count**:
   - The default number of retries for a failed email processing attempt is 3. You can modify this value by changing the `retries` parameter in the `process_email_with_retry()` function.

### 3. **PST File Location**:
   - If no `--pst_file` argument is provided, the PST file is saved as `emails.pst` in the current working directory. You can specify a custom directory for the PST file using the `--pst_file` argument.

### 4. **Logging**:
   - All errors and key operations are logged to `import_log.txt`. You can review this log for information about any errors that occurred during the process.

## Troubleshooting

### 1. **Outlook Errors**:
   If you encounter errors related to Outlook, ensure that Outlook is installed, configured correctly, and not being used by other processes during the script's execution.

### 2. **Permission Issues**:
   Ensure that you have permission to write to the output directory and PST file location. If the directory for the PST file doesn't exist, the script will create it automatically.

### 3. **Large MBOX Files**:
   The script is designed to handle large MBOX files using batch processing to prevent memory overload. However, if you encounter memory issues, try reducing the batch size.

### 4. **Missing Attachments**:
   If attachments are not appearing in the final PST file, ensure that the email parts are correctly formatted in the MBOX file. The script expects standard MIME-formatted attachments.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

---

## Example Usage

```bash
# Basic usage with a default PST file name
python convert.py file.mbox D:\output

# Usage with a specified PST file path
python convert.py file.mbox D:\output --pst_file D:\myPSTs\converted_emails.pst
```

---

### Additional Notes

1. **Test Small MBOX Files First**: If you're running the script for the first time, it's recommended to test it with a smaller MBOX file before processing large files.

2. **Backup Important Data**: Always back up important data before running large-scale conversions.

3. **PST File Limits**: PST files have size limits based on the version of Outlook you're using. Ensure that the size of your final PST file doesn't exceed these limits (e.g., 20 GB for Outlook 2003, 50 GB for Outlook 2010 and later).

---
