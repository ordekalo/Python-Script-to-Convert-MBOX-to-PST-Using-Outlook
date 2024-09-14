
# MBOX to PST Converter

This Python script converts emails from an MBOX file to a PST file using Microsoft Outlook on Windows. It is designed to efficiently handle large MBOX files, preserve email attachments, and provides robust error handling and progress tracking.

## Features

- **Batch Processing**: Processes emails in configurable batches to avoid memory overload.
- **Attachment Handling**: In-memory processing of attachments to minimize disk I/O.
- **Graceful Exit**: Supports signal handling (e.g., Ctrl+C) to gracefully shut down and save progress.
- **Checkpointing**: Save and resume progress after interruptions.
- **Auto-detect MBOX Files**: If no MBOX file is specified, the script will auto-detect `.mbox` files in the current directory.
- **Backup PST Files**: Automatically backs up existing PST files to avoid overwriting.
- **Multi-threading**: Email processing is parallelized for improved performance.
- **Configurable Workers**: Set the number of workers to control parallelization based on system resources.
- **Interactive Prompts**: Prevent accidental overwriting by confirming before replacing existing PST files.
- **Summary Report**: A detailed report is generated at the end, summarizing the number of emails processed and any failures.
- **Progress Bars**: Real-time progress bars are provided for email extraction and processing.

## Prerequisites

- **Microsoft Outlook**: Ensure Outlook is installed and configured on your Windows machine.
- **Python 3.x**: Install Python on your system.

### Required Python Packages

- `pywin32`: For interacting with Microsoft Outlook.
- `tqdm`: For displaying progress bars.
- `retrying`: For handling retryable errors with exponential backoff.

You can install the required packages using:

```bash
pip install -r requirements.txt
```

## Usage

```bash
python convert.py [output_folder] [--mbox_file <path_to_mbox>] [--pst_file <path_to_pst>] [--log-level <log_level>] [--batch-size <batch_size>] [--workers <num_workers>]
```

### Arguments

- `output_folder` (required): Directory where attachments will be saved.
- `--mbox_file`: Path to the MBOX file. If not provided, the script will auto-detect `.mbox` files in the current directory.
- `--pst_file`: Path to save the PST file. Default is `emails.pst` in the current directory.
- `--log-level`: Set the log verbosity (`DEBUG`, `INFO`, `WARNING`, `ERROR`, `CRITICAL`). Default is `INFO`.
- `--batch-size`: Number of emails to process in a batch. Default is 500.
- `--workers`: Number of parallel workers for processing emails. Default is the number of CPU cores.

### Examples

1. **Basic Usage**:
   ```bash
   python convert.py D:\output --pst_file D:\output\emails.pst
   ```

2. **With a Specific MBOX File**:
   ```bash
   python convert.py D:\output --mbox_file D:\mails\file.mbox --pst_file D:\output\emails.pst
   ```

3. **Verbose Logging (Debug Mode)**:
   ```bash
   python convert.py D:\output --log-level DEBUG
   ```

4. **Configuring Batch Size and Workers**:
   ```bash
   python convert.py D:\output --batch-size 1000 --workers 8
   ```

### Progress and Checkpointing

The script displays progress bars for email extraction and batch processing. It also saves progress after every batch, allowing you to resume the process if interrupted (e.g., system shutdown or Ctrl+C).

### Backup of Existing PST Files

Before overwriting an existing PST file, the script creates a backup in the same directory (e.g., `emails.pst.backup`).

### Graceful Exit

If interrupted (e.g., by pressing Ctrl+C), the script will save the progress and allow resuming from where it left off.

## License

This project is licensed under the MIT License.
