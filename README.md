# Rename-ByMetadata.ps1

A forensic-grade PowerShell script to rename files using embedded metadata, designed for recovery sets and chaotic folders. Supports documents, images, audio, and video. Features preview-first workflow, confidence scoring, batch processing, and robust error handling for long runs.

## Features
- Scans folders recursively or non-recursively
- Extracts deep metadata from supported file types
- Generates safe, meaningful filenames
- Preview mode with TXT/CSV reports
- Batch processing for large sets
- Retry and error logging
- Skips files already renamed (_renamed suffix)
- Interactive and non-interactive modes

## Usage
```powershell
# Basic preview
./Rename-ByMetadata.ps1 -Path "C:\TargetFolder"

# Apply renames
./Rename-ByMetadata.ps1 -Path "C:\TargetFolder" -Apply

# Custom batch size and timeout
./Rename-ByMetadata.ps1 -Path "C:\TargetFolder" -BatchSize 500 -PerFileTimeoutSeconds 20
```

See script help for all options and examples.

## Requirements
- Windows PowerShell
- No external dependencies for most metadata extraction
- Microsoft Word required for some DOC/RTF scenarios

## License
MIT

## Contributing
Pull requests and issues welcome!

## Author
David Hewitt / RemarkablePC

---
For more details, see the script header and comments.
