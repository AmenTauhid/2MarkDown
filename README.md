# Document to Markdown Converter

A Python script that recursively converts Word (.docx) and PowerPoint (.pptx) files to Markdown format using Microsoft's MarkItDown library.

## Features

- **Recursive Directory Scanning**: Automatically finds all Word and PowerPoint files in a directory tree
- **LLM-Powered Image Descriptions**: Optional OpenAI integration for generating descriptions of images in documents
- **Progress Tracking**: Visual progress bars showing conversion status
- **Error Logging**: Comprehensive error logging to track failed conversions
- **Flexible Output**: Saves Markdown files alongside original documents
- **Batch Processing**: Efficiently handles multiple files in one run

## Prerequisites

- Python 3.12.12 (or Python 3.10+)
- Conda environment (recommended)
- OpenAI API key (optional, for image descriptions)

## Installation

### 1. Activate your Conda environment

```bash
conda activate your_env_name
```

### 2. Install dependencies

```bash
pip install -r requirements.txt
```

### 3. Configure OpenAI API key (optional)

If you want LLM-powered image descriptions:

1. Copy the example environment file:
   ```bash
   cp .env .env
   ```

2. Edit `.env` and add your OpenAI API key:
   ```
   OPENAI_API_KEY=sk-your-actual-api-key-here
   ```

3. Get your API key from: https://platform.openai.com/api-keys

> Note: If you skip this step, the script will still work but won't generate descriptions for images.

## Usage

### Basic Usage

Convert all Word and PowerPoint files in the current directory:

```bash
python convert_to_markdown.py
```

### Specify a Directory

Convert files in a specific directory:

```bash
python convert_to_markdown.py --directory ./documents
```

or using the short form:

```bash
python convert_to_markdown.py -d ./path/to/files
```

### Skip Image Descriptions

To save API costs and speed up conversion, disable LLM image descriptions:

```bash
python convert_to_markdown.py --skip-images
```

### Convert Only Specific File Types

Convert only Word documents:

```bash
python convert_to_markdown.py --extensions .docx
```

Convert only PowerPoint presentations:

```bash
python convert_to_markdown.py --extensions .pptx
```

### Combine Options

```bash
python convert_to_markdown.py --directory ./documents --skip-images --extensions .docx
```

## Command-Line Arguments

| Argument | Short | Default | Description |
|----------|-------|---------|-------------|
| `--directory` | `-d` | `.` (current) | Directory to search for files |
| `--skip-images` | - | `False` | Disable LLM-powered image descriptions |
| `--extensions` | `-e` | `.docx .pptx` | File extensions to convert |

## Output

- Markdown files (`.md`) are saved in the same location as the original files
- Original files are not modified or deleted
- Conversion errors are logged to `conversion_errors.log`
- Progress is displayed in the terminal with a progress bar

## Example Output

```
======================================================================
MarkItDown Converter - Document to Markdown Conversion
======================================================================
Start time: 2025-01-11 14:30:00
Searching directory: C:\Users\Documents
File extensions: .docx, .pptx
LLM integration enabled using model: gpt-4o
Scanning for files...
Found 15 file(s) to convert
Converting files: 100%|██████████████████████| 15/15 [00:45<00:00,  3.00s/file]
======================================================================
Conversion Summary
======================================================================
Total files processed: 15
Successful conversions: 14
Failed conversions: 1
End time: 2025-01-11 14:30:45
Check conversion_errors.log for error details
======================================================================
```

## Project Structure

```
2MarkDown/
├── convert_to_markdown.py    # Main conversion script
├── requirements.txt           # Python dependencies
├── .env.example              # Template for environment variables
├── .env                      # Your API keys (create from .env.example)
├── README.md                 # This file
└── conversion_errors.log     # Error log (created on first run)
```

## What Gets Converted?

The script converts various document elements to Markdown:

### Word Documents (.docx)
- Headings and text formatting (bold, italic)
- Lists (ordered and unordered)
- Tables
- Links
- Images (with optional AI-generated descriptions)
- Document structure and hierarchy

### PowerPoint Presentations (.pptx)
- Slide titles and content
- Text boxes and shapes
- Lists and formatting
- Speaker notes
- Images (with optional AI-generated descriptions)
- Maintains slide order

## Troubleshooting

### "OPENAI_API_KEY not found in environment"

This warning appears if you haven't set up the `.env` file. The script will still work but won't generate image descriptions. To fix:

1. Copy `.env.example` to `.env`
2. Add your OpenAI API key to the `.env` file

### "No files found"

Make sure:
- You're in the correct directory
- Files have the correct extensions (`.docx` or `.pptx`)
- Files are not temporary Office files (starting with `~$`)

### Conversion Failures

Check `conversion_errors.log` for detailed error messages. Common issues:
- Corrupted or password-protected files
- Files in use by another application
- Insufficient permissions

### Import Errors

If you get module import errors, ensure all dependencies are installed:

```bash
pip install -r requirements.txt
```

## Cost Considerations

If using LLM image descriptions:
- Each image requires an API call to OpenAI
- Costs vary based on model (default: gpt-4o)
- Use `--skip-images` flag to avoid API costs

## Supported Python Versions

- Python 3.10+
- Tested with Python 3.12.12

## Dependencies

- `markitdown` - Microsoft's document converter
- `openai` - OpenAI API client
- `tqdm` - Progress bar library
- `python-dotenv` - Environment variable management

## License

This project uses the MarkItDown library developed by Microsoft.

## Support

For issues with:
- **This script**: Check `conversion_errors.log` or review the code
- **MarkItDown library**: Visit https://github.com/microsoft/markitdown
- **OpenAI API**: Visit https://platform.openai.com/docs

## Additional Resources

- [MarkItDown GitHub Repository](https://github.com/microsoft/markitdown)
- [OpenAI API Documentation](https://platform.openai.com/docs)
- [Markdown Guide](https://www.markdownguide.org/)
