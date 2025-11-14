#!/usr/bin/env python3
"""
Convert Word (.docx) and PowerPoint (.pptx) files to Markdown format.

This script recursively searches directories for Office documents and converts
them to Markdown using the MarkItDown library with optional LLM-powered image descriptions.
"""

import argparse
import logging
import os
import sys
from pathlib import Path
from typing import List, Tuple
from datetime import datetime

from dotenv import load_dotenv
from markitdown import MarkItDown
from tqdm import tqdm

# Load environment variables from .env file
load_dotenv()

# Configure logging
LOG_FILE = 'conversion_errors.log'
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)


def setup_markitdown(use_llm: bool = True) -> MarkItDown:
    """
    Initialize the MarkItDown converter with optional LLM integration.

    Args:
        use_llm: Whether to enable LLM-powered image descriptions

    Returns:
        Configured MarkItDown instance
    """
    if use_llm:
        api_key = os.getenv('OPENAI_API_KEY')
        model = os.getenv('OPENAI_MODEL', 'gpt-5')

        if not api_key:
            logger.warning("OPENAI_API_KEY not found in environment. Image descriptions will be disabled.")
            logger.warning("To enable, copy .env to .env and add your API key.")
            return MarkItDown(enable_plugins=False)

        try:
            from openai import OpenAI
            client = OpenAI(api_key=api_key)
            logger.info(f"LLM integration enabled using model: {model}")
            return MarkItDown(llm_client=client, llm_model=model)
        except Exception as e:
            logger.error(f"Failed to initialize OpenAI client: {e}")
            logger.warning("Falling back to conversion without LLM features.")
            return MarkItDown(enable_plugins=False)
    else:
        logger.info("LLM integration disabled.")
        return MarkItDown(enable_plugins=False)


def find_office_files(directory: Path, extensions: Tuple[str, ...] = ('.docx', '.pptx')) -> List[Path]:
    """
    Recursively find all Office files in the given directory.

    Args:
        directory: Root directory to search
        extensions: Tuple of file extensions to search for

    Returns:
        List of Path objects for found files
    """
    files = []
    for ext in extensions:
        files.extend(directory.rglob(f'*{ext}'))

    # Filter out temporary Office files (start with ~$)
    files = [f for f in files if not f.name.startswith('~$')]

    return sorted(files)


def normalize_to_ascii(text: str) -> str:
    """
    Normalize Unicode characters to their basic ASCII equivalents.

    Args:
        text: Input text with potential Unicode characters

    Returns:
        Text with Unicode characters replaced by ASCII equivalents
    """
    # Define character replacements
    replacements = {
        # Curly quotes to straight quotes
        '\u2018': "'",  # Left single quotation mark
        '\u2019': "'",  # Right single quotation mark
        '\u201A': "'",  # Single low-9 quotation mark
        '\u201B': "'",  # Single high-reversed-9 quotation mark
        '\u201C': '"',  # Left double quotation mark
        '\u201D': '"',  # Right double quotation mark
        '\u201E': '"',  # Double low-9 quotation mark
        '\u201F': '"',  # Double high-reversed-9 quotation mark
        # Dashes
        '\u2013': '-',  # En dash
        '\u2014': '--', # Em dash
        '\u2015': '--', # Horizontal bar
        # Spaces
        '\u00A0': ' ',  # Non-breaking space
        '\u2000': ' ',  # En quad
        '\u2001': ' ',  # Em quad
        '\u2002': ' ',  # En space
        '\u2003': ' ',  # Em space
        '\u2004': ' ',  # Three-per-em space
        '\u2005': ' ',  # Four-per-em space
        '\u2006': ' ',  # Six-per-em space
        '\u2007': ' ',  # Figure space
        '\u2008': ' ',  # Punctuation space
        '\u2009': ' ',  # Thin space
        '\u200A': ' ',  # Hair space
        # Other punctuation
        '\u2026': '...', # Horizontal ellipsis
        '\u2022': '*',   # Bullet
        '\u2023': '>',   # Triangular bullet
        '\u2032': "'",   # Prime
        '\u2033': '"',   # Double prime
        '\u2035': "'",   # Reversed prime
        '\u2036': '"',   # Reversed double prime
    }

    # Apply replacements
    for unicode_char, ascii_char in replacements.items():
        text = text.replace(unicode_char, ascii_char)

    return text


def convert_file(md: MarkItDown, input_file: Path, output_file: Path) -> bool:
    """
    Convert a single file to Markdown.

    Args:
        md: MarkItDown instance
        input_file: Path to input file
        output_file: Path to output Markdown file

    Returns:
        True if conversion successful, False otherwise
    """
    try:
        logger.info(f"Converting: {input_file}")
        result = md.convert(str(input_file))

        # Normalize Unicode characters to ASCII
        normalized_content = normalize_to_ascii(result.text_content)

        # Write the converted Markdown to file
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(normalized_content)

        logger.info(f"Successfully converted to: {output_file}")
        return True

    except Exception as e:
        logger.error(f"Failed to convert {input_file}: {str(e)}")
        return False


def main():
    """Main execution function."""
    parser = argparse.ArgumentParser(
        description='Convert Word and PowerPoint files to Markdown format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s                                    # Convert files in current directory
  %(prog)s --directory ./documents            # Convert files in specific directory
  %(prog)s --skip-images                      # Convert without LLM image descriptions
  %(prog)s --directory ./docs --extensions .docx  # Convert only Word documents
        """
    )

    parser.add_argument(
        '--directory', '-d',
        type=str,
        default='.',
        help='Directory to search for files (default: current directory)'
    )

    parser.add_argument(
        '--skip-images',
        action='store_true',
        help='Disable LLM-powered image descriptions to save API costs'
    )

    parser.add_argument(
        '--extensions', '-e',
        nargs='+',
        default=['.docx', '.pptx'],
        help='File extensions to convert (default: .docx .pptx)'
    )

    args = parser.parse_args()

    # Validate directory
    directory = Path(args.directory).resolve()
    if not directory.exists():
        logger.error(f"Directory does not exist: {directory}")
        sys.exit(1)

    if not directory.is_dir():
        logger.error(f"Path is not a directory: {directory}")
        sys.exit(1)

    # Initialize converter
    logger.info("="*70)
    logger.info("MarkItDown Converter - Document to Markdown Conversion")
    logger.info("="*70)
    logger.info(f"Start time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logger.info(f"Searching directory: {directory}")
    logger.info(f"File extensions: {', '.join(args.extensions)}")

    md = setup_markitdown(use_llm=not args.skip_images)

    # Find all Office files
    logger.info("Scanning for files...")
    extensions = tuple(ext if ext.startswith('.') else f'.{ext}' for ext in args.extensions)
    files = find_office_files(directory, extensions)

    if not files:
        logger.warning(f"No files with extensions {extensions} found in {directory}")
        sys.exit(0)

    logger.info(f"Found {len(files)} file(s) to convert")

    # Convert files with progress bar
    successful = 0
    failed = 0

    with tqdm(total=len(files), desc="Converting files", unit="file") as pbar:
        for input_file in files:
            # Create output filename (same location, .md extension)
            output_file = input_file.with_suffix('.md')

            # Update progress bar description
            pbar.set_description(f"Converting {input_file.name}")

            # Convert file
            if convert_file(md, input_file, output_file):
                successful += 1
            else:
                failed += 1

            pbar.update(1)

    # Print summary
    logger.info("="*70)
    logger.info("Conversion Summary")
    logger.info("="*70)
    logger.info(f"Total files processed: {len(files)}")
    logger.info(f"Successful conversions: {successful}")
    logger.info(f"Failed conversions: {failed}")
    logger.info(f"End time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    if failed > 0:
        logger.warning(f"Check {LOG_FILE} for error details")

    logger.info("="*70)


if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        logger.warning("\nConversion interrupted by user")
        sys.exit(130)
    except Exception as e:
        logger.error(f"Unexpected error: {e}", exc_info=True)
        sys.exit(1)
