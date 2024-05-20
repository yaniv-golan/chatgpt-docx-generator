# chatgpt-docx-generator

`chatgpt-docx-generator` is a Python script designed to convert ChatGPT output into `.docx` documents, preserving headings, bullet points, code blocks, and other formatting elements. This tool is ideal for generating well-structured Word documents from markdown-like text files, making it easier to share and present structured content.

## Features

- Convert ChatGPT output to `.docx` format
- Preserve headings, bullet points, and code blocks
- Command line arguments for input and output files
- Clipboard integration for input and output
- Customizable heading levels

## Usage

To use the script, run the following command:

```sh
python scripts/convert_to_docx.py --input-file <input_file> --output-file <output_file> --base-heading-level <level> [--input-clipboard] [--output-clipboard] [--debug]
```

### Parameters

- `--input-file`: The input text file.
- `--output-file`: The output `.docx` file.
- `--base-heading-level`: The base heading level (default: 1).
- `--input-clipboard`: Take input from the clipboard.
- `--output-clipboard`: Put output in the clipboard.
- `--debug`: Enable debug logging.

## Examples

### Convert from file to file

```sh
python scripts/convert_to_docx.py --input-file test/test1.txt --output-file test/out1.docx --base-heading-level 3 --debug
```

### Convert from clipboard to file

```sh
python scripts/convert_to_docx.py --input-clipboard --output-file test/out1.docx --base-heading-level 3 --debug
```

### Convert from file to clipboard

```sh
python scripts/convert_to_docx.py --input-file test/test1.txt --output-clipboard --base-heading-level 3 --debug
```

### Convert from clipboard to clipboard

```sh
python scripts/convert_to_docx.py --input-clipboard --output-clipboard --base-heading-level 3 --debug
```

## Installation

1. **Clone the repository:**

   ```sh
   git clone https://github.com/yaniv-golan/chatgpt-docx-generator.git
   cd chatgpt-docx-generator
   ```

2. **Install dependencies:**

   Ensure you have Python and the required libraries installed:

   ```sh
   pip install python-docx pyperclip
   ```

## Project Structure

```
chatgpt-docx-generator/
├── scripts/
│   └── convert_to_docx.py
├── LICENSE
├── README.md
└── .gitignore
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request with any improvements or bug fixes.

## Acknowledgments

- Thanks to the OpenAI team for developing ChatGPT.
- Thanks to the contributors of the `python-docx` and `pyperclip` libraries.
