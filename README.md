# Directory Content Generator

Are you tired of waiting on Windows search function to find files in subdirectories? I originally created this script so I had a quick reference for my 3D model STLs, but then modified it so it can be used for any directory on your Windows machine. It creates an XLSX file containing information about the files in the specified directory.

## Usage

Pass it a directory to scan and an output filename:

![Command Line Interface](./images/cmd.png)

The following information will be scraped and stored in a spreadsheet:

- Relative file path
- Filename
- Filetype
- Clickable URL
- Size

![Spreadsheet Screenshot](./images/output.png)

You can optionally pass the `-i` or `--images` flag to also embed all images into the spreadsheet:

![Spreadsheet Screenshot](./images/output_with_images.png)
