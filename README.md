# PDFExtractCombiner

This is a tool developed to automate the monotone process of converting `.xlsx (Excel)` files into `.pdf` files,
then extracting specific pages and merging them into another `.pdf` file to create a merged `.pdf` file

## Requirements
You must run this on a **native Windows OS environment**, it will **NOT** work under any other OS.  
You must have **Excel** installed on your computer (it will be launched in the background).

All remaining dependencies are packaged within the `.zip` archive.

## Usage
1. Download the latest release from [here](https://github.com/HarryXChen3/PDFExtractCombiner/releases) 
2. Extract the `.zip` file and place the `./main` folder in the intended working directory
(where your `.xlsx` and `.pdf` files are)
3. Locate the `main.exe` file at `./main/main.exe`
4. Run `main.exe` and follow the prompts

## Building
1. Navigate to the project root directory `./`
2. Install `pyinstaller` via pip from PyPI - [GitHub](https://github.com/pyinstaller/pyinstaller)
3. Run `pyinstaller main.spec`