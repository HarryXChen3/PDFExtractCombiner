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

### Mode Usage
- .xlsx & .pdf combine (monthly STR & RPM; STR binder report)
  - When prompted to select the mode, select the mode above
  - If `Unexpected duplicate files exist recursively/directly under {working-dir}. Please
  remove/rename them to continue.`, remove duplicate files (files with the same root name and same
  extension, i.e. `C:/duplicate.pdf` and `C:/test/duplicate.pdf`)
  - Check `Working Directory`, `Intended Output Directory`, `TMP Directory` (temp) directory
  - Ensure matched xlsx (Excel) and pdf file pairs count is what you expect
  - Ensure lone files count is what you expect (lone files are files that do not have a matching pdf/Excel file)
    - Lone files names/paths are displayed below this line; ex. `C:/lone_file_1.pdf, C:/lone_file_2.xlsx`
  - If `Directory {output-dir} already exists and is NOT empty` appears, confirm that files inside the directory might
  be overwritten
  - Answer `y` (default) if correct, `n` if not
  - `y` to start, `n` to cancel
  - Wait for progress bar ` 51%|█████     | 47/93 [00:55<00:52,  1.14s/it]`
    - `n%|bar| completed/total [time-spent<time-remaining, 1.14sec/iteration (or n-iteration/sec)]`
  - Wait for `Merge all newly combined .pdfs (n) into a single pdf?` and make choice
    - All newly merged pdfs are independent files distinguished by Excel pdf file pairs, you have the
    option to merge them all together as 1 `.pdf` file (in ascending order)
    - `y` to merge, `n` to ignore and continue
  - If merging as 1 pdf, wait for `Attempting merge of n .pdfs...`
    - Wait for `Wrote combined pdf to {pdf-path}.`
    - `Press [Enter] to exit...`
- .pdf combine (1st page; P&L First Pages)
  - When prompted to select the mode, select the mode above
  - Check `Working Directory`, `Intended Output Directory`, `TMP Directory` (temp) directory
  - Ensure matched pdf files count is what you expect
  - Answer `y` (default) if correct, `n` if not
  - `y` to start, `n` to cancel
  - Wait for '`Wrote combined pdf to {pdf-path}.`
  - `Press [Enter] to exit...`
## Building
1. Navigate to the project root directory `./`
2. Install `pyinstaller` via pip from PyPI - [GitHub](https://github.com/pyinstaller/pyinstaller)
3. Run `pyinstaller main.spec`