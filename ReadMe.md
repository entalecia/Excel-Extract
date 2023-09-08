# Extract Data From Excel

<img src="https://img.shields.io/badge/Python-green?style=flat&logo=python" />

#### This code extracts data from an excel sheets and presents it in a word document.


## Run
- Clone the repository `git clone https://github.com/entalecia/excel-extract`
- Install all requirements `pip install -r requirements.txt`
- Create a file **`doc.xlsx`** with your data
- Run `python main.py`

## Presentation

All the columns and it's data per row will be presented in one page seperated by page break. The Column Heading will be in **bold**. <br>
You can exclude some Column Headings by adding the heading's name *(case-sensitive)* in **titles_to_ignore** variable