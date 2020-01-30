#---------------------- IMPORTING REQUIRED MODULES ------------------------#
import xlrd
import xlwt
from xlutils.copy import copy
import requests
from bs4 import BeautifulSoup
import re


#---------------------- DEFINING THE REQUIRED FUNCTIONS ------------------------#
# getUrls : Gets the https://www.biblegateway.com/ URL that we need to scrape
# getVerse : Scrapes the given https://www.biblegateway.com/ URL to get the verse
# createStyles : Create styles for the cells
# writeToExcel : Writes the Verse URL and the verse to the Execel sheet
#-------------------------------------------------------------------------------#
## Getting the URLs
def getUrls(read_sheet, row):
    verse_url = "https://www.biblegateway.com/passage/?search=" + str(read_sheet.cell_value(row, 0)) + "+" + str(int(read_sheet.cell_value(row, 1))) + "%3A" + str(int(read_sheet.cell_value(row, 2))) + "&version=" + str(read_sheet.cell_value(row,4))
    verse_url_formula = 'HYPERLINK(CONCATENATE("https://www.biblegateway.com/passage/?search=",A'+ str(row+1) +',"+", B'+ str(row+1) +',"%3A", C'+ str(row+1) +',"&version=",E' + str(row+1) + '))'
    return (verse_url, verse_url_formula)


## Scrapping the Bible Verse from biblegateway
def getVerse(verse_url):
    page = requests.get(verse_url)
    soup = BeautifulSoup(page.text, 'html.parser')
    tag = soup.find_all('span', id = lambda x: x and x.startswith("en-"))
    verse = re.sub(r"\d+", "", tag[0].get_text()).strip()
    return verse


## Creating Styles for the Cells
def createStyles():
    ## Heading -- White text on a blue background
    h1_style = xlwt.easyxf('font: name Times New Roman, bold on, height 280, color black;'
                            'borders: left thin, right thin, top thin, bottom thin;'
                            'pattern: pattern fine_dots, fore_color white, back_color orange;'
                            'align: vertical center, horizontal center;')

    ## Style for the cells
    hyperlink_style = xlwt.easyxf('font: name Times New Roman, height 280, color blue, underline on;'
                            'borders: top_color gray25, bottom_color gray25, right_color gray25, left_color gray25, left thin, right thin, top thin, bottom thin;'
                            'pattern: pattern solid, fore_colour white;')

    return h1_style, hyperlink_style

    
## Writing Verse URL and Verse to the Excel Sheet
def writeToExcel(read_sheet, write_sheet):

    ## Setting a column widths
    col_width = 256 * 30    # 30 chars
    write_sheet.col(0).width = col_width
    write_sheet.col(1).width = col_width - 20
    write_sheet.col(2).width = col_width - 20
    write_sheet.col(3).width = col_width
    write_sheet.col(4).width = col_width - 20
    write_sheet.col(5).width = col_width * 3
    write_sheet.col(6).width = col_width * 3

    ## Formatting the Headings in the Excel Sheet
    h1_style, hyperlink_style = createStyles()
    write_sheet.write(0, 0, 'BOOK', h1_style)
    write_sheet.write(0, 1, 'CHAPTER', h1_style)
    write_sheet.write(0, 2, 'VERSE', h1_style)
    write_sheet.write(0, 3, 'REFERENCE', h1_style)
    write_sheet.write(0, 4, 'VERSION', h1_style)
    write_sheet.write(0, 5, 'VERSE URL', h1_style)
    write_sheet.write(0, 6, 'VERSE', h1_style)

    for row in range(1, read_sheet.nrows):
        if read_sheet.cell(row, 0).value:

            ## Get URLs
            verse_url, verse_url_formula = getUrls(read_sheet, row)

            ## Writing the URL to the cell
            write_sheet.write(row, 5, xlwt.Formula(verse_url_formula), hyperlink_style)

            ## Getting and writing entire verse
            verse = getVerse(verse_url)
            write_sheet.write(row, 6, verse)


#---------------------- DEFINING THE MAIN FUNCTION -----------------------------#
def main():
    ## Opening an existing workbook to write into
    xlrd_book = xlrd.open_workbook("bible_memory_verses.xls", formatting_info=True)
    xlwt_workbook = copy(xlrd_book)

    ## Getting the worksheet by index to write
    write_sheet = xlwt_workbook.get_sheet(0)

    ## Getting the worksheet by index to read
    read_sheet = xlrd_book.sheet_by_index(0)
    writeToExcel(read_sheet, write_sheet)

    ## Saving the updated Excel workbook
    xlwt_workbook.save("bible_memory_verses.xls")


#---------------------- EXECUTING THE MAIN FUNCTION -----------------------------#
if __name__ == "__main__":
    main()

