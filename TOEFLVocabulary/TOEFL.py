# libraries
from bs4 import BeautifulSoup
import requests
import lxml
import xlwt

# initializes Excel spreadsheet file (.xls), defines different styles for header and content
def initXls():
    # open spreadsheet file
    book = xlwt.Workbook()
    sheet = book.add_sheet("TOEFL", True)
    
    # header format
    headingStyle = xlwt.XFStyle()
    headingStyle.font.bold = True
    headingStyle.font.height = 14 * 20
    headingStyle.alignment.wrap = 1
    headingStyle.alignment.vert = xlwt.Alignment.VERT_CENTER
    headingStyle.alignment.horz = xlwt.Alignment.HORZ_CENTER

    # content format
    contentStyle = xlwt.XFStyle()
    contentStyle.font.height = 12 * 20
    contentStyle.alignment.wrap = 1
    contentStyle.alignment.vert = xlwt.Alignment.VERT_CENTER
    contentStyle.alignment.horz = xlwt.Alignment.HORZ_LEFT

    # sheet formatting
    sheet.row(0).height_mismatch = True
    sheet.row(0).height = 30 * 30
    sheet.col(0).width = 20 * 367
    sheet.col(1).width = 80 * 367
    sheet.col(2).width = 80 * 367

    # print header
    sheet.write(0, 0, "Words", headingStyle)
    sheet.write(0, 1, "Definitions", headingStyle)
    sheet.write(0, 2, "Examples", headingStyle)

    return(book, sheet, contentStyle)

# data is extracted from url response, loaded into dictionary, and printed
def scrape(soup):
    # grab master list of word entries
    entries = soup.find_all('li', {'class' : 'entry learnable'})

    # declare empty dictionary and iterate through
    dict = {}
    for word in entries:
        value = []
        key = word.find('a', {'class' : 'word dynamictext'}).text.strip()
        definition = word.find('div', {'class' : 'definition'}).text.strip()
        example = word.find('div', {'class' : 'example'})

        # if example sentence has source, it is deleted
        if(len(example.find_all('a', {'class' : 'source'})) > 0):
            example.a.decompose()

        # cleaning up scraped example sentence (some contain '\n' in the middle)
        example = example.text.strip()
        example = example.replace("\n", "")

        value.append(definition)
        value.append(example)
        
        dict[key] = value

    printDict(dict)

    return (dict)

# prints dictionary
def printDict(dict):
    for key in sorted(dict):
        print(key)
        for item in dict.get(key):
            print(item)
        print("\n")

# writes to spreadsheet and saves book
def writeXls(book, sheet, contentStyle, words):
    rowIndex = 1
    for key in sorted(words):
        # formatting for each row
        sheet.row(rowIndex).height_mismatch = True
        sheet.row(rowIndex).height = 30 * 30

        # write key to leftmost column
        sheet.write(rowIndex, 0, key, contentStyle)

        # iterate through each key's corresponding value, which is a list of both a definition and example sentence
        colIndex = 1
        for item in words.get(key):
            sheet.write(rowIndex, colIndex, item, contentStyle)
            colIndex += 1

        # update row index
        rowIndex += 1
    
    book.save("TOEFL Vocabulary (" + str(rowIndex - 1) + " Words).xls")

def main():
    url = "https://www.vocabulary.com/lists/52473"
    response = requests.get(url, "html.parser")
    book, sheet, contentStyle = initXls()
    words = scrape(BeautifulSoup(response.text, "lxml"))
    writeXls(book, sheet, contentStyle, words)

if __name__ == "__main__":
  main()