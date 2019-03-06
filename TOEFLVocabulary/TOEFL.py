# TODO
# handle apostrophes and multiple meanings better
# do commenting

# libraries
from bs4 import BeautifulSoup
import requests
import lxml
import xlwt

def scrape(soup):
    entries = soup.find('div', {'class' : 'entry-content'})
    words = []
    definitions = []
    sentences = []
    for entry in entries.find_all('tr'):
        if(len(entry.find_all('div', {'class' : 'example-blue'})) > 0):
            entry.div.decompose()

        word = entry.find_all('td')[0].text.strip()
        words.append(word)
        print(word)

        definition = entry.find_all('td')[1].text.strip()
        definitions.append(definition)
        print(definition)

        sentence = entry.find_all('td')[2].text.strip()
        sentences.append(sentence)
        print(sentence)

        print("\n")

    return (words, definitions, sentences)

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

    return(book, sheet, headingStyle, contentStyle)

def writeXls(book, sheet, headingStyle, contentStyle, words, definitions, sentences):
    index = 0
    sheet.col(0).width = 20 * 367
    sheet.col(1).width = 80 * 367
    sheet.col(2).width = 80 * 367
    while (index < len(words)):
        sheet.row(index).height_mismatch = True
        sheet.row(index).height = 30 * 30
        if(index == 0):
            sheet.write(index, 0, words[index], headingStyle)
            sheet.write(index, 1, definitions[index], headingStyle)
            sheet.write(index, 2, sentences[index], headingStyle)
        else:
            sheet.write(index, 0, words[index], contentStyle)
            sheet.write(index, 1, definitions[index], contentStyle)
            sheet.write(index, 2, sentences[index], contentStyle)
        index += 1
    book.save("TOEFL_Vocabulary.xls")

def main():
    url = "https://www.prepscholar.com/toefl/blog/toefl-vocabulary-list/"
    response = requests.get(url, "html.parser")
    book, sheet, headingStyle, contentStyle = initXls()
    words, definitions, sentences = scrape(BeautifulSoup(response.text, "lxml"))
    writeXls(book, sheet, headingStyle, contentStyle, words, definitions, sentences)

if __name__ == "__main__":
  main()