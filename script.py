from win32com import client
import os


def initialize():
    wordApp = client.Dispatch('Word.Application')
    wordApp.Visible = 1
    wordApp.DisplayAlerts = 0
    return wordApp


class DocumentStyle:
    def __init__(self, doc):
        self.document = doc
        self.paragraphs = doc.paragraphs
        self.paragraphStyles = self.getParagraphStyle()
        self.margins = self.getMargins()

    def getParagraphStyle(self):
        paragraphs = []
        for p in self.paragraphs:
            paragraphs.append(ParagraphStyle(p))
        return paragraphs

    def getMargins(self):
        margins = [self.document.PageSetup.LeftMargin, self.document.PageSetup.RightMargin,
                   self.document.PageSetup.TopMargin, self.document.PageSetup.BottomMargin]
        return margins

    def setMargins(self, m):
        self.document.PageSetup.LeftMargin = m[0]
        self.document.PageSetup.RightMargin = m[1]
        self.document.PageSetup.TopMargin = m[2]
        self.document.PageSetup.BottomMargin = m[3]


class ParagraphStyle:
    def __init__(self, paragraph):
        self.paragraph = paragraph
        self.name = str(paragraph.style)
        self.characterStyles = self.getCharacterStyle()
        # self.font = paragraph.style.font
        self.format = paragraph.style.paragraphformat

    def getCharacterStyle(self):
        characters = []
        for c in self.paragraph.Range.Characters:
            characters.append(CharacterStyle(c))
        return characters


class CharacterStyle:
    def __init__(self, character):
        self.character = character
        self.font = character.Font


def checkFormatting(testdocument, sampledocument):
    print('Text formatting:')
    checkStyle(testdocument, sampledocument)
    print('\nPage formatting:')
    checkMargins(testdocument, sampledocument)


def checkStyle(testdocument, sampledocument):
    usedStyles = []
    testlen = len(testdocument.paragraphs)
    samplelen = len(sampledocument.paragraphs)

    for par in range(testlen):
        currentStyle = testdocument.paragraphStyles[par].name
        if currentStyle not in usedStyles:
            usedStyles.append(currentStyle)

            if par == 0:
                print('\nTitle paragraph(s):')
                checkParagraphStyle(testdocument.paragraphStyles[par].format, sampledocument.paragraphStyles[par].format)
                for character in range(len(testdocument.paragraphStyles[par].characterStyles)):
                    checkCharacterStyle(testdocument.paragraphStyles[par].characterStyles[character].font,
                                    sampledocument.paragraphStyles[par].characterStyles[1].font)

            else:
                print('\nBody paragraph(s):')
                checkParagraphStyle(testdocument.paragraphStyles[par].format, sampledocument.paragraphStyles[1].format)
                for character in range(len(testdocument.paragraphStyles[par].characterStyles)):
                    checkCharacterStyle(testdocument.paragraphStyles[par].characterStyles[character].font,
                                    sampledocument.paragraphStyles[1].characterStyles[1].font)


def checkMargins(testdocument, sampledocument):
    wrongMarginsCount = 0
    for m in range(4):
        if testdocument.margins[m] != sampledocument.margins[m]:
            wrongMarginsCount += 1
    if wrongMarginsCount > 0:
        # testdocument.setMargins(sampledocument.margins)
        print('Wrong margins')
    else:
        print('Correct')


def checkCharacterStyle(testfont, samplefont):

    if testfont.color != samplefont.color:
        print(f'Wrong text color: {testfont.color} - should be {samplefont.color}')
        # testfont.color = samplefont.color
        testfont.color = '255'

    if testfont.name != samplefont.name:
        print(f'Wrong font name: {testfont.name} - should be {samplefont.name}')
        # testfont.name = samplefont.name
        testfont.color = '255'

    if testfont.size != samplefont.size:
        print(f'Wrong font size: {testfont.size} - should be {samplefont.size}')
        # testfont.size = samplefont.size
        testfont.color = '255'

    if testfont.bold != samplefont.bold:
        print(f'Wrong font weight: {testfont.bold} - should be {samplefont.bold}')
        # testfont.bold = samplefont.bold
        testfont.color = '255'

    if testfont.italic != samplefont.italic:
        print(f'Wrong cursive: {testfont.italic} - should be {samplefont.italic}')
        # testfont.italic = samplefont.italic
        testfont.color = '255'


def checkParagraphStyle(testformat, sampleformat):
    if testformat.alignment != sampleformat.alignment:
        print(f'Wrong text alignment: {testformat.alignment} - should be {sampleformat.alignment}')
        testformat.alignment = sampleformat.alignment

    if testformat.firstlineindent != sampleformat.firstlineindent:
        print(f'Wrong first line indent: {testformat.firstlineindent} - should be {sampleformat.firstlineindent}')
        testformat.firstlineindent = sampleformat.firstlineindent

    if testformat.linespacing != sampleformat.linespacing:
        print(f'Wrong line spacing: {testformat.linespacing} - should be {sampleformat.linespacing}')
        testformat.linespacing = sampleformat.linespacing

    if testformat.spaceafter != sampleformat.spaceafter:
        print(f'Wrong spacing after paragraph: {testformat.spaceafter} - should be {sampleformat.spaceafter}')
        testformat.spaceafter = sampleformat.spaceafter

    if testformat.spacebefore != sampleformat.spacebefore:
        print(f'Wrong spacing before paragraph: {testformat.spacebefore} - should be {sampleformat.spacebefore}')
        testformat.spacebefore = sampleformat.spacebefore


app = initialize()

dir_path = os.path.dirname(os.path.realpath(__file__))
sampleDoc = app.Documents.Open(dir_path + '\\sample.docx')
testDoc = app.Documents.Open(dir_path + '\\test.docx')

testDocument = DocumentStyle(testDoc)
sampleDocument = DocumentStyle(sampleDoc)
checkFormatting(testDocument, sampleDocument)

testDoc.SaveAs(dir_path + '\\test_formatted.docx')
sampleDoc.Close()
