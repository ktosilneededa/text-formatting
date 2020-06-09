from win32com import client
import os

wordApp = client.Dispatch('Word.Application')
wordApp.Visible = 1
wordApp.DisplayAlerts = 0

dir_path = os.path.dirname(os.path.realpath(__file__))

sampleDoc = wordApp.Documents.Open(dir_path + '\\sample.docx')
testDoc = wordApp.Documents.Open(dir_path + '\\test.docx')

class DocumentStyle:
    def __init__(self, doc):
        self.document = doc
        self.paragraphs = doc.Paragraphs
        self.paragraphStyles = self.getStyle()
        self.margins = self.getMargins()

    def getStyle(self):
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
        self.name = paragraph.style
        self.font = paragraph.style.font
        self.format = paragraph.format


def checkFormatting(testdocument, sampledocument):
    print('Text formatting:')
    checkTextStyle(testdocument, sampledocument)
    print('\nPage formatting:')
    checkMargins(testdocument, sampledocument)


def checkTextStyle(testdocument, sampledocument):
    for p in range(len(testdocument.paragraphs)):
        print(f'\nParagraph {p + 1}:')
        checkFormatProperties(testdocument.paragraphStyles[p].format, sampledocument.paragraphStyles[0].format)
        checkFontProperties(testdocument.paragraphStyles[p].font, sampledocument.paragraphStyles[0].font)


def checkMargins(testdocument, sampledocument):
    wrongMarginsCount = 0
    for m in range(4):
        if testdocument.margins[m] != sampledocument.margins[m]:
            wrongMarginsCount += 1
    if wrongMarginsCount > 0:
        testdocument.setMargins(sampledocument.margins)
        print('Wrong margins')


def checkFontProperties(testfont, samplefont):
    if testfont.name != samplefont.name:
        print(f'Wrong font name: {testfont.name} - should be {samplefont.name}')
        testfont.name = samplefont.name

    if testfont.size != samplefont.size:
        print(f'Wrong font size: {testfont.size} - should be {samplefont.size}')
        testfont.size = samplefont.size

    if testfont.bold != samplefont.bold:
        print(f'Wrong font weight: {testfont.bold} - should be {samplefont.bold}')
        testfont.bold = samplefont.bold

    if testfont.italic != samplefont.italic:
        print(f'Wrong cursive: {testfont.italic} - should be {samplefont.italic}')
        testfont.italic = samplefont.italic

    if testfont.color != samplefont.color:
        print(f'Wrong text color: {testfont.color} - should be {samplefont.color}')
        testfont.color = samplefont.color


def checkFormatProperties(testformat, sampleformat):
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


testDocument = DocumentStyle(testDoc)
sampleDocument = DocumentStyle(sampleDoc)
checkFormatting(testDocument, sampleDocument)
sampleDoc.Close()
