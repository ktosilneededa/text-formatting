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
    pars = len(sampledocument.paragraphs)

    for par_i in range(pars):
        chars_in_pars = len(testdocument.paragraphStyles[par_i].characterStyles)
        char_diffs = [0] * (chars_in_pars + 1)
        red_messages = []
        red_phrases = []
        count = 0

        print(f'\nParagraph {par_i + 1}: ')
        checkParagraphStyle(testdocument.paragraphStyles[par_i].format, sampledocument.paragraphStyles[par_i].format)
        for char_i in range(chars_in_pars):
            character, diff = checkCharacterStyle(testdocument.paragraphStyles[par_i].characterStyles[char_i],
                            sampledocument.paragraphStyles[par_i].characterStyles[char_i])
            char_diffs[char_i] = diff
            if diff != '' and diff != char_diffs[char_i - 1]:
                count += 1
                red_phrases.append('\n')
                red_phrases.append(str(character))
                red_messages.append(diff)
            if diff == char_diffs[char_i - 1] and diff != '':
                red_phrases.append(str(character))

        if count == 0:
            print('Character style is correct')
        else:
            red_phrases = ''.join(red_phrases)
            red_phrases = red_phrases.splitlines()

            for m in range(len(red_messages)):
                print(f'Red area \'{red_phrases[m + 1]}\': {red_messages[m]}')


def checkMargins(testdocument, sampledocument):
    wrongMarginsCount = 0
    for m in range(4):
        if testdocument.margins[m] != sampledocument.margins[m]:
            wrongMarginsCount += 1
    if wrongMarginsCount > 0:
        print('Wrong margins')
    else:
        print('Margins are correct')


def checkCharacterStyle(testchar, samplechar):
    color_wrong = '255'
    message = ''
    character = ''

    if testchar.font.color != samplechar.font.color:
        message += f'Wrong text color: {testchar.font.color} - should be {samplechar.font.color} '
        testchar.font.color = color_wrong

    if testchar.font.name != samplechar.font.name:
        message += f'Wrong font name: {testchar.font.name} - should be {samplechar.font.name} '
        testchar.font.color = color_wrong

    if testchar.font.size != samplechar.font.size:
        message += f'Wrong font size: {testchar.font.size} - should be {samplechar.font.size} '
        testchar.font.color = color_wrong

    if testchar.font.bold != samplechar.font.bold:
        message += f'Wrong font weight: {testchar.font.bold} - should be {samplechar.font.bold} '
        testchar.font.color = color_wrong

    if testchar.font.italic != samplechar.font.italic:
        message += f'Wrong cursive: {testchar.font.italic} - should be {samplechar.font.italic} '
        testchar.font.color = color_wrong

    if message != '':
        character = samplechar.character

    return character, message


def checkParagraphStyle(testformat, sampleformat):
    count = 0

    if testformat.alignment != sampleformat.alignment:
        print(f'Wrong text alignment: {testformat.alignment} - should be {sampleformat.alignment}')
        count += 1

    if testformat.firstlineindent != sampleformat.firstlineindent:
        print(f'Wrong first line indent: {testformat.firstlineindent} - should be {sampleformat.firstlineindent}')
        count += 1

    if testformat.linespacing != sampleformat.linespacing:
        print(f'Wrong line spacing: {testformat.linespacing} - should be {sampleformat.linespacing}')
        count += 1

    if testformat.spaceafter != sampleformat.spaceafter:
        print(f'Wrong spacing after paragraph: {testformat.spaceafter} - should be {sampleformat.spaceafter}')
        count += 1

    if testformat.spacebefore != sampleformat.spacebefore:
        print(f'Wrong spacing before paragraph: {testformat.spacebefore} - should be {sampleformat.spacebefore}')
        count += 1

    if count == 0:
        print('Paragraph style is correct')


app = initialize()

dir_path = os.path.dirname(os.path.realpath(__file__))
sampleDoc = app.Documents.Open(dir_path + '\\sample.docx')
testDoc = app.Documents.Open(dir_path + '\\test.docx')

testDocument = DocumentStyle(testDoc)
sampleDocument = DocumentStyle(sampleDoc)
checkFormatting(testDocument, sampleDocument)

testDoc.SaveAs(dir_path + '\\test_formatted.docx')
sampleDoc.Close()
