import os
import json
from docx import Document
from dictdiffer import diff

filePath = 'sample.docx'
formDataJson = 'formData.json'

sectionOrientationDict = {
    0: "portrait",
    1: "landscape"
}

paragraphAlignmentDict = {
    0: "center",
    1: "left",
    2: "right",
    3: "justify",
    None: "left"
}

firstLineIndentDict = {}
lineSpacingDict = {}


def toCm(value):
    return round(value.cm, 2)


def toPt(value):
    return value.pt


def getBaseStyleProperty(paragraph, property):
    value = getattr(paragraph.paragraph_format, property)
    if value is None:
        value = getattr(paragraph.style.paragraph_format, property)
        if value is None and paragraph.style.base_style is not None:
            value = getattr(paragraph.style.base_style.paragraph_format, property)
            if value is None and paragraph.style.base_style.base_style is None:
                value = None
    return value


def zeroIfNone(value):
    return value if not None else 0


def checkFormatting(formdatapath, sampledatapath):
    formData = open(formdatapath)
    sampleData = open(sampledatapath)
    f, s = json.load(formData), json.load(sampleData)
    d = diff(s['sections'], f['sections'])
    # os.remove(sampleDataPath)
    for i in d:
        print(i)


class Data:
    def __init__(self, file):
        self.document = file
        self.sections = self.document.sections
        self.paragraphs = self.document.paragraphs
        self.text = 0
        self.data = None

    def getSectionProperties(self):
        sectionProperties = [
            {
                "orientation": sectionOrientationDict[self.sections[s].orientation],
                "leftMargin": toCm(self.sections[s].left_margin),
                "rightMargin": toCm(self.sections[s].right_margin),
                "topMargin": toCm(self.sections[s].top_margin),
                "bottomMargin": toCm(self.sections[s].bottom_margin)
            }
            for s in range(len(self.sections))
        ]

        return sectionProperties

    def getParagraphProperties(self):
        paragraphProperties = [
            {
                self.paragraphs[p].style.name: {
                    "paragraphFormat": {
                        "alignment": paragraphAlignmentDict[getBaseStyleProperty(self.paragraphs[p], "alignment")],
                        "firstLineIndent": zeroIfNone(getBaseStyleProperty(self.paragraphs[p], "first_line_indent")),
                        "lineSpacing": zeroIfNone(getBaseStyleProperty(self.paragraphs[p], "line_spacing")),
                        "spaceAfter": toPt(zeroIfNone(getBaseStyleProperty(self.paragraphs[p], "space_after"))),
                        # "spaceBefore": toPt(zeroIfNone(getBaseStyleProperty(self.paragraphs[p], "space_before"))),
                    }
                }
            }
            for p in range(len(self.paragraphs))
        ]

        return paragraphProperties

    def collectData(self):
        self.data = {
            "sections": self.getSectionProperties(),
            # "styles": self.getParagraphProperties()
        }

    def makeJsonFile(self):
        self.collectData()
        file = open('sampleData.json', 'w')
        json.dump(self.data, file, indent=4)
        file.close()
        return file.name


def app():
    document = Document(filePath)
    sampleDataJson = Data(document).makeJsonFile()
    checkFormatting(formDataJson, sampleDataJson)


if __name__ == '__main__':
    app()
