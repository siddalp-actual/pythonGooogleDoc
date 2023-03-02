#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  gdocHelper.py
#
#  Copyright 2020 Pete Siddall <pete.siddall@gmail.com>
#
#  This program is free software; you can redistribute it and/or modify
#  it under the terms of the GNU General Public License as published by
#  the Free Software Foundation; either version 2 of the License, or
#  (at your option) any later version.
#
#  This program is distributed in the hope that it will be useful,
#  but WITHOUT ANY WARRANTY; without even the implied warranty of
#  MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
#  GNU General Public License for more details.
#
#  You should have received a copy of the GNU General Public License
#  along with this program; if not, write to the Free Software
#  Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
#  MA 02110-1301, USA.
#
#

import gdriveFile as gdf


class GdocHelper(gdf.gdriveFile):

    GDOC_DOC_MIMETYPE = "application/vnd.google-apps.document"

    @classmethod
    def assertIsDoc(cls, obj):
        assert obj.isDocument
        obj.__class__ = cls
        obj.cacheFileInfo()
        obj.parseBodyContent()

    def refresh(self):
        '''
        After an update operation, the local copy is stale, need to
        force a re-cache from drive, and update the outline.
        '''
        self.cacheFileInfo(force = True)
        self.parseBodyContent()

    def appendToDoc(self, text):
        appendRequest = {
            "insertText": {
                "text": f"\n{text}",  # this newline pushes text into a new para
                "endOfSegmentLocation": {"segmentId": ""},
            }
        }

        self.docs_service.documents().batchUpdate(
            documentId=self.gdocId, body={"requests": [appendRequest]}
        ).execute()

    def appendTextWithHeader(self, header, text):
        '''
        Inserts the text at the end of the current document
        '''
        self.insertTextWithHeader(header, text, self.docExtent + 1)

    def insertTextWithHeader(self, header, text, startPos):
        insertLen = len(header)
        textInsert1 = GdocHelper.buildInsertText("\n" + header + "\n*", startPos)
        setInsert1Format = GdocHelper.buildStyleUpdate(
            "HEADING_2", startPos + 1, startPos + insertLen
        )
        setAdjunctNormal = GdocHelper.buildStyleUpdate(
            "NORMAL_TEXT", startPos + insertLen + 2, startPos + insertLen + 3
        )
        # next insert jumps header, 2 newlines, and the '*' adjunct
        textInsert2 = GdocHelper.buildInsertText(text, startPos + insertLen + 3)
        deleteAdjunct = GdocHelper.buildDeleteRange(
            startPos + insertLen + 2, startPos + insertLen + 3
        )
        operations = [
            textInsert1,
            setInsert1Format,
            setAdjunctNormal,
            textInsert2,
            deleteAdjunct,
        ]
        resp = (
            self.docs_service.documents()
            .batchUpdate(documentId=self.gdocId, body={"requests": operations})
            .execute()
        )
        print(resp)
        self.refresh()  # reload from drive and rebuild outline

    def deleteText(self, start, end):
        delOp = GdocHelper.buildDeleteRange(start, end)
        resp = (
            self.docs_service.documents()
            .batchUpdate(documentId=self.gdocId, body={"requests": [delOp]})
            .execute()
        )
        print(resp)
        self.refresh()  # reload from drive and rebuild outline


    @staticmethod
    def buildAppendText(stuff):
        rb = {"insertText": {"text": stuff, "endOfSegmentLocation": {"segmentId": ""}}}
        return rb

    @staticmethod
    def buildInsertText(stuff, where):
        rb = {"insertText": {"text": stuff, 'location': {'index': where } }}
        return rb

    @staticmethod
    def buildStyleUpdate(textStyle, startIndex, endIndex):
        assert endIndex > startIndex
        rb = {
            "updateParagraphStyle": {
                "fields": "namedStyleType",
                "paragraphStyle": {"namedStyleType": textStyle},
                "range": {"startIndex": startIndex, "endIndex": endIndex},
            }
        }
        return rb

    @staticmethod
    def buildDeleteRange(startIndex, endIndex):
        assert endIndex > startIndex
        rb = {
            "deleteContentRange": {
                "range": {
                    "startIndex": startIndex,
                    "endIndex": endIndex,
                    "segmentId": "",
                }
            }
        }
        return rb

    def parseBodyContent(self):
        self.outline = DocumentOutline()
        objectList = []
        for section in self.fileInfo["body"]["content"]:
            attribs = section.keys()
            if "paragraph" in attribs:
                p = Paragraph(section)
                if p.isHeading:
                    self.outline.addSection(p)
            elif "sectionBreak" in attribs:
                p = SectionBreak(section)
            elif "table" in attribs:
                print("table still unparsed")
                raise Error
            elif "tableofcontents" in attribs:
                print("toc still unparsed")
                raise Error
            else:
                print(attribs)
                raise Error
            objectList.append(p)
        self.objectList = objectList
        self.docExtent = objectList[-1].endPos

    def __len__(self):
        return self.docExtent


class Section(object):
    def __init__(self, attrDict):
        self.endPos = attrDict["endIndex"]
        self.attrs = attrDict
        self.modified = False

    def __str__(self):
        return f"({self.startPos:3d}, {self.endPos:3d}) : {type(self)}"


class SectionBreak(Section):
    def __init__(self, attrDict):
        super().__init__(attrDict)
        assert "sectionBreak" in attrDict.keys()
        self.style = attrDict["sectionBreak"]["sectionStyle"]
        self.startPos = self.endPos


class Paragraph(Section):
    def __init__(self, attrDict):
        super().__init__(attrDict)
        assert "paragraph" in attrDict.keys()
        self.startPos = attrDict["startIndex"]
        self.style = attrDict["paragraph"]["paragraphStyle"]
        self.isHeading = self.style["namedStyleType"][:7] == "HEADING"
        self.elements = []
        for el in self.attrs["paragraph"]["elements"]:
            self.elements.append(TextElement(el))
        if self.isHeading:
            self.heading = self.elements[0].content

    def __str__(self):
        # s = f"({self.startPos:3d}, {self.endPos:3d}) : {type(self)}"
        s = Section.__str__(self)
        if self.isHeading:
            s += f"<{self.heading}>"
        for el in self.elements:
            s += "\n" + "   " + el.__str__()
        return s


class TextElement(Section):
    def __init__(self, attrDict):
        super().__init__(attrDict)
        assert "textRun" in attrDict.keys()
        self.startPos = attrDict["startIndex"]
        self.style = attrDict["textRun"]["textStyle"]
        self.content = attrDict["textRun"]["content"]

    def __str__(self):
        return Section.__str__(self) + ": >" + self.content


class DocumentOutline(object):

    import re

    # import collections

    dateMatch = re.compile(r"(\w{3}\s\d+\s\w{3}\s\d{4})")

    def __init__(self):
        self.headings = []  # an ordered list of Paragraphs
        # self.headingsIndex = collections.OrderedDict()
        # NOTE Python > 3.7 guarantees insertion order of keys
        self.headingsIndex = dict()  # a hash of the titles, pointing at it's index

    def addSection(self, s):
        assert s.isHeading
        self.headings.append(s)
        i = self.headings.index(s)
        # i = s.startPos
        self.headingsIndex[s.heading] = i
        return i

    def findFirstDate(self, after=0):
        for h, i in self.headingsIndex.items():
            if i <= after:
                continue
            matchObj = DocumentOutline.dateMatch.search(h)
            if matchObj:
                return (matchObj[0], i)
        print("date not found in headingsIndex")
        raise Exception('no date found in outline headings')


def main(args):
    print("use import gdocHelper ONLY")
    return 0


if __name__ == "__main__":
    import sys

    sys.exit(main(sys.argv))
