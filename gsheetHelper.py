#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  gsheetHelper.py
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
import pandas as pd
import re


class GSheetHelper(gdf.gdriveFile):

    RESULTS_SHEET = "Results"

    @classmethod
    def assertIsSheet(cls, obj):
        try:
            assert obj.attribs["mimeType"] == gdf.gdriveFile.GDOC_SHEET_MIMETYPE
        except:
            print(obj.attribs["mimeType"])
        assert obj.isSpreadSheet
        obj.__class__ = cls
        obj.cacheFileInfo()

    @classmethod
    def newgdf(cls, access, title="newGdriveFile"):
        """
        Create a new sheet, share it, and access it
        """
        spreadsheetprops = {"properties": {"title": title}}

        spreadsheet = (
            access.sheet_service.spreadsheets()
            .create(
                body=spreadsheetprops,
            )
            .execute()
        )
        fid = spreadsheet["spreadsheetId"]

        user_permission = {
            "type": "user",
            "role": "writer",
            "emailAddress": "pete.siddall@gmail.com",
        }
        access.drive_service.permissions().create(
            fileId=fid, body=user_permission, fields="id"
        ).execute()

        doc = cls({"id": fid, "mimeType": "spreadsheet"})
        doc.cacheAccess(access)
        doc.cacheFileInfo()
        return doc

    def appendSheet(self, newSheetName):
        if newSheetName in self.sheets:
            print(f"appendSheet({newSheetName}) sheet already exists")
            return
        sheetProperties = {"title": newSheetName}

        params = {
            "spreadsheetId": self.gdocId,
            "body": {"requests": [{"addSheet": {"properties": sheetProperties}}]},
        }
        resp = self.sheet_service.spreadsheets().batchUpdate(**params).execute()
        self.sheets.append(newSheetName)

    def publishDF(self, df, startRow=2, resultsSheet=None, growSheet=False):
        if resultsSheet is None:
            resultsSheet = GSheetHelper.RESULTS_SHEET
        # First the 'frame'
        if type(df.index) == pd.pandas.core.indexes.multi.MultiIndex:
            self.addData2d(
                "A",
                startRow + 1,
                [list(i) for i in list(df.index)],
                arrayOf="ROW",
                sheet=resultsSheet,
                growSheet=growSheet,
            )
            firstColOffset = 1
        else:
            self.addData(
                "A",
                startRow + 1,
                list(df.index),
                arrayRepresents="COLUMN",
                sheet=resultsSheet,
            )
            firstColOffset = 0

        self.addData(
            2 + firstColOffset,
            startRow,
            list(df.columns),
            arrayRepresents="ROW",
            sheet=resultsSheet,
            growSheet=growSheet,
        )

        # Then we start data in Column B = 2, but enumerate index is 0 based
        m = df.values.tolist()  # The matrix of values turned into a list of rowlists
        self.addData2d(
            2 + firstColOffset,
            startRow + 1,
            m,
            arrayOf="ROW",
            sheet=resultsSheet,
            growSheet=growSheet,
        )


class GSheetPublisher(object):
    """
    This class defines a publisher for a dataframe.
    There is a mapping from df column names to sheet column titles
    There is a dictionary of formatters to apply to named columns
    """

    CELL_NAME_PATTERN = re.compile("(\w+)(\d+)")

    def __init__(self, df: pd.DataFrame):
        self.dataFrame = df
        self.formatters = {}
        self.columns = {}
        self.doc = None

    def writeLocation(self, doc, sheet, cell):
        """
        links the publisher to a spreadsheet location
        """
        assert doc.isSpreadSheet == True
        self.doc = doc
        if sheet not in doc.sheets:
            print(f"{sheet} is not in {doc.sheets}")
            raise ValueError
        match = GSheetPublisher.CELL_NAME_PATTERN.match(cell)
        if not match:
            print(f"dodgy cell name {cell}")
            raise ValueError
        else:
            self.sheet = sheet
            self.row = int(match[2])
            self.column = match[1]
            self.colNum = gdf.gdriveFile.string_colnum(self.column)

    def writeData(self):
        if self.doc is None:
            print("No location for writer: use writeLocation() first")
        # First the 'frame'
        # Row titles
        if type(self.dataFrame.index) == pd.pandas.core.indexes.multi.MultiIndex:
            self.doc.addData2d(
                self.column,
                self.row + 1,
                [list(i) for i in list(self.dataFrame.index)],
                arrayOf="ROW",
                sheet=self.sheet,
                growSheet=True,
            )
            firstColOffset = 1
        else:
            self.doc.addData(
                self.column,
                self.row + 1,
                list(self.dataFrame.index),
                arrayRepresents="COLUMN",
                sheet=self.sheet,
            )
            firstColOffset = 0
        # Column Headings
        self.doc.addData(
            self.colNum + firstColOffset + 1,
            self.row,
            list(self.dataFrame.columns),
            arrayRepresents="ROW",
            sheet=self.sheet,
        )
        # Now the data
        self.doc.addData2d(
            self.colNum + firstColOffset + 1,
            self.row + 1,
            self.renderData(),
            arrayOf="ROW",
            sheet=self.sheet,
            growSheet=False,
        )  # row titles must have grown the sheet

    def renderData(self):
        data = self.dataFrame.values.tolist()
        outData = []
        for row in data:
            outRow = []
            for i, col in enumerate(row):
                dfColname = self.dataFrame.columns[i]
                fn = self.formatters.get(dfColname)
                if fn is None:
                    res = col
                else:
                    res = fn(col)
                outRow.append(res)

            # print(outRow)
            outData.append(outRow.copy())
        # print(outData)
        return outData

    def addFormatter(self, colname, function):
        if colname in self.dataFrame.columns:
            self.formatters[colname] = function
        else:
            print(f"{colname} is not a column name")
            raise ValueError


def main(args):
    print(f"Can only import gSheetHelper.py")
    return 0


if __name__ == "__main__":
    import sys

    sys.exit(main(sys.argv))
