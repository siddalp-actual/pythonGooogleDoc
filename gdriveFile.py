#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
#  gdriveFile.py
#
#  Copyright 2019 Pete Siddall <pete.siddall@gmail.com>
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
import oauth2client.file
import oauth2client.client
import oauth2client.tools

# google-api-python-client provides the next two
import apiclient.discovery
import apiclient.http
import pprint
import pandas as pd
import os.path


class gdriveFile:
    GDOC_SHEET_MIMETYPE = "application/vnd.google-apps.spreadsheet"
    GDOC_DOC_MIMETYPE = "application/vnd.google-apps.document"

    @staticmethod
    def findDriveFile(access, query):
        """
        search google drive for a file with name matching query
        if a single one is found, instantiate it as a gdriveFile object
        """

        def show_file_info(a):
            print(
                "Found file: %s (%s) last change %s\ntype %s"
                % (a["name"], a["id"], a["modifiedTime"], a["mimeType"])
            )

        # pass in a document query, and return the (hopefully) only
        # corresponding file id
        print(type(access))
        drive = access.drive_service
        page_token = None
        fileList = []
        while True:
            # response looks like a dict with:
            # files: an array containing the requested attributes for each one
            # nextPageToken: there's more to come indicator
            response = (
                drive.files()
                .list(
                    q=query,
                    spaces="drive",
                    fields="nextPageToken, files(id, name, modifiedTime, mimeType)",
                    pageToken=page_token,
                )
                .execute()
            )

            fileList.extend(response["files"])

            page_token = response.get("nextPageToken", None)
            if page_token is None:
                break

        if len(fileList) == 1:
            show_file_info(fileList[0])
            newObj = gdriveFile(fileList[0])
            newObj.cacheAccess(access)
            return newObj
        else:
            print("Multiple files returned by search")
            for f in fileList:
                show_file_info(f)

            return 0

    @staticmethod
    def colnum_string(n: int) -> str:
        """
        Convert a column number (starting at 1) the character column string
        """
        string = ""
        assert n > 0
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    @staticmethod
    def string_colnum(s: int) -> int:
        """
        Convert a character column string to the column number
        """
        assert len(s) <= 2
        n = 0
        for c in s:
            n *= 26
            n += ord(c) - 64
        return n

    @staticmethod
    def createValueRange(
        colname: str, startrow: int, data, arrayRepresents, sheet
    ):
        """
        Marshalls the start column, row and data into the form expected by googlesheet API
        enforces only a single ROW|COLUMN is updated
        """
        if type(data) != list:
            lenData = 1
            data = [data]
        else:
            if type(data[0]) == list:
                # need to flatten it
                if len(data) > 1:
                    print("too many entries in data list")
                    raise ValueError
                data = [i for i in data[0]]

            lenData = len(data)
        print(lenData, data)

        valueRange = {}
        if arrayRepresents == "COLUMN":
            valueRange.update(
                {
                    "range": "{}!{}{}:{}{}".format(
                        sheet,
                        colname,
                        startrow,
                        colname,
                        startrow + lenData - 1,
                    )
                }
            )
        else:
            colnum = gdriveFile.string_colnum(colname)
            valueRange.update(
                {
                    "range": "{}!{}{}:{}{}".format(
                        sheet,
                        colname,
                        startrow,
                        gdriveFile.colnum_string(colnum + lenData - 1),
                        startrow,
                    )
                }
            )
        valueRange.update({"majorDimension": arrayRepresents + "S"})
        valueRange.update({"values": [data]})

        return valueRange

    @staticmethod
    def createValueRange2d(
        colname: str, startrow: int, data, arrayOf, sheet
    ):
        """
        Marshalls the start column, row and data into the form expected by googlesheet API
        2d data can be provided

        arrayOf = ROW data = [[A1, B1], [A2, B2]]
        arrayOf = COLUMN data = [[A1, A2, A3]]
        """
        if type(data) != list:
            lenData = 1
            widthData = 1
            data = [[data]]
        else:
            lenData = len(data)
            if type(data[0]) == list:
                # maybe need to flatten it
                widthData = len(data[0])
                if widthData > 1:
                    # assume 2d data
                    for n, i in enumerate(data):
                        if len(i) != widthData:
                            print(
                                f"{n}th element >{i}< is {len(i)} rest are {widthData}"
                            )
                            raise ValueError

                    data = [i for i in data]
            else:
                data = [[i for i in data]]
                widthData = 1

        print(f"({lenData}, {widthData})", data)

        valueRange = {}
        colnum = gdriveFile.string_colnum(colname)
        if arrayOf == "ROW":
            valueRange.update(
                {
                    "range": "{}!{}{}:{}{}".format(
                        sheet,
                        colname,
                        startrow,
                        gdriveFile.colnum_string(
                            colnum + widthData - 1
                        ),
                        startrow + lenData - 1,
                    )
                }
            )
        else:
            valueRange.update(
                {
                    "range": "{}!{}{}:{}{}".format(
                        sheet,
                        colname,
                        startrow,
                        gdriveFile.colnum_string(colnum + lenData - 1),
                        startrow + widthData - 1,
                    )
                }
            )
        valueRange.update({"majorDimension": arrayOf + "S"})
        valueRange.update({"values": data})
        print(valueRange)
        return valueRange

    def __init__(self, fileDict):
        """
        create a gdriveFile object based on attributes
        at least needs an 'id' for further operations
        """
        DEF_MIME_TYPE = "application/octet-stream"
        self.attribs = fileDict
        self.gdocId = self.attribs["id"]
        self.isSpreadSheet = (
            self.attribs.get("mimeType", DEF_MIME_TYPE)[-11:]
            == "spreadsheet"
        )
        self.isDocument = (
            self.attribs.get("mimeType", DEF_MIME_TYPE)[-8:]
            == "document"
        )
        self.fileInfo = None
        self.fileData = None
        self.sheetDict = {}
        self.sheetLen = {}

    @classmethod
    def gdfFromId(cls, fid, access, docType="spreadsheet"):
        """
        classmethod allows alternate constructor
        """
        assert (
            docType[-11:] == "spreadsheet" or docType[-8:] == "document"
        )
        if fid[:4] == "http":
            print("found url:", fid)
            url = fid
            p = url.find("?id=")  # pos of substring
            print("contains id at:", p)
            fid = url[p + 4 :]

        if docType == "spreadsheet":
            docType = gdriveFile.GDOC_SHEET_MIMETYPE
        if docType == "document":
            docType = gdriveFile.GDOC_DOC_MIMETYPE

        doc = cls({"id": fid, "mimeType": docType})
        doc.cacheAccess(access)
        doc.cacheFileInfo()
        return doc

    @classmethod
    def newgdf(cls, access, title="newGdriveFile"):
        """
        Create a new file, share it, and access it
        """
        fileprops = {"name": title, "originalFilename": "noname.yet"}

        spreadsheet = (
            access.drive_service.files()
            .create(
                body=fileprops,
            )
            .execute()
        )
        fid = spreadsheet["id"]

        user_permission = {
            "type": "user",
            "role": "writer",
            "emailAddress": "pete.siddall@gmail.com",
        }
        access.drive_service.permissions().create(
            fileId=fid, body=user_permission, fields="id"
        ).execute()

        doc = cls(
            {
                "id": fid,
            }
        )
        doc.cacheAccess(access)
        doc.cacheFileInfo()
        return doc

    def cacheAccess(self, access):
        """
        cache the api access objects with the file
        """
        self.access = access
        if type(access) != gdriveAccess:
            print("type {} is not a gdriveAcess object")
            self.sheet_service = None
            raise (typeError)
        else:
            self.sheet_service = self.access.sheet_service
            self.docs_service = self.access.docs_service

    def cacheFileInfo(self, force=False):
        """
        pull down all sorts of useful info about the file
        in particular, we're after the title, number of sheets and their names

        force: if true, then re-cache the file
        """
        # assert(self.isSpreadSheet is True)
        if self.fileInfo and not force:
            return
        else:
            if self.isSpreadSheet:
                self.fileInfo = (
                    self.sheet_service.spreadsheets()
                    .get(spreadsheetId=self.gdocId)
                    .execute()
                )
                self.title = self.fileInfo["properties"]["title"]
                self.sheets = [
                    s["properties"]["title"]
                    for s in self.fileInfo["sheets"]
                ]
                # set columnCount, rowCount
                self.sheetMaxSize = [
                    s["properties"]["gridProperties"]
                    for s in self.fileInfo["sheets"]
                ]
                self.defaultSheet = self.sheets[0]
            elif self.isDocument:
                self.fileInfo = (
                    self.docs_service.documents()
                    .get(documentId=self.gdocId)
                    .execute()
                )
                self.title = self.fileInfo["title"]
            else:
                self.fileInfo = (
                    self.access.drive_service.files()
                    .get(fileId=self.gdocId, fields="*")
                    .execute()
                )

            self.versionInfo = self.getVersions()

    def getVersions(self):
        """
        ask for information about the document versions
        """
        resp = (
            self.access.drive_service.revisions()
            .list(
                fileId=self.gdocId,
                fields="*",
                pageSize=1000,
            )
            .execute()
        )
        return [
            {
                f: n[f]
                for f in ["id", "modifiedTime", "lastModifyingUser"]
            }
            for n in resp["revisions"]
        ]

    def uploadNewFile(self, filename, mimetype="text/csv"):
        """
        replace the existing version of the text with the named file
        """
        assert os.path.isfile(filename)
        fileData = apiclient.http.MediaFileUpload(
            filename, mimetype=mimetype
        )
        current_version = self.versionInfo[-1]["id"]
        try:
            new_version = current_version + 1
        except (ValueError, TypeError):
            new_version = 1

        fileMetaData = {
            "properties": {"revisionId": new_version, "name": filename}
        }
        update = (
            self.access.drive_service.files()
            .update(
                fileId=self.gdocId,
                body=fileMetaData,
                media_body=fileData,
            )
            .execute()
        )
        self.cacheFileInfo(force=True)

    def cacheFileData(self):
        """
        pull down cell values into one valuerange per sheet
        valueRenderOption determines whether values are formatted,
        unformatted, or formulae returned
        """
        assert self.isSpreadSheet is True
        self.cacheFileInfo()
        if self.fileData:
            return
        else:
            params = {
                "spreadsheetId": self.gdocId,
                "ranges": self.sheets,
                "majorDimension": "ROWS",
            }
            self.fileData = (
                self.sheet_service.spreadsheets()
                .values()
                .batchGet(**params)
                .execute()
            )
            self.setSheetExtents()

    def setSheetExtents(self):
        rangeList = self.fileData["valueRanges"]
        lastCol = []
        lastRow = []

        # print(rangeList)
        assert len(rangeList) == len(self.sheets)  # or problem cos
        # everything I've seen so far, says 1 valueRange per sheet
        for entry in rangeList:
            assert entry["majorDimension"] == "ROWS"
            values = entry.get("values", [])
            maxlen = 0
            for row in values:
                rowlen = len(row)
                maxlen = max(rowlen, maxlen)
            lastCol.append(maxlen)
            lastRow.append(len(values))

        self.lastCol = dict(zip(self.sheets, lastCol))
        self.lastRow = dict(zip(self.sheets, lastRow))

    def addData(
        self,
        startCol,
        startRow,
        dataArray,
        arrayRepresents="ROW",
        sheet=None,
        growSheet=False,
    ):
        """
        Add a list of cells to a spreadsheet - assumed to be first sheet
        startCol: is the top left coordinate, either a 1 based number, or letter
        startRow: is the top left row number
        dataArray: a list containing a single list of values to enter
        arrayRepresents: defaults to ROW -> the values are for consecutive columns
                                    COLUMN -> the values are for consecutive rows
        growSheet: False throw an error if data beyond end, True then append 5 rows
        """
        if type(startCol) == int:
            startCol = gdriveFile.colnum_string(startCol)

        if arrayRepresents != "ROW" and arrayRepresents != "COLUMN":
            print("arrayRepresents parameter must be ROW|COLUMN")
            raise ValueError

        sheetIndex = None
        if sheet is None:
            sheet = self.defaultSheet
            sheetIndex = 0
        else:
            for n, s in enumerate(self.sheets):
                if sheet == s:
                    sheetIndex = n
                    break
        if sheetIndex is None:
            print("invalid sheet specified")
            raise ValueError

        if startRow > (self.sheetMaxSize[sheetIndex]["rowCount"] - 5):
            if growSheet:
                self.append5Rows(sheet, sheetIndex)
            else:
                raise ValueError

        data = gdriveFile.createValueRange(
            startCol,
            startRow,
            dataArray,
            arrayRepresents=arrayRepresents,
            sheet=sheet,
        )

        params = {
            "spreadsheetId": self.gdocId,
            "body": {"data": data, "valueInputOption": "user_entered"},
        }
        resp = (
            self.sheet_service.spreadsheets()
            .values()
            .batchUpdate(**params)
            .execute()
        )

    def addData2d(
        self,
        startCol,
        startRow,
        dataArray,
        arrayOf="ROW",
        sheet=None,
        growSheet=False,
    ):
        """
        Add a list of cells to a spreadsheet - assumed to be first sheet
        startCol: is the top left coordinate, either a 1 based number, or letter
        startRow: is the top left row number
        dataArray: a list containing a single list of values to enter
        arrayRepresents: defaults to ROW -> the values are for consecutive columns
                                    COLUMN -> the values are for consecutive rows
        growSheet: False throw an error if data beyond end, True then append 5 rows
        """
        if type(startCol) == int:
            startCol = gdriveFile.colnum_string(startCol)

        if arrayOf != "ROW" and arrayOf != "COLUMN":
            print("arrayRepresents parameter must be ROW|COLUMN")
            raise ValueError

        sheetIndex = None
        if sheet is None:
            sheet = self.defaultSheet
            sheetIndex = 0
        else:
            for n, s in enumerate(self.sheets):
                if sheet == s:
                    sheetIndex = n
                    break
        if sheetIndex is None:
            print("invalid sheet specified")
            raise ValueError

        if startRow > (self.sheetMaxSize[sheetIndex]["rowCount"] - 5):
            if growSheet:
                self.append5Rows(sheet, sheetIndex)
            else:
                raise ValueError

        data = gdriveFile.createValueRange2d(
            startCol, startRow, dataArray, arrayOf=arrayOf, sheet=sheet
        )

        params = {
            "spreadsheetId": self.gdocId,
            "body": {"data": data, "valueInputOption": "user_entered"},
        }
        resp = (
            self.sheet_service.spreadsheets()
            .values()
            .batchUpdate(**params)
            .execute()
        )

    def append5Rows(self, sheet, sheetIndex):
        # append at the end of the sheet
        aP = self.sheetMaxSize[sheetIndex][
            "rowCount"
        ]  # append position
        appendData = gdriveFile.createValueRange(
            "A",
            aP,
            ["" for i in range(5)],
            arrayRepresents="COLUMN",
            sheet=sheet,
        )

        appendParm = {
            "spreadsheetId": self.gdocId,
            "range": "{}!{}{}:{}{}".format(sheet, "A", aP, "A", aP + 4),
            "body": appendData,
            "valueInputOption": "USER_ENTERED",  # ['INPUT_VALUE_OPTION_UNSPECIFIED', 'RAW', 'USER_ENTERED']
            "insertDataOption": "INSERT_ROWS",  # overwrite
        }
        print(appendParm)

        resp = (
            self.sheet_service.spreadsheets()
            .values()
            .append(**appendParm)
            .execute()
        )

    def toDataFrame(self, usecols=None):
        """
        iterate over the sheets, building a dataframe for each
        these are stashed into a dictionary
        """
        self.cacheFileData()

        for n in range(len(self.fileData["valueRanges"])):
            df = self.sheetToDataFrame(n, usecols=usecols)
            self.sheetDict.update({df.name: df})
        return self.sheetDict

    def sheetToDataFrame(self, i, usecols=None):
        """
        return a pandas dataframe containing the data, and
        the number of columns it contains
        """

        def addrow(row, usecols):
            tempDict = {}
            n = 0
            for y, item in enumerate(row):
                if usecols is None or y in usecols:
                    # try converting numbers instead of leaving as string
                    if len(item) > 0 and item[0].isdigit():
                        try:
                            item = float(item)
                        except:
                            pass
                    tempDict.update({y: item})
                    n += 1
            for y in range(n, len(usecols)):
                tempDict.update(
                    {usecols[y]: ""}
                )  # add dummy value for wanted cols
            return (tempDict, n)

        stuff = []
        maxcols = 0

        for n, row in enumerate(
            self.fileData["valueRanges"][i].get("values", [])
        ):
            # print("row({}): {}".format(n,row))
            (newRow, cols) = addrow(row, usecols)
            stuff.append(newRow)
            maxcols = max(maxcols, cols)

        # look at first cell in first row to guess whether the first row
        # contains labels
        try:
            if stuff[0][0] == "" or stuff[0][0][0].isdigit():
                # probably not
                df = pd.DataFrame(stuff)
            else:
                # make sure we have enough labels for all the columns used
                c = list(stuff[0].values())  # based on first row labels
                e = []
                for n in range(maxcols):
                    if n >= len(c) or not c[n] or c[n] == "":
                        e.append(n)
                    else:
                        e.append(c[n])
                # print(e)
                df = pd.DataFrame(stuff[1:])
                df.columns = e
        except IndexError:
            print(
                f"IndexError. Stuff: {stuff} assigning null df for sheet {self.sheets[i]}"
            )
            df = pd.DataFrame()

        df.name = self.sheets[i]
        self.sheetLen.update({df.name: len(df)})
        return df

    def showFileInfo(self):
        self.cacheFileInfo()
        pp = pprint.PrettyPrinter()
        pp.pprint(self.attribs)
        print("spreadsheet: {}".format(self.isSpreadSheet))
        pp.pprint(self.fileInfo)
        print("Title:   {}\nSheets: {}".format(self.title, self.sheets))

    def itersheets(self):
        for name, vr in zip(self.sheets, self.fileData["valueRanges"]):
            yield (self.title, name), vr["values"]

    #  Untested stuff for CSV
    @staticmethod
    def write_csv(fd, rows, dialect="excel"):
        csvfile = csv.writer(fd, dialect=dialect)
        csvfile.writerows(rows)

    def export_csv(
        self,
        service,
        docid,
        filename_template="%(title)s : %(sheet)s.csv",
    ):
        if not self.isSpreadSheet:
            print("object is not a spreadsheet")
            raise (typeError)
        for (doc, sheet), rows in self.itersheets(service, docid):
            filename = filename_template % {
                "title": doc,
                "sheet": sheet,
            }
            with open(filename, "w", newline="") as fd:
                self.write_csv(fd, rows)


class gdriveAccess:
    import os

    SCOPE = [
        "https://www.googleapis.com/auth/drive.metadata.readonly",
        #'https://www.googleapis.com/auth/drive.readonly',
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/memento",
        "https://www.googleapis.com/auth/reminders",
        #'https://www.googleapis.com/auth/spreadsheets.readonly',
        # added next to try to allow ROBOT to update my sheet
        "https://www.googleapis.com/auth/spreadsheets",
        #'https://www.googleapis.com/auth/documents.readonly'
        "https://www.googleapis.com/auth/documents",
    ]
    # TOKENFILE = "/home/siddalp/Dropbox/pgm/googledocs/token.json"
    # CREDFILE = "/home/siddalp/Dropbox/pgm/googledocs/client_id.json"

    # __file__ is the absolute path to this module when it has been imported
    TOKENFILE = os.path.join(os.path.dirname(__file__), "token.json")
    CREDFILE = os.path.join(os.path.dirname(__file__), "client_id.json")

    def __init__(self):
        # we have a cached oauth token
        credStore = oauth2client.file.Storage(gdriveAccess.TOKENFILE)
        creds = credStore.get()
        # but if it doesn't exist, we need to request a new one, passing the
        # application's credentials and requested priviledges
        # This potentially opens a browser dialog to grant permission
        if not creds or creds.invalid:
            flow = oauth2client.client.flow_from_clientsecrets(
                gdriveAccess.CREDFILE, gdriveAccess.SCOPE
            )
            creds = oauth2client.tools.run_flow(flow, credStore)

        self.credentials = creds
        # create an application end point for interaction with google drive
        # cache_discovery=False added 25/2/22 to remove logging warning about
        #  file_cache only supported with client < 4.0.0
        self.drive_service = apiclient.discovery.build(
            "drive", "v3", credentials=creds, cache_discovery=False
        )
        # and another for the sheets API used for pulling out different tabs
        # and data
        self.sheet_service = apiclient.discovery.build(
            "sheets",
            version="v4",
            credentials=creds,
            cache_discovery=False,
        )

        self.docs_service = apiclient.discovery.build(
            "docs", "v1", credentials=creds, cache_discovery=False
        )

    def __enter__(self):
        """
        enable resource manager function:
        with gdriveAccess() as access:
        """
        return self

    def __exit__(self, *args):
        """
        enable resource manager function
        """
        pass

    def get_drive_service(self):
        return self.drive_service

    def get_sheet_service(self):
        return self.sheet_service


def maybeInvokeAuthDialog():
    # When scope updated, delete the token.json token file and run
    #
    # python -m gdriveFile
    #
    # we have a cached oauth token
    credStore = oauth2client.file.Storage(gdriveAccess.TOKENFILE)
    creds = credStore.get()
    # but if it doesn't exist, we need to request a new one, passing the
    # application's credentials and requested priviledges
    # This potentially opens a browser dialog to grant permission
    if not creds or creds.invalid:
        flow = oauth2client.client.flow_from_clientsecrets(
            gdriveAccess.CREDFILE, gdriveAccess.SCOPE
        )
        creds = oauth2client.tools.run_flow(flow, credStore)


def main(args):
    print("only use as a module")
    maybeInvokeAuthDialog()
    return 0


if __name__ == "__main__":
    import sys

    sys.exit(main(sys.argv))
