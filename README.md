# pythonGooogleDoc
local python client layers for handling my Google drive and documents.  With particular emphasis on turning sheets into pandas tables

## Setup

You will need your google id set up for api access - part of this process involves getting a `token.json` file for use in your code.

With `token.json` in the same directory as this code do:

```
python -m gdriveFile
```

which should take you through browser-based authentication.

## Use

``` python
import sys

searchstring = "name contains 'my doc'"

sys.path.append("<path to gdriveFile>")
import gdriveFile as gf

access = gf.gdriveAccess()
gdoc = gf.driveFile.findDriveFile(access, searchstring)

gdoc.showFileInfo()

gdf = gdoc.toDataFrame(usecols = [0,1,2,4])

primarySheet = gdf['Sheet One']
```
