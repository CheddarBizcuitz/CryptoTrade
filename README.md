# CryptoTrade

Welcome to CryptoTrade! For the latest release please see:
[![Latest Release](https://img.shields.io/github/release/CheddarBizcuitz/CryptoTrade.svg?label=latest%20release)](https://github.com/CheddarBizcuitz/CryptoTrade/releases/latest)

Questions, feedback and suggestions?
[![Join the chat at https://gitter.im/CryptoTradeTalk](https://badges.gitter.im/CryptoTradeTalk/Lobby.svg)](https://gitter.im/CryptoTradeTalk?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

## Table of Contents

- [What is CryptoTrade?](#what-is-cryptotrade)
- [Live-Feed Data](#live-feed-data)
- [Transferring Data](#transferring-data)
- [Tips](#tips)
- [Downloads](#downloads)
- [License](#license)

## What is CryptoTrade?

CryptoTrade is an interactive spreadsheet that tracks your cryptocurrency trades through the use of VBA and other excel functions. 

CryptoTrade is free to use and can run on Excel Versions 2010 or later.

## Live-Feed Data

CryptoTrade currently uses [Nomics API](https://nomics.com/) to generate free, live-feed data on thousands of cryptocurrencies.

Reference to query data can be found at: https://nomics.com/docs/#tag/Currencies

You can request a free API key at: https://p.nomics.com/cryptocurrency-bitcoin-api

If required, you will need to enter your key into the Excel data query as: https://api.nomics.com/v1/currencies/ticker?key=enterkeyhere

## Transferring Data

Transferred data can be separated into three categories:

1. Source code data (VBA)
2. Formatting data (workbook)
3. Personal data

### Source Code Data

Source code can be updated by running [mImportVB.bas](https://github.com/CheddarBizcuitz/CryptoTrade/blob/main/VBA/mImportVB.bas). This module searches for the folder <mark>"VisualBasic"</mark> on the user's C:\ Drive. You may alter the final path by adjusting the lines below:

**Note:** All Excel objects will be imported as Class Modules (.cls). You will need to manually copy & paste the code into the workbook's corresponding object, then delete the class module(s). Updates to original objects should not be common for future releases, and will most likely be accomplished through updating [Formatting Data](#formatting-data).

#### Root
```VBA
SpecialPath = WshShell.SpecialFolders("C:\")
```

#### Directory
```VBA
    If fso.FolderExists(SpecialPath & "VisualBasic") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VisualBasic"
        On Error GoTo 0
    End If
    
    If fso.FolderExists(SpecialPath & "VisualBasic") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VisualBasic"
    Else
```

### Formatting Data

Any updates to the Excel workbook will need to be accomplished through the use of the latest version. If there are few formatting changes within the release, directions on how to update the user's current version will be listed. Although this is not the best solution, it will hopefully reduce the overwrite process and prevent the need to transfer [Personal Data](#personal-data).

### Personal Data

If the transfer of personal data is required, you will need to manually copy & paste the data into the new workbook. This can be achieved through the following steps:

DWB = Destination Workbook (new release)  
SWB = Source Workbook (current user release)  

1. In the DWB, add coins and 'Generate'
2. Determine line count in SWB
3. Copy & paste SWB data to DWB

To determine the proper line count per worksheet, use the formula ((n-4)+1)-m), where 

n = last row # in SWB  
m = # of rows already present in DWB (preset as 3)  

## Tips

1. All coin sheets are populated from the 'Temp' worksheet. You can unhide and edit this sheet prior to generating. This will ensure that each sheet has both the correct and identical user 'preset' data.
2. If the coin fee is consistent among a number of data entries, you can auto-populate the column by adding a formula to cell **"G4"**. Reference to 'Transaction Type' is likely the most efficient option and can be linked through the corresponding named range.  
For example:

Transaction Type | Named Range
------------ | -------------
Stake | cStake

```
= IF($D4="","",IF($F4=cStake,0,$D4*0.001))
```
3. Say you're not interested in 'Estimated Profit' data and would like to pull something else into the 'Master'. Through the module [mPullData.bas](https://github.com/CheddarBizcuitz/CryptoTrade/blob/main/VBA/mPullData.bas), you can alter the selected pull value and table column. The referenced cell should always relate to the 'Temp' worksheet. 

    **Note:** Referenced cell value(s) needs to be converted to text before pulling.

```
ABC100 = TEXT(Temp!A1,"$#,##0.00")
```

```VBA
    For Each sh In Sheets
        For Each c In Rng.Cells
            If sh.Name = c Then
                c.Offset(0, 1) = sh.Range("R3").Value ' Offset to right from 0 column 'Coin' (1 = Investment)
                c.Offset(0, 2) = sh.Range("S3").Value '                                      (2 = Coin Balance)
                c.Offset(0, 3) = sh.Range("T3").Value '                                      (3 = USD Balance)
                c.Offset(0, 4) = sh.Range("ABC100").Value '                                  (4 = 'Example')
                ... '                                                                        ...
            End If
        Next c
    Next sh
```

## Downloads

For Excel versions 2010 & 2013, the Power Query add-in can be downloaded [here](https://www.microsoft.com/en-us/download/details.aspx?id=39379).

## License

CryptoTrade is licensed under [GNU General Public License v3.0](https://github.com/CheddarBizcuitz/CryptoTrade/blob/main/LICENSE).

