

## 1️⃣ controlProc()

- This is the starting button.
- It checks Param_Sheet_last_su row by row.
- If column 2 has “Y”, it runs the full process.


## 2️⃣ mainProc()

- This is the brain.
- It:
- Builds LastTable
- Builds SuTable
- Builds ScriptTable
- Processes SU records
- Counts account types
- Saves output file
- Updates SU count sheet

## 3️⃣ makeLastTable()
- Opens log file
- Removes unwanted lines
- Stores clean data in LastTable

## 4️⃣ makeSuTable()
- Opens NEO file
- Copies first 10 columns
- Pastes into SuTable

## 5️⃣ makeScriptTable()

- Reads all script text files
- Extracts name parts
- Stores metadata in ScriptTable

## 6️⃣ editSU_Count()
- Updates summary comparison formulas
- Shows match / mismatch automatically



### You must have these sheets:
- Param_Sheet_last_su ← (Main control sheet)
- LastTable
- SuTable
- ScriptTable
- SU_Count
- Menu

- If even one is missing → macro will fail.


## ✅ 1️⃣ What INPUT sheets/files you must provide

- Your macro depends on one control sheet + external files.

### 🔹 A) Mandatory Excel Sheets Inside This Workbook

- You must have these sheets:
  - Param_Sheet_last_su ← (Main control sheet)
  - LastTable
  - SuTable
  - ScriptTable
  - SU_Count
  - Menu

### 🔹 B) Param_Sheet_last_su (VERY IMPORTANT)
- This sheet controls everything.
- It reads:

| Column | Purpose                                       |
| ------ | --------------------------------------------- |
| Col 2  | Execution flag → must be `Y` or `y`           |
| Col 4  | Folder path                                   |
| Col 5  | Sub folder / file prefix                      |
| Col 6  | LAST log file name                            |
| Col 7  | File suffix (used for NEO & output file name) |

- Macro starts reading from row 2.
- If column 2 = Y → it will process that row.

## ✅ 2️⃣ External Input Files Required 
- Macro automatically opens these files from the folder path you provide.


### 🔹 1) LAST Log File

- Opened in:

```vba
Workbooks.OpenText Filename:= path + file
```

📌 Comes from:
```vba
.Cells(paramCnt, 4)  → Folder
.Cells(paramCnt, 5)  → Prefix
.Cells(paramCnt, 6)  → Log file name
```

### 🔹 2) NEO File

- Opened in:
  ```vba
  Workbooks.Open Filename:= path + prefix + "NEO" + suffix
  ```

📌 Comes from:
```vba
.Cells(paramCnt, 4)
.Cells(paramCnt, 5)
"NEO"
.Cells(paramCnt, 7)
```

- Used for:
  - Creating SuTable
  - Copying first 10 columns
 
### 🔹 3) Script Files

- Pattern searched:
```vba
*script*.txt
```
- In the same folder:
  ```vba
  path + prefix + "*script*.txt"
  ```

- Used for:
- Creating ScriptTable
- Extracting script metadata


## ✅ 3️⃣ What OUTPUT You Will Get

- There are 3 main outputs.

### 🔹 OUTPUT 1: Processed SuTable File (MAIN OUTPUT)

- At end of mainProc():

```
ThisWorkbook.Worksheets("SuTable").Copy
ActiveWorkbook.SaveAs path + prefix + "SIN" + suffix
```

- 🔥 Output file name format:
```
[path][prefix]SIN[suffix]
```

- Example:
  ```
  C:\Logs\ABC_SIN_20240627.xlsx
  ```

- 📌 Output type:
  ```
  Excel Workbook (.xlsx)
  ```

- 📌 Stored:
  - Same folder where your input files are located.


### 🔹 OUTPUT 2: Updated SU_Count Sheet
- Sheet inside same workbook:
```
SU_Count
```

- It updates:
  - Match / mismatch formulas
  - X count
  - Service count
  - Other count
  - This is internal summary.
  - Not saved separately.


### 🔹 OUTPUT 3: Message Box

- If warning happened:
  → Shows Exclamation message

- If everything fine:
  → Shows Information message

## ✅ 4️⃣ What Gets Modified Inside Workbook

- During execution:

| Sheet       | Action            |
| ----------- | ----------------- |
| LastTable   | Deleted & Rebuilt |
| SuTable     | Deleted & Rebuilt |
| ScriptTable | Deleted & Rebuilt |
| SU_Count    | Updated           |

## ✅ 5️⃣ Full Flow Summary (Very Clear)

- Here is exactly what happens:
  ```vba
  controlProc()
     ↓
  Read Param_Sheet_last_su
       ↓
  If Y found
       ↓
  mainProc()
       ↓
  makeLastTable()   → From LAST log file
  makeSuTable()     → From NEO file
  makeScriptTable() → From script files
       ↓
  Process Accounts
       ↓
  Count X / Service / Other
       ↓
  Save New File → SIN file
       ↓
  Update SU_Count
       ↓
  Show Message

  ```


## ✅ 6️⃣ Final Checklist Before Running

- Make sure:
  - ✔ Param_Sheet_last_su filled properly
  - ✔ Column 2 has Y
  - ✔ Folder path correct
  - ✔ LAST file exists
  - ✔ NEO file exists
  - ✔ Script files exist
  - ✔ All required sheets exist


## ✅ Final Answer in One Line
- You provide:
  - Param_Sheet_last_su sheet
  - LAST log file
  - NEO file
  - Script files

- You get:
  - New Excel file named → SIN
  - Updated SU_Count summary
  - Match / mismatch results

- Stored in:
  - 👉 Same folder path from Param_Sheet_last_su

```mathematica
                    ┌──────────────────────────────┐
                    │  Param_Sheet_last_su (Excel) │
                    │  (Control Sheet)             │
                    └──────────────┬───────────────┘
                                   │
                                   │  Reads:
                                   │  - Folder Path
                                   │  - Prefix
                                   │  - LAST file
                                   │  - Suffix
                                   │
                                   ▼
        ┌─────────────────────────────────────────────────┐
        │                mainProc()                       │
        └─────────────────────────────────────────────────┘
                     │               │               │
                     ▼               ▼               ▼

        ┌───────────────┐   ┌───────────────┐   ┌────────────────┐
        │ LAST Log File │   │   NEO File    │   │ Script Files   │
        │ (Text file)   │   │ (Excel file)  │   │ *script*.txt   │
        └───────┬───────┘   └───────┬───────┘   └────────┬───────┘
                │                   │                      │
                ▼                   ▼                      ▼

        ┌───────────────┐   ┌───────────────┐   ┌────────────────┐
        │  LastTable    │   │   SuTable     │   │  ScriptTable   │
        │ (Sheet)       │   │  (Sheet)      │   │   (Sheet)      │
        └───────────────┘   └───────────────┘   └────────────────┘
                          \          │           /
                           \         │          /
                            \        ▼         /
                             ─── Processing ───
                                      │
                                      ▼

                     ┌──────────────────────────────┐
                     │   SIN Output File (Excel)    │
                     │  path + prefix + SIN + suffix│
                     └──────────────────────────────┘
                                      │
                                      ▼
                          SU_Count Sheet Updated
```


```vba
Option Explicit
' ============================================================
' Global Variable Declarations
' These variables are used across multiple procedures
' ============================================================

Dim vbExclamationFlg As Boolean   ' Flag to check if any warning occurred
Dim vbExclamationMsg As String    ' Stores warning message text

Dim paramCnt As Long              ' Loop counter for Param_Sheet_last_su

Dim xCnt As Long                  ' Counter for X type accounts
Dim serviceCnt As Long            ' Counter for service accounts
Dim otherCnt As Long              ' Counter for other accounts


' ============================================================
' Main Control Procedure
' Entry point of the entire process
' ============================================================
Sub controlProc()

    ' Work only inside this sheet
    With ThisWorkbook.Worksheets("Param_Sheet_last_su")
    
        ' Reset flags before starting
        vbExclamationFlg = False
        vbExclamationMsg = ""
        
        ' Start reading parameters from row 2
        paramCnt = 2
        
        ' Loop until column 2 becomes empty
        Do Until .Cells(paramCnt, 2).Value = ""
        
            ' If execution flag is Y or y
            If .Cells(paramCnt, 2).Value = "y" Or _
               .Cells(paramCnt, 2).Value = "Y" Then
            
                Application.ScreenUpdating = False
                
                ' Run main processing logic
                Call mainProc
                
                Application.ScreenUpdating = True
            End If
            
            paramCnt = paramCnt + 1
        
        Loop
        
        ' Show result message
        If vbExclamationFlg = True Then
            MsgBox vbExclamationMsg, vbExclamation
        Else
            MsgBox "Process Completed Successfully", vbInformation
        End If
        
        ' Return user to Menu sheet
        ThisWorkbook.Worksheets("Menu").Select
        ThisWorkbook.Worksheets("Menu").Cells(1, 1).Select
        
    End With

End Sub


' ============================================================
' Main Processing Procedure
' This controls full data processing flow
' ============================================================
Sub mainProc()

    Dim lastCnt As Long
    Dim lastAccount As String
    Dim lastAccountMulti As String
    Dim lastLogin As String
    Dim lastLogout As String
    
    Dim suCnt As Long
    Dim suHHMM As String
    Dim suFromTo
    Dim suFrom As String
    Dim suTo As String
    
    Dim scriptCnt As Long
    Dim scriptControlNumberMulti As String
    Dim scriptLogNameMulti As String
    
    Dim colonPosition As Long
    
    With ThisWorkbook.Worksheets("Param_Sheet_last_su")
    
        ' Step 1: Build required tables
        Call makeLastTable
        Call makeSuTable
        Call makeScriptTable
        
        ' Initialize counters
        suCnt = 1
        xCnt = 0
        serviceCnt = 0
        otherCnt = 0
        
        ' Loop through SuTable
        Do Until ThisWorkbook.Worksheets("SuTable").Cells(suCnt, 1).Value = ""
        
            ' Extract SU time and account information
            suHHMM = Left(Format(Replace(CDate(ThisWorkbook.Worksheets("SuTable").Cells(suCnt, 3).Value), ":", ""), "000000"), 4)
            
            suFromTo = Split(ThisWorkbook.Worksheets("SuTable").Cells(suCnt, 6).Value, "-")
            suFrom = suFromTo(0)
            suTo = suFromTo(1)
            
            ' Check if system account
            If Left(suFrom, 1) = "x" And _
               IsNumeric(Mid(suFrom, 2, 6)) = True Then
               
                xCnt = xCnt + 1
            
            Else
                ' Process last table data
                lastCnt = 1
                lastAccountMulti = ""
                
                Do Until ThisWorkbook.Worksheets("LastTable").Cells(lastCnt, 1).Value = ""
                
                    colonPosition = InStr(ThisWorkbook.Worksheets("LastTable").Cells(lastCnt, 1).Value, ":")
                    
                    If Left(ThisWorkbook.Worksheets("LastTable").Cells(lastCnt, 1).Value, 1) = "x" _
                       And IsNumeric(Mid(ThisWorkbook.Worksheets("LastTable").Cells(lastCnt, 1).Value, 2, 6)) = True Then
                        
                        lastAccount = Left(ThisWorkbook.Worksheets("LastTable").Cells(lastCnt, 1).Value, 7)
                        
                        ' Collect multiple accounts
                        If InStr(lastAccountMulti, lastAccount) = 0 Then
                            lastAccountMulti = lastAccountMulti & "," & lastAccount
                        End If
                        
                    End If
                    
                    lastCnt = lastCnt + 1
                
                Loop
                
            End If
            
            suCnt = suCnt + 1
        
        Loop
        
        ' Save final file
        Application.DisplayAlerts = False
        
        ThisWorkbook.Worksheets("SuTable").Copy
        ActiveWorkbook.SaveAs .Cells(paramCnt, 4).Value & _
                              .Cells(paramCnt, 5).Value & _
                              "SIN" & _
                              .Cells(paramCnt, 7).Value
        ActiveWorkbook.Close
        
        Application.DisplayAlerts = True
        
        ' Update SU count summary
        Call editSU_Count
        
    End With

End Sub


' ============================================================
' Create LastTable
' Reads input log file and extracts useful records
' ============================================================
Sub makeLastTable()

    Dim End_row_last As Integer
    Dim Loop_Cnt As Integer
    Dim Last_Cnt As Integer
    
    With ThisWorkbook.Worksheets("Param_Sheet_last_su")
    
        ThisWorkbook.Worksheets("LastTable").Select
        Cells.Select
        Selection.Delete
        
        Loop_Cnt = 1
        Last_Cnt = 1
        
        ' Open source file
        Workbooks.OpenText Filename:=.Cells(paramCnt, 4).Value & _
                                     .Cells(paramCnt, 5).Value & _
                                     .Cells(paramCnt, 6).Value
        
        End_row_last = Cells(1, 1).SpecialCells(xlLastCell).Row
        
        ' Extract useful rows
        Do Until Loop_Cnt > End_row_last
        
            If Left(Cells(Loop_Cnt, 1), 3) <> "[*-" Then
            
                ThisWorkbook.Worksheets("LastTable").Cells(Last_Cnt, 1) = Cells(Loop_Cnt, 1)
                Last_Cnt = Last_Cnt + 1
            
            End If
            
            Loop_Cnt = Loop_Cnt + 1
        
        Loop
        
        Workbooks(.Cells(paramCnt, 6).Value).Close
    
    End With

End Sub


' ============================================================
' Create SuTable
' Opens NEO file and copies first 10 columns
' ============================================================
Sub makeSuTable()

    With ThisWorkbook.Worksheets("Param_Sheet_last_su")
    
        ThisWorkbook.Worksheets("SuTable").Select
        Cells.Select
        Selection.Delete
        
        Workbooks.Open Filename:=.Cells(paramCnt, 4).Value & _
                                  .Cells(paramCnt, 5).Value & _
                                  "NEO" & _
                                  .Cells(paramCnt, 7).Value
        
        Range(Cells(1, 1), Cells(Cells(1, 1).End(xlDown).Row, 10)).Select
        Selection.Copy
        
        ThisWorkbook.Activate
        ThisWorkbook.Worksheets("SuTable").Select
        Cells(1, 1).Select
        
        ActiveSheet.Paste
        
        Application.CutCopyMode = False
        
        Workbooks("NEO" & .Cells(paramCnt, 7).Value).Close
    
    End With

End Sub


' ============================================================
' Create ScriptTable
' Reads script files and builds script metadata
' ============================================================
Sub makeScriptTable()

    Dim scriptCnt As Long
    Dim fName As String
    Dim fSplit
    
    With ThisWorkbook.Worksheets("Param_Sheet_last_su")
    
        ThisWorkbook.Worksheets("ScriptTable").Select
        Cells.Select
        Selection.Delete
        
        scriptCnt = 1
        
        fName = Dir(.Cells(paramCnt, 4).Value & _
                    .Cells(paramCnt, 5).Value & _
                    "*script*.txt", vbNormal)
        
        Do Until fName = ""
        
            fSplit = Split(fName, "_")
            
            ThisWorkbook.Worksheets("ScriptTable").Cells(scriptCnt, 1) = fSplit(0)
            ThisWorkbook.Worksheets("ScriptTable").Cells(scriptCnt, 5) = fSplit(1)
            
            scriptCnt = scriptCnt + 1
            fName = Dir()
        
        Loop
    
    End With

End Sub


' ============================================================
' Update SU Count Summary Sheet
' ============================================================
Sub editSU_Count()

    Dim j As Long
    
    With ThisWorkbook.Worksheets("SU_Count")
    
        ' Loop through columns and update match/mismatch formulas
        j = 2
        
        Do Until .Cells(1, j).Value = ""
        
            .Cells(11, j).Formula = _
                "=IF(R[-8]C=R[-4]C,""match"",""mismatch"")"
            
            j = j + 1
        
        Loop
    
    End With

End Sub
```
