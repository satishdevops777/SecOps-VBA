# VBA Macro: Excel Data Merge and Sort

## 📂 Purpose

This macro automates the process of merging Excel data from multiple `.xlsx` files in a folder and outputs a sorted, formatted result in another Excel file. It is written in **VBA (Visual Basic for Applications)**.

---

## 🧠 Main Procedure: `ProcessMergeDatas`

```vb
Sub ProcessMergeDatas(outputFilePath As String)
```

### 🔄 Purpose

Open a folder, loop through `.xlsx` files, merge content into an output workbook, sort the result, and save it.

---

## 🔧 Variable Definitions

```vb
Dim folderPath As String, fileName As String
Dim wbInput As Workbook, wsInput As Worksheet
Dim wbOutput As Workbook, wsOutput As Worksheet
Dim inputData As Variant
Dim i As Long, j As Long, outputRow As Long
Dim lastRow As Long
Dim outputBU As String, currentFileBU As String
Dim userChoice As VbMsgBoxResult
```

- **folderPath**: Path to folder containing input files
- **fileName**: Current file being processed
- **wbInput/wsInput**: Input workbook/worksheet
- **wbOutput/wsOutput**: Output workbook/worksheet
- **inputData**: Data read from input sheet
- **outputBU/currentFileBU**: BU key extracted from filename

---

## 🗂 Folder Selection

```vb
Set fDialog = Application.FileDialog(msoFileDialogFolderPicker)
If fDialog.Show <> -1 Then Exit Sub
folderPath = fDialog.SelectedItems(1)
```

- Prompts user to select a folder.
- Exits if canceled.

---

## 📄 Output Workbook Handling

```vb
Set wbOutput = Workbooks.Open(outputFilePath)
Set wsOutput = wbOutput.Sheets(1)
outputBU = GetBUKeyFromName(Dir(outputFilePath))
```

- Opens output file
- Extracts BU key from output filename

---

## 🔄 Loop Through Input Files

```vb
fileName = Dir(folderPath & "*.xlsx")
Do While fileName <> ""
```

- Loops through all `.xlsx` files

### 💡 Check for Duplicate BU

```vb
currentFileBU = GetBUKeyFromName(fileName)
If currentFileBU = outputBU Then
    userChoice = MsgBox(..., vbYesNo)
    If userChoice = vbNo Then Exit Sub
End If
```

- Skips files with same BU as output unless user agrees

---

## 📅 Read Data from Input

```vb
Set wbInput = Workbooks.Open(folderPath & "\" & fileName)
Set wsInput = wbInput.Sheets(1)
inputData = wsInput.Range("A2:AC" & lastRow).Value
```

- Reads rows starting from A2 to ACx

---

## 🧮 Copy Valid Rows

```vb
For i = 1 To UBound(inputData, 1)
    If Trim(inputData(i, 8)) <> "" Or Trim(inputData(i, 9)) <> "" Then
        For j = 1 To 29
            wsOutput.Cells(outputRow, j).Value = inputData(i, j)
        Next j
        outputRow = outputRow + 1
    End If
Next i
```

- Filters rows with data in column 8 or 9
- Appends them to the output

---

## 🔄 Sort the Output

```vb
With wsOutput.Sort
    .SortFields.Clear
    .SortFields.Add Key:=wsOutput.Range("AB2:AB" & ...)
    .Apply
End With
```

- Sorts data by column `AB`

---

## 🪜 Finalize and Save

```vb
wsOutput.Columns("A:AC").AutoFit
wbOutput.Save
wbOutput.Close
MsgBox "マージ処理が完了しました。", vbInformation
```

- Autofits columns
- Saves and closes output file
- Displays "Merge process completed" in Japanese

---

## 🔧 Helper Function: `GetBUKeyFromName`

```vb
Function GetBUKeyFromName(fileName As String) As String
    Dim s As String, i As Long
    s = Split(fileName, ".")(0)
    For i = 1 To Len(s)
        If Not Mid(s, i, 1) Like "[A-Z]" Then Exit For
    Next i
    GetBUKeyFromName = Left(s, i - 1)
End Function
```

- Extracts uppercase prefix from filename
- E.g. `HRData2023.xlsx` ➔ `HR`

---
