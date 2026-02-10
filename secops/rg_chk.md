## 📌 What this macro does (high level)

- This VBA code performs a regression check between two CSV files:

```
Accounted.csv
NEOAccounted.csv
```

It:
- Reads both CSVs
- Extracts account name + time
- Compares records between both
- Marks OK / NG
- Writes the final judgement into an Excel sheet
- Shows popup alerts if mismatches are found
- Japanese comments/messages are preserved exactly.

### 📥 INPUT FILES

| Output                   | Where                 |
| ------------------------ | --------------------- |
| OK / NG judgement        | `CHK_Accounted` sheet |
| Color-coded system names | Excel cells           |
| Popup alerts             | VBA `MsgBox`          |
| No new CSV files created | Only read             |


### 🧠 Key Logic Summary

- Matching is done on:
  - Account name
  - Time (hh:mm)
- If Accounted exists but not in NEO → ❌ NG
- If match found → ✅ OK

```
Option Explicit
# Forces explicit variable declaration

Public Acc_time(999) As String
# Stores time values from Accounted.csv

Public Acc_Account(999) As String
# Stores account names from Accounted.csv

Public Acc_Idx As Integer
# Index counter for Accounted arrays

Public NEOAcc_time(999) As String
# Stores time values from NEOAccounted.csv

Public NEOAcc_Account(999) As String
# Stores account names from NEOAccounted.csv

Public NEOAcc_Cheked(999) As String
# Stores OK/NG result for NEO records

Public NEOAcc_Idx As Integer
# Index counter for NEO arrays

Public Judgement As String
# Stores final judgement OK/NG

Public In_File_name_Accounted As String
# Input file name for Accounted CSV

Sub regression_chk()

    ' Accounted.csvに存在し、NEOAccounted.csvに存在しないパターンを防ぐための確認
    # Regression check to detect records present in Accounted but missing in NEO

    Dim loop_i As Integer
    # Loop index for parameter sheet

    Dim Wk_Activecell As String
    # Used as loop break marker

    Dim Chk_Ctl As String
    # Control flag (Y/N)

    Dim Hit_Flg As String
    # Match found flag

    Dim Accounted_Idx As Byte
    # Index for final result arrays

    Dim Accounted_System(50) As String
    # Stores system names

    Dim Accounted_System_Color(50) As Long
    # Stores system colors

    Dim Accounted_FileName(50) As String
    # Stores processed file names

    Dim Accounted_Judgement(50) As String
    # Stores final judgement per system

    Dim In_File_name_NEOAccounted As String
    # Input file name for NEOAccounted

    Wk_Activecell = "9999"
    # Dummy initial value

    loop_i = 2
    # Start reading from row 2

    Accounted_Idx = 0
    # Reset result index

    Erase Accounted_System
    Erase Accounted_System_Color
    Erase Accounted_Judgement
    # Clear result arrays

    Application.ScreenUpdating = False
    # Performance optimization

    Do Until Wk_Activecell = ""

        In_Systeme = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1)
        # Read system name

        In_Systeme_Color = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1).Interior.Color
        # Read system color

        Chk_Ctl = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 2)
        # Read control flag

        If (Chk_Ctl = "Y" Or Chk_Ctl = "y") Then
            # Process only enabled rows

            In_Path1 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 4)
            In_Path2 = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 5)
            In_File_name_Accounted = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 7)
            In_File_name_NEOAccounted = "NEO" & In_File_name_Accounted
            # Build input paths

            Call Sec_Accounted_Spool
            # Read Accounted.csv

            Call Sec_NEOAccounted_Spool
            # Read NEOAccounted.csv

            Acc_Idx = 0
            # Reset Accounted index

            Judgement = "OK"
            # Default judgement

            Do Until Acc_time(Acc_Idx) = ""

                NEOAcc_Idx = 0
                Hit_Flg = ""
                # Reset search state

                Do Until NEOAcc_time(NEOAcc_Idx) = "" Or Hit_Flg = "1"

                    If NEOAcc_Cheked(NEOAcc_Idx) <> "OK" Then
                        # Skip already matched entries

                        If (Acc_time(Acc_Idx) = NEOAcc_time(NEOAcc_Idx)) And _
                           (Acc_Account(Acc_Idx) = NEOAcc_Account(NEOAcc_Idx)) Then

                            NEOAcc_Cheked(NEOAcc_Idx) = "OK"
                            # Mark as matched

                            Hit_Flg = "1"
                            # Match found
                        End If
                    End If

                    NEOAcc_Idx = NEOAcc_Idx + 1
                Loop

                If Hit_Flg <> "1" Then
                    Judgement = "NG"
                    # Mark failure

                    MsgBox "確認要あり：" & In_File_name_Accounted & _
                           " の " & Acc_time(Acc_Idx) & " を確認願います。"
                    # Alert user
                End If

                Acc_Idx = Acc_Idx + 1
            Loop

            Accounted_System(Accounted_Idx) = In_Systeme
            Accounted_System_Color(Accounted_Idx) = In_Systeme_Color
            Accounted_FileName(Accounted_Idx) = "NEO" & In_File_name_Accounted
            Accounted_Judgement(Accounted_Idx) = Judgement
            Accounted_Idx = Accounted_Idx + 1
        End If

        loop_i = loop_i + 1
        Wk_Activecell = ThisWorkbook.Sheets("Param_Sheet_last_su").Cells(loop_i, 1)
    Loop

    Sheets("CHK_Accounted").Select
    Range("A1").Select
    # Output results

    Accounted_Idx = 0

    Do Until Accounted_System(Accounted_Idx) = ""

        ActiveCell.Offset(0, 0) = Accounted_System(Accounted_Idx)
        ActiveCell.Offset(0, 0).Interior.Color = Accounted_System_Color(Accounted_Idx)
        ActiveCell.Offset(1, 0) = Accounted_FileName(Accounted_Idx)

        If Accounted_Judgement(Accounted_Idx) = "NG" Then
            ActiveCell.Offset(2, 0) = "確認要"
        Else
            ActiveCell.Offset(2, 0) = "OK"
        End If

        ActiveCell.Offset(0, 1).Select
        Accounted_Idx = Accounted_Idx + 1
    Loop

    Sheets("Menu").Select
    MsgBox "fin"
    # End message

End Sub
```
