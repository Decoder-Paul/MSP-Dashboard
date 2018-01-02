Attribute VB_Name = "Counting"
Public Rng As Range
Sub ticketCount()

'========================================================================================================
' TicketCount
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To get no. of Tickets depending on conditions from Data File to Dashboard .
'
' Author    :   Subhankar Paul, 9th February, 2017
    ' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'ACT', 'PRB' are string constant, gonna updated soon
'
' Parameter :   N/A
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
Application.ScreenUpdating = False
Application.DisplayAlerts = False


Dim sheetData As String
Dim sheetDbd As String
Dim BI As Variant
Dim R, lro As Long
Dim C As Long

sheetData = "Consolidated Report"
sheetDbd = "Summary"

'------------ Checking for the Data & Dashboard Sheets -----------
If fSheetExists(sheetData) = True Then
    Sheets(sheetData).Activate
    If fSheetExists(sheetDbd) = True Then
        Sheets(sheetDbd).Activate
    Else
        MsgBox "Dashboard Sheet doesn't Exist"
    End If
Else
    MsgBox "Data Sheet doesn't Exist"
End If

Sheets(sheetDbd).Select

'------------ Cleaning Previous Data from the cells -----------
Dim clean As Range

Set clean = Range("B4:K12")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B14:K22")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B24:K32")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B34:K42")
clean.Select
Selection.Cells.ClearContents

Set clean = Range("B44:K51")
clean.Select
Selection.Cells.ClearContents
'------------------ Range Selection ----------
Sheets(sheetData).Select
lro = Sheets(sheetData).Cells(Rows.Count, "A").End(xlUp).Row
Set Rng = Sheets(sheetData).Range("O2:O" & lro)
For i = 2 To lro
    Cells(i, 18).Value = CLng(Cells(i, 10).Value) 'Creation Date Converted to Integer
    If Cells(i, 12).Value <> "" Then
        Cells(i, 19).Value = CLng(Cells(i, 12).Value) 'Finish Date Converted to Integer
    End If
Next i
Call Trans_SRQ_ver1
Call Trans_INC_ver1
Call Trans_PRB_ver1
Call Trans_ACT_ver1
Call Atlas_SRQ_ver1
Call Atlas_INC_ver1
Call Atlas_PRB_ver1
Call Atlas_ACT_ver1

Call Trans_SRQ_ver2
Call Trans_INC_ver2
Call Trans_PRB_ver2
Call Trans_ACT_ver2
Call Atlas_SRQ_ver2
Call Atlas_INC_ver2
Call Atlas_PRB_ver2
Call Atlas_ACT_ver2

Call Trans_SRQ_ver3
Call Trans_INC_ver3
Call Trans_PRB_ver3
Call Trans_ACT_ver3
Call Atlas_SRQ_ver3
Call Atlas_INC_ver3
Call Atlas_PRB_ver3
Call Atlas_ACT_ver3

Call Trans_SRQ_ver4
Call Trans_INC_ver4
Call Trans_PRB_ver4
Call Trans_ACT_ver4
Call Atlas_SRQ_ver4
Call Atlas_INC_ver4
Call Atlas_PRB_ver4
Call Atlas_ACT_ver4


End Sub

Sub Trans_SRQ_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer/
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 2).Value = Count
     '   Sheets("Summary").Cells(4, 14).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 2).Value = Count
        Sheets("Summary").Cells(5, 14).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 2).Value = Count
        Sheets("Summary").Cells(6, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 2).Value = Count
        Sheets("Summary").Cells(7, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 2).Value = Count
        Sheets("Summary").Cells(8, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_INC_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 2).Value = Count
        Sheets("Summary").Cells(10, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 2).Value = Count
        Sheets("Summary").Cells(11, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 2).Value = Count
        Sheets("Summary").Cells(12, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 2).Value = Count
        Sheets("Summary").Cells(13, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 2).Value = Count
        Sheets("Summary").Cells(14, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_PRB_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 2).Value = Count
        Sheets("Summary").Cells(16, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 2).Value = Count
        Sheets("Summary").Cells(17, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 2).Value = Count
        Sheets("Summary").Cells(18, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 2).Value = Count
        Sheets("Summary").Cells(19, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 2).Value = Count
        Sheets("Summary").Cells(20, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Trans_ACT_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 2).Value = Count
        Sheets("Summary").Cells(22, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 2).Value = Count
        Sheets("Summary").Cells(23, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 2).Value = Count
        Sheets("Summary").Cells(24, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 2).Value = Count
        Sheets("Summary").Cells(25, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 2).Value = Count
        Sheets("Summary").Cells(26, 14).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 2).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 2).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 2).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_SRQ_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 7).Value = Count
        Sheets("Summary").Cells(4, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 7).Value = Count
        Sheets("Summary").Cells(5, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 7).Value = Count
        Sheets("Summary").Cells(6, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 7).Value = Count
        Sheets("Summary").Cells(7, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 7).Value = Count
        Sheets("Summary").Cells(8, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_INC_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 7).Value = Count
        Sheets("Summary").Cells(10, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 7).Value = Count
        Sheets("Summary").Cells(11, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 7).Value = Count
        Sheets("Summary").Cells(12, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 7).Value = Count
        Sheets("Summary").Cells(13, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 7).Value = Count
        Sheets("Summary").Cells(14, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_PRB_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 7).Value = Count
        Sheets("Summary").Cells(16, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 7).Value = Count
        Sheets("Summary").Cells(17, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 7).Value = Count
        Sheets("Summary").Cells(18, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 7).Value = Count
        Sheets("Summary").Cells(19, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 7).Value = Count
        Sheets("Summary").Cells(20, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_ACT_ver1()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 7).Value = Count
        Sheets("Summary").Cells(22, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 7).Value = Count
        Sheets("Summary").Cells(23, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 7).Value = Count
        Sheets("Summary").Cells(24, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 7).Value = Count
        Sheets("Summary").Cells(25, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 7).Value = Count
        Sheets("Summary").Cells(26, 19).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver1_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 7).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver1_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver1_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 7).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver1_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver1_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 7).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_SRQ_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer/
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 3).Value = Count
        Sheets("Summary").Cells(4, 15).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 3).Value = Count
        Sheets("Summary").Cells(5, 15).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 3).Value = Count
        Sheets("Summary").Cells(6, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 3).Value = Count
        Sheets("Summary").Cells(7, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 3).Value = Count
        Sheets("Summary").Cells(8, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 3).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 3).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 3).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_INC_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 3).Value = Count
        Sheets("Summary").Cells(10, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 3).Value = Count
        Sheets("Summary").Cells(11, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 3).Value = Count
        Sheets("Summary").Cells(12, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 3).Value = Count
        Sheets("Summary").Cells(13, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 3).Value = Count
        Sheets("Summary").Cells(14, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 3).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 3).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 3).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_PRB_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 3).Value = Count
        Sheets("Summary").Cells(16, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 3).Value = Count
        Sheets("Summary").Cells(17, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 3).Value = Count
        Sheets("Summary").Cells(18, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 3).Value = Count
        Sheets("Summary").Cells(19, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 3).Value = Count
        Sheets("Summary").Cells(20, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 3).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 3).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 3).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Trans_ACT_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 3).Value = Count
        Sheets("Summary").Cells(22, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 3).Value = Count
        Sheets("Summary").Cells(23, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 3).Value = Count
        Sheets("Summary").Cells(24, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 3).Value = Count
        Sheets("Summary").Cells(25, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 3).Value = Count
        Sheets("Summary").Cells(26, 15).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 3).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 3).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 3).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_SRQ_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 8).Value = Count
        Sheets("Summary").Cells(4, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 8).Value = Count
        Sheets("Summary").Cells(5, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 8).Value = Count
        Sheets("Summary").Cells(6, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 8).Value = Count
        Sheets("Summary").Cells(7, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 8).Value = Count
        Sheets("Summary").Cells(8, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 8).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 8).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 8).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_INC_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 8).Value = Count
        Sheets("Summary").Cells(10, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 8).Value = Count
        Sheets("Summary").Cells(11, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 8).Value = Count
        Sheets("Summary").Cells(12, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 8).Value = Count
        Sheets("Summary").Cells(13, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 8).Value = Count
        Sheets("Summary").Cells(14, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 8).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 8).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 8).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_PRB_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 8).Value = Count
        Sheets("Summary").Cells(16, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 8).Value = Count
        Sheets("Summary").Cells(17, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 8).Value = Count
        Sheets("Summary").Cells(18, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 8).Value = Count
        Sheets("Summary").Cells(19, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 8).Value = Count
        Sheets("Summary").Cells(20, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 8).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 8).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 8).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_ACT_ver2()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 8).Value = Count
        Sheets("Summary").Cells(22, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 8).Value = Count
        Sheets("Summary").Cells(23, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 8).Value = Count
        Sheets("Summary").Cells(24, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 8).Value = Count
        Sheets("Summary").Cells(25, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 8).Value = Count
        Sheets("Summary").Cells(26, 20).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver2_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 8).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver2_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver2_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 8).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver2_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver2_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 8).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Trans_SRQ_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer/
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 4).Value = Count
        Sheets("Summary").Cells(4, 16).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 4).Value = Count
        Sheets("Summary").Cells(5, 16).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 4).Value = Count
        Sheets("Summary").Cells(6, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 4).Value = Count
        Sheets("Summary").Cells(7, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 4).Value = Count
        Sheets("Summary").Cells(8, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 4).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 4).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 4).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_INC_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 4).Value = Count
        Sheets("Summary").Cells(10, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 4).Value = Count
        Sheets("Summary").Cells(11, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 4).Value = Count
        Sheets("Summary").Cells(12, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 4).Value = Count
        Sheets("Summary").Cells(13, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 4).Value = Count
        Sheets("Summary").Cells(14, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 4).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 4).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 4).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_PRB_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 4).Value = Count
        Sheets("Summary").Cells(16, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 4).Value = Count
        Sheets("Summary").Cells(17, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 4).Value = Count
        Sheets("Summary").Cells(18, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 4).Value = Count
        Sheets("Summary").Cells(19, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 4).Value = Count
        Sheets("Summary").Cells(20, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 4).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 4).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 4).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Trans_ACT_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 4).Value = Count
        Sheets("Summary").Cells(22, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 4).Value = Count
        Sheets("Summary").Cells(23, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 4).Value = Count
        Sheets("Summary").Cells(24, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 4).Value = Count
        Sheets("Summary").Cells(25, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 4).Value = Count
        Sheets("Summary").Cells(26, 16).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 4).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 4).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 4).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_SRQ_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 9).Value = Count
        Sheets("Summary").Cells(4, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 9).Value = Count
        Sheets("Summary").Cells(5, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 9).Value = Count
        Sheets("Summary").Cells(6, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 9).Value = Count
        Sheets("Summary").Cells(7, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 9).Value = Count
        Sheets("Summary").Cells(8, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 9).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 9).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 9).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_INC_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 9).Value = Count
        Sheets("Summary").Cells(10, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 9).Value = Count
        Sheets("Summary").Cells(11, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 9).Value = Count
        Sheets("Summary").Cells(12, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 9).Value = Count
        Sheets("Summary").Cells(13, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 9).Value = Count
        Sheets("Summary").Cells(14, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 9).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 9).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 9).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_PRB_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 9).Value = Count
        Sheets("Summary").Cells(16, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 9).Value = Count
        Sheets("Summary").Cells(17, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 9).Value = Count
        Sheets("Summary").Cells(18, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 9).Value = Count
        Sheets("Summary").Cells(19, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 9).Value = Count
        Sheets("Summary").Cells(20, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 9).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 9).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 9).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_ACT_ver3()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 9).Value = Count
        Sheets("Summary").Cells(22, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 9).Value = Count
        Sheets("Summary").Cells(23, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 9).Value = Count
        Sheets("Summary").Cells(24, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 9).Value = Count
        Sheets("Summary").Cells(25, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 9).Value = Count
        Sheets("Summary").Cells(26, 21).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver3_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 9).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver3_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver3_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 9).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver3_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver3_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 9).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_SRQ_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer/
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 5).Value = Count
        Sheets("Summary").Cells(4, 17).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 5).Value = Count
        Sheets("Summary").Cells(5, 17).Value = WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 5).Value = Count
        Sheets("Summary").Cells(6, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 5).Value = Count
        Sheets("Summary").Cells(7, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 5).Value = Count
        Sheets("Summary").Cells(8, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 5).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 5).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 5).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_INC_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 5).Value = Count
        Sheets("Summary").Cells(10, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 5).Value = Count
        Sheets("Summary").Cells(11, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 5).Value = Count
        Sheets("Summary").Cells(12, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 5).Value = Count
        Sheets("Summary").Cells(13, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 5).Value = Count
        Sheets("Summary").Cells(14, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 5).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 5).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 5).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Trans_PRB_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 5).Value = Count
        Sheets("Summary").Cells(16, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 5).Value = Count
        Sheets("Summary").Cells(17, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 5).Value = Count
        Sheets("Summary").Cells(18, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 5).Value = Count
        Sheets("Summary").Cells(19, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 5).Value = Count
        Sheets("Summary").Cells(20, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 5).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 5).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 5).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Trans_ACT_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Transformer
    .AutoFilter Field:=9, Criteria1:="Transformers"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 5).Value = Count
        Sheets("Summary").Cells(22, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 5).Value = Count
        Sheets("Summary").Cells(23, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 5).Value = Count
        Sheets("Summary").Cells(24, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 5).Value = Count
        Sheets("Summary").Cells(25, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 5).Value = Count
        Sheets("Summary").Cells(26, 17).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 5).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 5).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 5).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_SRQ_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All SRQ Tickets
    .AutoFilter Field:=2, Criteria1:="SRQ"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' SRQ-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(4, 10).Value = Count
        Sheets("Summary").Cells(4, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(5, 10).Value = Count
        Sheets("Summary").Cells(5, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        
        ' SRQ-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(6, 10).Value = Count
        Sheets("Summary").Cells(6, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))

        ' SRQ-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(7, 10).Value = Count
        Sheets("Summary").Cells(7, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' SRQ-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(10, 10).Value = Count
        Sheets("Summary").Cells(8, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' SRQ-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(8, 10).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' SRQ-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(9, 10).Value = Count
            
        .AutoFilter Field:=18
        
        ' SRQ-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(11, 10).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_INC_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="INC"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' INC-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(14, 10).Value = Count
        Sheets("Summary").Cells(10, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(15, 10).Value = Count
        Sheets("Summary").Cells(11, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(16, 10).Value = Count
        Sheets("Summary").Cells(12, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(17, 10).Value = Count
        Sheets("Summary").Cells(13, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' INC-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(20, 10).Value = Count
        Sheets("Summary").Cells(14, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' INC-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(18, 10).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' INC-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(19, 10).Value = Count
            
        .AutoFilter Field:=18
        
        ' INC-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(21, 10).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
Sub Atlas_PRB_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="PRB"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' PRB-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(24, 10).Value = Count
        Sheets("Summary").Cells(16, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(25, 10).Value = Count
        Sheets("Summary").Cells(17, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(26, 10).Value = Count
        Sheets("Summary").Cells(18, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(27, 10).Value = Count
        Sheets("Summary").Cells(19, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' PRB-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(30, 10).Value = Count
        Sheets("Summary").Cells(20, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' PRB-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(28, 10).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' PRB-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(29, 10).Value = Count
            
        .AutoFilter Field:=18
        
        ' PRB-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(31, 10).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub

Sub Atlas_ACT_ver4()
    
    Sheets("Consolidated Report").Range("A1:S1").Select
    Selection.AutoFilter
    '------------------ Filtering Data for Version 1 ----------------------
    With Selection
    ' For All Atlas
    .AutoFilter Field:=9, Criteria1:="Atlas"
    ' For All INC Tickets
    .AutoFilter Field:=2, Criteria1:="ACT"
        ' For Resolved Ticket
        ' end_Date >= 'Finish Date' >= start_Date
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        
        ' ACT-P1_Resolved
        .AutoFilter Field:=13, Criteria1:="1"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(34, 10).Value = Count
        Sheets("Summary").Cells(22, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P2_Resolved
        .AutoFilter Field:=13, Criteria1:="2"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(35, 10).Value = Count
        Sheets("Summary").Cells(23, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P3_Resolved
        .AutoFilter Field:=13, Criteria1:="3"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(36, 10).Value = Count
        Sheets("Summary").Cells(24, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-P4 & P5 Resolved
        .AutoFilter Field:=13, Criteria1:="4", Operator:=xlOr, Criteria2:="5"
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(37, 10).Value = Count
        Sheets("Summary").Cells(25, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        ' ACT-Total_Resolved
        .AutoFilter Field:=13
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(40, 10).Value = Count
        Sheets("Summary").Cells(26, 22).Value = Application.WorksheetFunction.Sum(Rng.SpecialCells(xlCellTypeVisible))
        .AutoFilter Field:=19
        
        ' ACT-Total_Opening Balance
        .AutoFilter Field:=18, Criteria1:="<" & CLng(ver4_stDt)
        .AutoFilter Field:=19, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(38, 10).Value = Count
        
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
        ' ACT-Total_Received
        .AutoFilter Field:=18, Criteria1:=">=" & CLng(ver4_stDt), Operator:=xlAnd, Criteria2:="<=" & CLng(ver4_enDt)
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(39, 10).Value = Count
            
        .AutoFilter Field:=18
        
        ' ACT-Total_Carry Forward
        .AutoFilter Field:=18, Criteria1:="<=" & CLng(ver4_enDt)
        .AutoFilter Field:=19, Criteria1:=">" & CLng(ver4_enDt), Operator:=xlOr, Criteria2:=""
        Count = ActiveSheet.AutoFilter.Range.Columns(5). _
                SpecialCells(xlCellTypeVisible).Count - 1
        Sheets("Summary").Cells(41, 10).Value = Count
        .AutoFilter Field:=18
        .AutoFilter Field:=19
        
    End With
    
End Sub
