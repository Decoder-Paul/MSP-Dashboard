Attribute VB_Name = "Staging"
Sub mainDataStaging()
'========================================================================================================
' Main Data for Staging
' -------------------------------------------------------------------------------------------------------
' Purpose   :   Contains formated raw data for the data extraction
'
' Author    :   Shambhavi B M, 5th January, 2018
' Notes     :   N/A
'
' Parameter :   N/A
' Returns   :   N/A
' -------------------------------------------------------------------------------------------------------
' Revision History
'
' -> Converting CreationDate,ActualstartDate,ActualfinishDate,CW Start Date,
'    CW End Date,Finish Date to date format
' -> Priority column converted from string to number format
'========================================================================================================

Dim RDlro As Long
Dim DAlro As Long
Dim R As Long
Dim c As Long

WS_DA.Activate

DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

'Clearing all data in MainData
If DAlro >= 1 Then
    WS_DA.Range("A1:Z" & (DAlro)).Clear
End If

WS_RD.Activate
RDlro = WS_RD.Cells(WS_RD.Rows.Count, "A").End(xlUp).Row

'Copying Raw Data to MainData
WS_RD.Range(Cells(R + 1, c + 2), Cells(R + RDlro, c + 19)).Copy
WS_DA.Select
WS_DA.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("A1").Select

DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

'Converting priority from string to number
WS_DA.Range(Cells(R + 2, c + 26), Cells(R + DAlro, c + 26)).Formula = "=NUMBERVALUE(LEFT(L2,1))"
WS_DA.Range(Cells(R + 1, c + 26), Cells(R + RDlro, c + 26)).Copy
WS_DA.Range("Z1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range(Cells(R + 2, c + 26), Cells(R + DAlro, c + 26)).Cut Destination:=WS_DA.Range("L2")

'CreationDate, ActualstartDate,Actualfinishdate,CW StartDate,CW Enddate are converted to Date format
WS_DA.Range(Cells(R + 2, c + 9), Cells(R + DAlro, c + 9)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, c + 10), Cells(R + DAlro, c + 10)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, c + 11), Cells(R + DAlro, c + 11)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, c + 16), Cells(R + DAlro, c + 16)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, c + 17), Cells(R + DAlro, c + 17)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

Application.CutCopyMode = False

'columns and rows alignment
WS_DA.Range(Cells(R + 1, c + 1), Cells(R + 1, c + 18)).Interior.Color = RGB(46, 139, 87)
WS_DA.Range(Cells(R + 1, c + 1), Cells(R + 1, c + 18)).RowHeight = 30
WS_DA.Range(Cells(R + 1, c + 1), Cells(R + 1, c + 18)).VerticalAlignment = xlCenter

Range(Cells(R + 1, c + 1), Cells(R + DAlro, c + 18)).Select
With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(R + 1, c + 1), Cells(R + DAlro, c + 18)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
Next BI
 
End Sub
