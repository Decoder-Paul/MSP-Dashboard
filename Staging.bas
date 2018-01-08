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
Dim WB As Workbook
Dim WS_RD As Worksheet
Dim WS_DA As Worksheet
Dim sRd As String
Dim sDa As String
Dim RDlro As Long
Dim DAlro As Long
Dim R As Long
Dim C As Long

sRd = "Raw Data"
sDa = "MainData"

Set WB = ActiveWorkbook
Set WS_RD = WB.Sheets(sRd)
Set WS_DA = WB.Sheets(sDa)

WS_DA.Activate

DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

'Clearing all data in MainData
If DAlro >= 1 Then
    WS_DA.Range("A1:Z" & (DAlro)).Clear
End If

WS_RD.Activate
RDlro = WS_RD.Cells(WS_RD.Rows.Count, "A").End(xlUp).Row

'Copying Raw Data to MainData
WS_RD.Range(Cells(R + 1, C + 2), Cells(R + RDlro, C + 19)).Copy
WS_DA.Select
WS_DA.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

WS_DA.Range("A1").Select

DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

'Converting priority from string to number
WS_DA.Range(Cells(R + 2, C + 26), Cells(R + DAlro, C + 26)).Formula = "=NUMBERVALUE(LEFT(L2,1))"
WS_DA.Range(Cells(R + 1, C + 26), Cells(R + RDlro, C + 26)).Copy
WS_DA.Range("Z1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
WS_DA.Range(Cells(R + 2, C + 26), Cells(R + DAlro, C + 26)).Cut Destination:=WS_DA.Range("L2")

'CreationDate, ActualstartDate,Actualfinishdate,CW StartDate,CW Enddate are converted to Date format
WS_DA.Range(Cells(R + 2, C + 9), Cells(R + DAlro, C + 9)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, C + 10), Cells(R + DAlro, C + 10)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, C + 11), Cells(R + DAlro, C + 11)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, C + 16), Cells(R + DAlro, C + 16)).NumberFormat = "[$-14009]dd-mm-yyyy;@"
WS_DA.Range(Cells(R + 2, C + 17), Cells(R + DAlro, C + 17)).NumberFormat = "[$-14009]dd-mm-yyyy;@"

Application.CutCopyMode = False

'columns and rows alignment
WS_DA.Range(Cells(R + 1, C + 1), Cells(R + 1, C + 18)).Interior.Color = RGB(46, 139, 87)
WS_DA.Range(Cells(R + 1, C + 1), Cells(R + 1, C + 18)).RowHeight = 30
WS_DA.Range(Cells(R + 1, C + 1), Cells(R + 1, C + 18)).VerticalAlignment = xlCenter

Range(Cells(R + 1, C + 1), Cells(R + DAlro, C + 18)).Select
With Selection
    .Columns.AutoFit
End With

For Each BI In Array(xlEdgeTop, xlEdgeLeft, xlEdgeBottom, xlEdgeRight, xlInsideHorizontal, xlInsideVertical)
    With Range(Cells(R + 1, C + 1), Cells(R + DAlro, C + 18)).Borders(BI)
         .Weight = xlThin
         .Color = RGB(148, 138, 84)
    End With
Next BI
 
End Sub
