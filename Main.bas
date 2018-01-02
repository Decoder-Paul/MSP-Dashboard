'Author- Subhankar paul 
Attribute VB_Name = "Main"
Public ver1_stDt, ver1_enDt, ver2_stDt, ver2_enDt As Date
Public ver3_stDt, ver3_enDt, ver4_stDt, ver4_enDt As Date
    Public Function fSheetExists(sheetToFind As String) As Boolean
        On Error GoTo ErrorHandler
        Dim sheet As Worksheet
        fSheetExists = False
        For Each sheet In Worksheets
            If sheetToFind = sheet.Name Then
                fSheetExists = True
                Exit Function
            End If
        Next sheet
ErrorHandler:
    End Function
Sub InputDate()

Dim WB As Workbook
Dim WS_In As Worksheet
Dim WS_se As Worksheet

Set WB = ActiveWorkbook

Set WS_Home = WB.Sheets("Home")
Set WS_In = WB.Sheets("Consolidated Report")
Set WS_Sum = WB.Sheets("Summary")
WS_Home.Activate
WS_In.Activate
WS_Sum.Activate
WS_Home.Select
'Taking Input for start and end date for 4 different versions

ver1_stDt = Cells(5, 4).Value
ver1_enDt = Cells(5, 6).Value
ver2_stDt = Cells(7, 4).Value
ver2_enDt = Cells(7, 6).Value
ver3_stDt = Cells(9, 4).Value
ver3_enDt = Cells(9, 6).Value
ver4_stDt = Cells(11, 4).Value
ver4_enDt = Cells(11, 6).Value

Call ticketCount
End Sub

