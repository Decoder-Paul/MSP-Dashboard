Attribute VB_Name = "Calls"
Public Function fSheetExists(sheetToFind As String) As Boolean
'========================================================================================================
' fSheetExists
' -------------------------------------------------------------------------------------------------------
' Purpose of this Function : To check if a sheet is existing or not
'
' Author : Subhankar Paul | 8th January, 2018
' Notes  : Public function
' Parameters : sSheetNameIN - Sheet Name
' Returns : True/False
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================
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
Sub pCleanDB()
'Procedure to clean the Dashboard

    'Home sheet content cleansing
    WS_HM.Range("D5:F33").ClearContents
    
    'Support Dashboard contents cleansing
    'Active Ticket's stat table
    WS_CSS.Range("D5:R9").ClearContents
    WS_CSS.Range("T5:X9").ClearContents
    
    'Aging Data Table
    WS_CSS.Range("D14:R23").ClearContents
    WS_CSS.Range("D28:R28").ClearContents
    
    'Quarter Stats Table
    WS_CSS.Range("D34:W48").ClearContents
    
End Sub

Sub CreateUniqueList()
'Populating the list of unique team names in Column V in Main Data Sheet
    WS_DA.Select
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "H").End(xlUp).Row
    ActiveSheet.Range("H1:H" & lastrow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("V1"), Unique:=True
End Sub

Sub QtrReplication()

'========================================================================================================
' Main Data for Staging
' -------------------------------------------------------------------------------------------------------
' Purpose   :   Qruarter replication for quarter based report generation
'
' Author    :   Shambhavi B M, 9th January, 2018
' Notes     :   N/A
'
' Parameter :   N/A
' Returns   :   N/A
' -------------------------------------------------------------------------------------------------------
' Revision History
'
'========================================================================================================

Dim lro As Long
Dim R As Long

WS_CSS.Select
CSSlro = WS_CSS.Cells(WS_CSS.Rows.Count, "C").End(xlUp).Row

If c > 1 And CSSlro > 48 Then
    
    'Selection and deletion of below dashboard except top
    Rows("49:" & CSSlro).Select
    Selection.Delete Shift:=xlUp
    
    'Selection of top dashboard template and replicating based on the number of quarter
    Range("A34:W48").Select
    Selection.AutoFill Destination:=Range("A34:W" & (48 + c * 15)), Type:=xlFillDefault

    'fixing row height and column width
    Rows("34:" & CSSlro).RowHeight = 30
    Rows("34:" & CSSlro).ColumnWidth = 6
    
    'fixing width of column A,B,C and S
    Columns("A:B").ColumnWidth = 8
    Columns("C").ColumnWidth = 14
    Columns("S").ColumnWidth = 9
    
End If
    
    'fixing row height and column width
    Rows("34:" & CSSlro).RowHeight = 30
    Rows("34:" & CSSlro).ColumnWidth = 6
    
    'fixing width of column A,B,C and S
    Columns("A:B").ColumnWidth = 8
    Columns("C").ColumnWidth = 14
    Columns("S").ColumnWidth = 9

End Sub
