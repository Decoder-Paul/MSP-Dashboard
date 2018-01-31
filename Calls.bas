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
Sub InputDate()
'========================================================================================================
' InputDate
' ------------------------------------------------------------------------------------------------------
' Purpose of this Function : Taking date input from home page as quarter
'
' Author : Subhankar Paul | 9th January, 2018
' Notes  : Procedure can take dates even from blank quarters in between
' Parameters :
' Returns :
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

    Dim i As Integer
    
    WS_HM.Select
    c = 0
    'quarters is a 2D array with 2 column start date and end date
    For i = 5 To 33
        If (Cells(i, 4).Value <> "" And Cells(i, 6).Value <> "") Then
            If Cells(i, 4).Value < Cells(i, 6).Value Then
                quarters(c, 0) = Cells(i, 4).Value
                quarters(c, 1) = Cells(i, 6).Value
                c = c + 1
            Else
                MsgBox "Start date can't be more than end date in Version " & (i + 1) / 2 - 2, vbExclamation
                End
            End If
            
        End If
        i = i + 1
    Next i
End Sub
Sub pCleanDB()
'Procedure to clean the Dashboard

    'Home sheet content cleansing
    'WS_HM.Range("D5:F33").ClearContents
    
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

If CSSlro > 48 Then
    'Selection and deletion of below dashboard except first quarter and above
    Rows("49:" & CSSlro).Select
    Selection.Delete Shift:=xlUp
End If

If c > 1 Then
    'Selection of top dashboard template and replicating based on the number of quarter
    Range("A34:W48").Select
    Selection.AutoFill Destination:=Range("A34:W" & (48 + (c - 1) * 15)), Type:=xlFillDefault
End If


    'fixing row height and column width
    Rows("34:" & CSSlro).RowHeight = 30
    Rows("34:" & CSSlro).ColumnWidth = 6
    
    'fixing width of column A,B,C and S
    Columns("A:B").ColumnWidth = 8
    Columns("C").ColumnWidth = 14
    Columns("S").ColumnWidth = 9
    
    'fixing row height and column width
    Rows("34:" & CSSlro).RowHeight = 30
    Rows("34:" & CSSlro).ColumnWidth = 6
    
    'fixing width of column A,B,C and S
    Columns("A:B").ColumnWidth = 8
    Columns("C").ColumnWidth = 14
    Columns("S").ColumnWidth = 9

End Sub

Sub CreateUniqueList()
'Populating the list of unique team names in Column V in Main Data Sheet
    WS_DA.Select
    Dim lastrow As Long
    lastrow = Cells(Rows.Count, "H").End(xlUp).Row
    ActiveSheet.Range("H1:H" & lastrow).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ActiveSheet.Range("V1"), Unique:=True
End Sub

Sub pCloseApp()
'========================================================================================================
' pCloseApp
' -------------------------------------------------------------------------------------------------------
' Purpose of this programm : To create a new sheet
'
' Author : Subhankar Paul 18th January, 2017
' Notes  : N/A
' Parameters : N/A
' Returns : N/A
' ---------------------------------------------------------------
' Revision History
'
'========================================================================================================

    'Hidding the File

    WS_DA.Visible = xlSheetHidden
    WS_CSS.Activate
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

End Sub

