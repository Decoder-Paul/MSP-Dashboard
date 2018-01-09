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
