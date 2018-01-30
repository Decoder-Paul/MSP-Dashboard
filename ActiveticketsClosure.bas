Attribute VB_Name = "ActiveticketsClosure"
Sub activeClosureDuration(ByVal team As String)
'========================================================================================================
' Closure duration calculation of closed tickets for each team
' -------------------------------------------------------------------------------------------------------
' Purpose   :   Calculating closure duration of closed tickets for each team
'
' Author    :   Shambhavi B M, 22nd January, 2018
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'PRB' are string constant
'
' Parameter :   N/A
' Returns   :   N/A
' -------------------------------------------------------------------------------------------------------
' Revision History
'
'========================================================================================================

Dim DAlro As Long
Dim i As Long
Dim ticket_type As String
Dim priority As Long
Dim closedtckts_age As Variant
Dim Sum As Long

WS_CSS.Activate

WS_DA.Activate
DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

For i = 2 To DAlro
    
     'comparing team between maindata team and the parameter team for teamwise aging calculation
     If Cells(i, 8).Value = team Then
             
        ticket_type = Cells(i, 1).Value
        priority = Cells(i, 12).Value
        closedtckts_age = Cells(i, 19).Value
        
        If Cells(i, 25).Value <> "" Then
        
        Select Case ticket_type
            Case "INC":
                Select Case priority
                    Case 1:
                            closed_INC_P1 (i)
10                    Case 2:
                    Case 3:
                    Case 4 And 5:
                End Select
            Case "SRQ":
                Select Case priority
                    Case 1:
                    Case 2:
                    Case 3:
                    Case 4 And 5:
                End Select
            Case "PRB":
                Select Case priority
                    Case 1:
                    Case 2:
                    Case 3:
                    Case 4 And 5:
                End Select
        End Select
End Sub
