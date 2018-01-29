Attribute VB_Name = "ActiveticketsClosure"
Sub activeCount(ByVal team As String)
'========================================================================================================
' activeCount
' -------------------------------------------------------------------------------------------------------
' Purpose   :   To get no. of active(open) Tickets on the date of report
'
' Author    :   Subhankar Paul, 29th January, 2018
' Notes     :   . Different Ticket Types: 'INC', 'SRQ', 'ACT', 'PRB' are string constant
'               . ResponseSLA %, ResolutionSLA % is calculated at the end of the count are
'
' Parameter :   team is for team based report
' Returns   :   N/A
' ---------------------------------------------------------------
' Revision History
'========================================================================================================
    
    Dim active_Inc(4) As Integer
    Dim rspSla_Inc(4) As Integer
    Dim rspSlaPrcnt_Inc(4) As Integer
    Dim resSla_Inc(4) As Integer
    Dim resSlaPrcnt_Inc(4) As Integer
    
    Dim active_Srq(4) As Integer
    Dim rspSla_Srq(4) As Integer
    Dim rspSlaPrcnt_Srq(4) As Integer
    Dim resSla_Srq(4) As Integer
    Dim resSlaPrcnt_Srq(4) As Integer
    
    Dim active_Prb(4) As Integer
    Dim rspSla_Prb(4) As Integer
    Dim rspSlaPrcnt_Prb(4) As Integer
    Dim resSla_Prb(4) As Integer
    Dim resSlaPrcnt_Prb(4) As Integer
    
    Dim active_Chg(4) As Integer
    Dim winMiss_Chg(4) As Integer
    Dim winMissPrcnt_Chg(4) As Integer
    
    Dim Data_rowCount As Long
    Dim Data_i As Long
    Dim j As Long
    
    Dim tkt_type As String
    Dim rspSLA As String
    Dim resSLA As String
    Dim prty As Integer

    Dim createDate As Long
    Dim finishDate As Long

    WS_DA.Select

    Data_rowCount = ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row

    For Data_i = 2 To Data_rowCount
    '------------------ Filtering Data for TEAM ----------------------
        If Cells(Data_i, 8).Value = team Then
            '------ Active Ticket Filteration -------
            finishDate = Cells(Data_i, 25).Value
            If finishDate = "" Then
                tkt_type = Cells(Data_i, 1).Value ' Ticket Type
                prty = Cells(Data_i, 12).Value ' Priority
                rspSLA = Cells(Data_i, 2).Value
                resSLA = Cells(Data_i, 3).Value
                Select Case tkt_type
                'If Incident ticket type
                Case "INC"
                    Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        'Active tickets
                        active_Inc(0) = active_Inc(0) + 1
                        If rspSLA = "N" Then
                            rspSla_Inc(0) = rspSla_Inc(0) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Inc(0) = resSla_Inc(0) + 1
                        End If
                    Case 2
                        active_Inc(1) = active_Inc(1) + 1
                        If rspSLA = "N" Then
                            rspSla_Inc(1) = rspSla_Inc(1) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Inc(1) = resSla_Inc(1) + 1
                        End If
                    Case 3
                        active_Inc(2) = active_Inc(2) + 1
                        If rspSLA = "N" Then
                            rspSla_Inc(2) = rspSla_Inc(2) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Inc(2) = resSla_Inc(2) + 1
                        End If
                    Case 4
                        active_Inc(3) = active_Inc(3) + 1
                        If rspSLA = "N" Then
                            rspSla_Inc(3) = rspSla_Inc(3) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Inc(3) = resSla_Inc(3) + 1
                        End If
                    Case 5
                        active_Inc(3) = active_Inc(3) + 1
                        If rspSLA = "N" Then
                            rspSla_Inc(3) = rspSla_Inc(3) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Inc(3) = resSla_Inc(3) + 1
                        End If
                    End Select
                Case "SRQ"
                    Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        'Active tickets
                        active_Srq(0) = active_Srq(0) + 1
                        If rspSLA = "N" Then
                            rspSla_Srq(0) = rspSla_Srq(0) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Srq(0) = resSla_Srq(0) + 1
                        End If
                    Case 2
                        active_Srq(1) = active_Srq(1) + 1
                        If rspSLA = "N" Then
                            rspSla_Srq(1) = rspSla_Srq(1) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Srq(1) = resSla_Srq(1) + 1
                        End If
                    Case 3
                        active_Srq(2) = active_Srq(2) + 1
                        If rspSLA = "N" Then
                            rspSla_Srq(2) = rspSla_Srq(2) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Srq(2) = resSla_Srq(2) + 1
                        End If
                    Case 4
                        active_Srq(3) = active_Srq(3) + 1
                        If rspSLA = "N" Then
                            rspSla_Srq(3) = rspSla_Srq(3) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Srq(3) = resSla_Srq(3) + 1
                        End If
                    Case 5
                        active_Srq(3) = active_Srq(3) + 1
                        If rspSLA = "N" Then
                            rspSla_Srq(3) = rspSla_Srq(3) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Srq(3) = resSla_Srq(3) + 1
                        End If
                    End Select
                Case "PRB"
                    Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        'Active tickets
                        active_Prb(0) = active_Prb(0) + 1
                        If rspSLA = "N" Then
                            rspSla_Prb(0) = rspSla_Prb(0) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Prb(0) = resSla_Prb(0) + 1
                        End If
                    Case 2
                        active_Prb(1) = active_Prb(1) + 1
                        If rspSLA = "N" Then
                            rspSla_Prb(1) = rspSla_Prb(1) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Prb(1) = resSla_Prb(1) + 1
                        End If
                    Case 3
                        active_Prb(2) = active_Prb(2) + 1
                        If rspSLA = "N" Then
                            rspSla_Prb(2) = rspSla_Prb(2) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Prb(2) = resSla_Prb(2) + 1
                        End If
                    Case 4
                        active_Prb(3) = active_Prb(3) + 1
                        If rspSLA = "N" Then
                            rspSla_Prb(3) = rspSla_Prb(3) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Prb(3) = resSla_Prb(3) + 1
                        End If
                    Case 5
                        active_Prb(3) = active_Prb(3) + 1
                        If rspSLA = "N" Then
                            rspSla_Prb(3) = rspSla_Prb(3) + 1
                        End If
                        If resSLA = "N" Then
                            resSla_Prb(3) = resSla_Prb(3) + 1
                        End If
                    End Select
                Case "CHG"
                    Select Case prty
                    'Priority 1 and so on same for below also
                    Case 1
                        'Active tickets
                        active_Chg(0) = active_Chg(0) + 1
                        'Change Window Missed: Today > Change Window End Date
                        If today > Cells(Data_i, 17).Value Then
                            winMiss_Chg(0) = winMiss_Chg(0) + 1
                        End If
                    Case 2
                        active_Chg(1) = active_Chg(1) + 1
                        If today > Cells(Data_i, 17).Value Then
                            winMiss_Chg(1) = winMiss_Chg(1) + 1
                        End If
                    Case 3
                        active_Chg(2) = active_Chg(2) + 1
                        If today > Cells(Data_i, 17).Value Then
                            winMiss_Chg(2) = winMiss_Chg(2) + 1
                        End If
                    Case 4
                        active_Chg(3) = active_Chg(3) + 1
                        If today > Cells(Data_i, 17).Value Then
                            winMiss_Chg(3) = winMiss_Chg(3) + 1
                        End If
                    Case 5
                        active_Chg(3) = active_Chg(3) + 1
                        If today > Cells(Data_i, 17).Value Then
                            winMiss_Chg(3) = winMiss_Chg(3) + 1
                        End If
                    End Select
                End Select
            End If
        End If
    Next Data_i
    For i = 0 To 3
        active_Inc(4) = active_Inc(4) + active_Inc(i)
        active_Srq(4) = active_Srq(4) + active_Srq(i)
        active_Prb(4) = active_Prb(4) + active_Prb(i)
        active_Chg(4) = active_Chg(4) + active_Chg(i)
        
        resSla_Inc(4) = resSla_Inc(4) + resSla_Inc(i)
        rspSla_Inc(4) = rspSla_Inc(4) + rspSla_Inc(i)
        
        resSla_Srq(4) = resSla_Srq(4) + resSla_Srq(i)
        rspSla_Srq(4) = rspSla_Srq(4) + rspSla_Srq(i)
        
        resSla_Prb(4) = resSla_Prb(4) + resSla_Prb(i)
        rspSla_Prb(4) = rspSla_Prb(4) + rspSla_Prb(i)
        
        winMiss_Chg(4) = winMiss_Chg(4) + winMiss_Chg(i)
    Next i
    For i = 0 To 4
        If active_Inc(i) <> 0 Then
            rspSlaPrcnt_Inc(i) = rspSla_Inc(i) * 100 / active_Inc(i)
            resSlaPrcnt_Inc(i) = resSla_Inc(i) * 100 / active_Inc(i)
        End If
        If active_Srq(i) <> 0 Then
            rspSlaPrcnt_Srq(i) = rspSla_Srq(i) * 100 / active_Srq(i)
            resSlaPrcnt_Srq(i) = resSla_Srq(i) * 100 / active_Srq(i)
        End If
        If active_Prb(i) <> 0 Then
            rspSlaPrcnt_Prb(i) = rspSla_Prb(i) * 100 / active_Prb(i)
            resSlaPrcnt_Prb(i) = resSla_Prb(i) * 100 / active_Prb(i)
        End If
        If active_Chg(i) <> 0 Then
            winMissPrcnt_Chg(i) = winMiss_Chg(i) * 100 / active_Chg(i)
        End If
    Next i
    WS_CSS.Select
    
    '------------------ VERSIONWISE Value Placement of the Variable in Excel sheet -------------
    '----------Incident----------
    Range("D5:H5").Value = active_Inc
    Range("D6:H6").Value = resSla_Inc
    Range("D7:H7").Value = rspSla_Inc
    Range("D8:H8").Value = resSlaPrcnt_Inc
    Range("D9:H9").Value = rspSlaPrcnt_Inc

    Range("I5:M5").Value = active_Srq
    Range("I6:M6").Value = resSla_Srq
    Range("I7:M7").Value = rspSla_Srq
    Range("I8:M8").Value = resSlaPrcnt_Srq
    Range("I9:M9").Value = rspSlaPrcnt_Srq

    Range("N5:R5").Value = active_Prb
    Range("N6:R6").Value = resSla_Prb
    Range("N7:R7").Value = rspSla_Prb
    Range("N8:R8").Value = resSlaPrcnt_Prb
    Range("N9:R9").Value = rspSlaPrcnt_Prb
    
    Range("T5:X5").Value = active_Chg
    Range("T7:X7").Value = winMiss_Chg
    Range("T9:X9").Value = winMissPrcnt_Chg
End Sub

