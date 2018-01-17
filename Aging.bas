Attribute VB_Name = "Aging"
Sub agingCount(ByVal team As String)
'========================================================================================================
' Aging calculation for each team
' -------------------------------------------------------------------------------------------------------
' Purpose   :   Calculating aging for the Active tickets
'
' Author    :   Shambhavi B M, 11th January, 2018
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'PRB' are string constant
'
' Parameter :   N/A
' Returns   :   N/A
' -------------------------------------------------------------------------------------------------------
' Revision History
'
'========================================================================================================

'----------Incident----------
Dim Days0_1_INC(4) As Integer
Dim Days2_3_INC(4) As Integer
Dim Days4_5_INC(4) As Integer
Dim Days6_7_INC(4) As Integer
Dim Days8_14_INC(4) As Integer
Dim Days15_30_INC(4) As Integer
Dim Days31_60_INC(4) As Integer
Dim Days61_90_INC(4) As Integer
Dim Days_GT_90_INC(4) As Integer
Dim Days_Sum_INC(4) As Integer

'--------Service Request--------
Dim Days0_1_SRQ(4) As Integer
Dim Days2_3_SRQ(4) As Integer
Dim Days4_5_SRQ(4) As Integer
Dim Days6_7_SRQ(4) As Integer
Dim Days8_14_SRQ(4) As Integer
Dim Days15_30_SRQ(4) As Integer
Dim Days31_60_SRQ(4) As Integer
Dim Days61_90_SRQ(4) As Integer
Dim Days_GT_90_SRQ(4) As Integer
Dim Days_Sum_SRQ(4) As Integer

'--------Problem Statement--------
Dim Days0_1_PRB(4) As Integer
Dim Days2_3_PRB(4) As Integer
Dim Days4_5_PRB(4) As Integer
Dim Days6_7_PRB(4) As Integer
Dim Days8_14_PRB(4) As Integer
Dim Days15_30_PRB(4) As Integer
Dim Days31_60_PRB(4) As Integer
Dim Days61_90_PRB(4) As Integer
Dim Days_GT_90_PRB(4) As Integer
Dim Days_Sum_PRB(4) As Integer

Dim INC_Dict_P1, INC_Dict_P2, INC_Dict_P3, INC_Dict_P4andP5, INC_Dict_total As Object
Dim SRQ_Dict_P1, SRQ_Dict_P2, SRQ_Dict_P3, SRQ_Dict_P4andP5, SRQ_Dict_total As Object
Dim PRB_Dict_P1, PRB_Dict_P2, PRB_Dict_P3, PRB_Dict_P4andP5, PRB_Dict_total As Object

Set INC_Dict_P1 = CreateObject("scripting.dictionary")
Set INC_Dict_P2 = CreateObject("scripting.dictionary")
Set INC_Dict_P3 = CreateObject("scripting.dictionary")
Set INC_Dict_P4andP5 = CreateObject("scripting.dictionary")
Set INC_Dict_total = CreateObject("scripting.dictionary")

Set SRQ_Dict_P1 = CreateObject("scripting.dictionary")
Set SRQ_Dict_P2 = CreateObject("scripting.dictionary")
Set SRQ_Dict_P3 = CreateObject("scripting.dictionary")
Set SRQ_Dict_P4andP5 = CreateObject("scripting.dictionary")
Set SRQ_Dict_total = CreateObject("scripting.dictionary")

Set PRB_Dict_P1 = CreateObject("scripting.dictionary")
Set PRB_Dict_P2 = CreateObject("scripting.dictionary")
Set PRB_Dict_P3 = CreateObject("scripting.dictionary")
Set PRB_Dict_P4andP5 = CreateObject("scripting.dictionary")
Set PRB_Dict_total = CreateObject("scripting.dictionary")

Dim today As Variant
Dim DAlro As Long
Dim i As Long
Dim ticket_type As String
Dim priority As Long
Dim age_of_tkt As Variant
Dim element As Long

WS_CSS.Activate
today = DateOfreport

WS_DA.Activate
DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

'cleaning Aging column in mainData before calculating aging for team
WS_DA.Range("S2:S" & DAlro).Clear

'Aging calculation logic
For i = 2 To DAlro

    'Aging for Active tickets
    'checking whether Actualfinish date is empty
    If Cells(i, 25).Value = "" Then
        'checking whether Actual start date is not empty or not
        If Cells(i, 24).Value <> "" Then
                'if Actual start date is empty then taking dufference between DateOfreport and Actualstartdate
                Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 24).Value
        Else
            'if Actual start date is empty then taking the difference between DateOfreport and Creation date
            If Cells(i, 24).Value = "" Then
                    Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 23).Value
            End If
        End If
    Else
        'Aging for resolved tickets
        'checking whether Actualfinish date is not empty and should be greater than date of report
        If Cells(i, 25).Value <> "" Or Cells(i, 25).Value >= CLng(DateOfreport) Then
        'checking whether Actual start date is not empty or not
            If Cells(i, 24).Value <> "" Then
                'if Actual start date is empty then taking dufference between DateOfreport and Actualstartdate
                Cells(i, 19).Value = Cells(i, 25).Value - Cells(i, 24).Value
            Else
            'if Actual start date is empty then taking the difference between DateOfreport and Creation date
                If Cells(i, 24).Value = "" Then
                    Cells(i, 19).Value = Cells(i, 25).Value - Cells(i, 23).Value
                End If
            End If
        End If
    End If
Next i

'------------Counting aging based on team, ticket type and priority-----------------
For i = 2 To DAlro
    
    If i = 297 Then
        Debug.Print i
    End If
     'comparing team between maindata team and the parameter team for teamwise aging calculation
     If Cells(i, 8).Value = team Then
             
        ticket_type = Cells(i, 1).Value
        priority = Cells(i, 12).Value
        age_of_tkt = Cells(i, 19).Value
        element = Cells(i, 19).Value

        Select Case ticket_type
            Case "INC":
                Select Case priority
                    Case 1:
                            If Cells(i, 25).Value <> "" Then
                                If INC_Dict_P1.Exists(element) Then
                                    INC_Dict_P1.Item(element) = INC_Dict_P1.Item(element) + 1
                                Else
                                    INC_Dict_P1.Add element, 1
                                End If
                            End If
                    
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(0) = Days0_1_INC(0) + 1
                                    'sum of aging count for the priority 1 of incident tickets
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(0) = Days2_3_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(0) = Days4_5_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(0) = Days6_7_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(0) = Days8_14_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(0) = Days15_30_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(0) = Days31_60_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(0) = Days61_90_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(0) = Days_GT_90_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                End If
                            End If
                    Case 2:
                          If Cells(i, 25).Value <> "" Then
                                If INC_Dict_P2.Exists(element) Then
                                    INC_Dict_P2.Item(element) = INC_Dict_P2.Item(element) + 1
                                Else
                                    INC_Dict_P2.Add element, 1
                                End If
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(1) = Days0_1_INC(1) + 1
                                    'sum of aging count for the priority 2 of incident tickets
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(1) = Days2_3_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(1) = Days4_5_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(1) = Days6_7_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(1) = Days8_14_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(1) = Days15_30_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(1) = Days31_60_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(1) = Days61_90_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(1) = Days_GT_90_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If Cells(i, 25).Value <> "" Then
                                If INC_Dict_P3.Exists(element) Then
                                    INC_Dict_P3.Item(element) = INC_Dict_P3.Item(element) + 1
                                Else
                                    INC_Dict_P3.Add element, 1
                                End If
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(2) = Days0_1_INC(2) + 1
                                    'sum of aging count for the priority 3 of incident tickets
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(2) = Days2_3_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(2) = Days4_5_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(2) = Days6_7_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(2) = Days8_14_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(2) = Days15_30_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(2) = Days31_60_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(2) = Days61_90_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(2) = Days_GT_90_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If Cells(i, 25).Value <> "" Then
                            If INC_Dict_P4andP5.Exists(element) Then
                                INC_Dict_P4andP5.Item(element) = INC_Dict_P4andP5.Item(element) + 1
                            Else
                                INC_Dict_P4andP5.Add element, 1
                            End If
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(3) = Days0_1_INC(3) + 1
                                    'sum of aging count for the priority 4 and 5 of incident tickets
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(3) = Days2_3_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(3) = Days4_5_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(3) = Days6_7_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(3) = Days8_14_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(3) = Days15_30_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(3) = Days31_60_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(3) = Days61_90_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(3) = Days_GT_90_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                End If
                            End If
                                
                End Select
            Case "SRQ":
                Select Case priority
                    Case 1:
                            If SRQ_Dict_P1.Exists(element) Then
                                SRQ_Dict_P1.Item(element) = SRQ_Dict_P1.Item(element) + 1
                            Else
                                SRQ_Dict_P1.Add element, 1
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(0) = Days0_1_SRQ(0) + 1
                                    'sum of aging count for the priority 1 of service requests tickets
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(0) = Days2_3_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(0) = Days4_5_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(0) = Days6_7_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(0) = Days8_14_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(0) = Days15_30_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(0) = Days31_60_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(0) = Days61_90_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(0) = Days_GT_90_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                End If
                            End If
                    Case 2:
                            If SRQ_Dict_P2.Exists(element) Then
                                SRQ_Dict_P2.Item(element) = SRQ_Dict_P2.Item(element) + 1
                            Else
                                SRQ_Dict_P2.Add element, 1
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(1) = Days0_1_SRQ(1) + 1
                                    'sum of aging count for the priority 2 of service requests tickets
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(1) = Days2_3_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(1) = Days4_5_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(1) = Days6_7_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(1) = Days8_14_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(1) = Days15_30_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(1) = Days31_60_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(1) = Days61_90_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(1) = Days_GT_90_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If SRQ_Dict_P3.Exists(element) Then
                                SRQ_Dict_P3.Item(element) = SRQ_Dict_P3.Item(element) + 1
                            Else
                                SRQ_Dict_P3.Add element, 1
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(2) = Days0_1_SRQ(2) + 1
                                    'sum of aging count for the priority 3 of service requests tickets
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(2) = Days2_3_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(2) = Days4_5_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(2) = Days6_7_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(2) = Days8_14_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(2) = Days15_30_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(2) = Days31_60_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(2) = Days61_90_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(2) = Days_GT_90_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If SRQ_Dict_P4andP5.Exists(element) Then
                                SRQ_Dict_P4andP5.Item(element) = SRQ_Dict_P4andP5.Item(element) + 1
                            Else
                                SRQ_Dict_P4andP5.Add element, 1
                            End If
                            
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(3) = Days0_1_SRQ(3) + 1
                                    'sum of aging count for the priority 4 and 5 of service requests tickets
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(3) = Days2_3_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(3) = Days4_5_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(3) = Days6_7_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(3) = Days8_14_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(3) = Days15_30_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(3) = Days31_60_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(3) = Days61_90_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(3) = Days_GT_90_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                End If
                            End If
                End Select
            Case "PRB":
                Select Case priority
                    Case 1:
                            If PRB_Dict_P1.Exists(element) Then
                                PRB_Dict_P1.Item(element) = PRB_Dict_P1.Item(element) + 1
                            Else
                                PRB_Dict_P1.Add element, 1
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(0) = Days0_1_PRB(0) + 1
                                    'sum of aging count for the priority 1 of problem tickets
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(0) = Days2_3_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(0) = Days4_5_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(0) = Days6_7_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(0) = Days8_14_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(0) = Days15_30_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(0) = Days31_60_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(0) = Days61_90_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(0) = Days_GT_90_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                End If
                            End If
                    Case 2:
                            If PRB_Dict_P2.Exists(element) Then
                                PRB_Dict_P2.Item(element) = PRB_Dict_P2.Item(element) + 1
                            Else
                                PRB_Dict_P2.Add element, 1
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(1) = Days0_1_PRB(1) + 1
                                    'sum of aging count for the priority 2 of problem tickets
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(1) = Days2_3_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(1) = Days4_5_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(1) = Days6_7_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(1) = Days8_14_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(1) = Days15_30_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(1) = Days31_60_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(1) = Days61_90_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(1) = Days_GT_90_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If PRB_Dict_P3.Exists(element) Then
                                PRB_Dict_P3.Item(element) = PRB_Dict_P3.Item(element) + 1
                            Else
                                PRB_Dict_P3.Add element, 1
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(2) = Days0_1_PRB(2) + 1
                                    'sum of aging count for the priority 3 of problem tickets
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(2) = Days2_3_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(2) = Days4_5_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(2) = Days6_7_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(2) = Days8_14_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(2) = Days15_30_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(2) = Days31_60_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(2) = Days61_90_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(2) = Days_GT_90_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            
                            If PRB_Dict_P4andP5.Exists(element) Then
                                PRB_Dict_P4andP5.Item(element) = PRB_Dict_P4andP5.Item(element) + 1
                            Else
                                PRB_Dict_P4andP5.Add element, 1
                            End If
                            
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(3) = Days0_1_PRB(3) + 1
                                    'sum of aging count for the priority 4 and 5 of problem tickets
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(3) = Days2_3_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(3) = Days4_5_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(3) = Days6_7_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(3) = Days8_14_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(3) = Days15_30_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(3) = Days31_60_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(3) = Days61_90_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(3) = Days_GT_90_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                End If
                            End If
                End Select
        End Select
     End If
Next i

Dim highestVal As Integer, s As Integer
    For Each j In INC_Dict_P3.Keys
    Debug.Print INC_Dict_P3(j)
        If INC_Dict_P3(j) > highestVal Then
            highestVal = INC_Dict_P3(j)
            s = j
        End If
    Next j
'------------Counting aging based on team, ticket type and priority-----------------

WS_CSS.Select

'values insertion to the respective cells in the dashboard
'----------Incident----------
Range("D14:G14").Value = Days0_1_INC
Range("D15:G15").Value = Days2_3_INC
Range("D16:G16").Value = Days4_5_INC
Range("D17:G17").Value = Days6_7_INC
Range("D18:G18").Value = Days8_14_INC
Range("D19:G19").Value = Days15_30_INC
Range("D20:G20").Value = Days31_60_INC
Range("D21:G21").Value = Days61_90_INC
Range("D22:G22").Value = Days_GT_90_INC
Range("D23:G23").Value = Days_Sum_INC

'--------Service Request--------
Range("I14:L14").Value = Days0_1_SRQ
Range("I15:L15").Value = Days2_3_SRQ
Range("I16:L16").Value = Days4_5_SRQ
Range("I17:L17").Value = Days6_7_SRQ
Range("I18:L18").Value = Days8_14_SRQ
Range("I19:L19").Value = Days15_30_SRQ
Range("I20:L20").Value = Days31_60_SRQ
Range("I21:L21").Value = Days61_90_SRQ
Range("I22:L22").Value = Days_GT_90_SRQ
Range("I23:L23").Value = Days_Sum_SRQ

'--------Problem Statement--------
Range("N14:Q14").Value = Days0_1_PRB
Range("N15:Q15").Value = Days2_3_PRB
Range("N16:Q16").Value = Days4_5_PRB
Range("N17:Q17").Value = Days6_7_PRB
Range("N18:Q18").Value = Days8_14_PRB
Range("N19:Q19").Value = Days15_30_PRB
Range("N20:Q20").Value = Days31_60_PRB
Range("N21:Q21").Value = Days61_90_PRB
Range("N22:Q22").Value = Days_GT_90_PRB
Range("N23:Q23").Value = Days_Sum_PRB

'Total count of aging(Incidents) for all priority based on days segragation
Cells(14, 8).Formula = "=Sum(D14:G14)"
Cells(15, 8).Formula = "=Sum(D15:G15)"
Cells(16, 8).Formula = "=Sum(D16:G16)"
Cells(17, 8).Formula = "=Sum(D17:G17)"
Cells(18, 8).Formula = "=Sum(D18:G18)"
Cells(19, 8).Formula = "=Sum(D19:G19)"
Cells(20, 8).Formula = "=Sum(D20:G20)"
Cells(21, 8).Formula = "=Sum(D21:G21)"
Cells(22, 8).Formula = "=Sum(D22:G22)"

'Total count of aging(service requests) for all priority based on days segragation
Cells(14, 13).Formula = "=Sum(I14:L14)"
Cells(15, 13).Formula = "=Sum(I15:L15)"
Cells(16, 13).Formula = "=Sum(I16:L16)"
Cells(17, 13).Formula = "=Sum(I17:L17)"
Cells(18, 13).Formula = "=Sum(I18:L18)"
Cells(19, 13).Formula = "=Sum(I19:L19)"
Cells(20, 13).Formula = "=Sum(I20:L20)"
Cells(21, 13).Formula = "=Sum(I21:L21)"
Cells(22, 13).Formula = "=Sum(I22:L22)"

'Total count of aging(problem statements) for all priority based on days segragation
Cells(14, 18).Formula = "=Sum(N14:Q14)"
Cells(15, 18).Formula = "=Sum(N15:Q15)"
Cells(16, 18).Formula = "=Sum(N16:Q16)"
Cells(17, 18).Formula = "=Sum(N17:Q17)"
Cells(18, 18).Formula = "=Sum(N18:Q18)"
Cells(19, 18).Formula = "=Sum(N19:Q19)"
Cells(20, 18).Formula = "=Sum(N20:Q20)"
Cells(21, 18).Formula = "=Sum(N21:Q21)"
Cells(22, 18).Formula = "=Sum(N22:Q22)"

'-----Sum of Total column------------
'----------Incident------------
Cells(23, 8).Formula = "=sum(H14:H22)"

'--------Service Request--------
Cells(23, 13).Formula = "=sum(M14:M22)"

'--------Problem Statement--------
Cells(23, 18).Formula = "=sum(R14:R22)"

End Sub

Sub agingCountForAll()
'========================================================================================================
' Aging calculation for Consolidated report
' -------------------------------------------------------------------------------------------------------
' Purpose   :   Calculating aging for the Active tickets
'
' Author    :   Shambhavi B M, 11th January, 2018
' Notes     :   Different Ticket Types: 'INC', 'SRQ', 'PRB' are string constant
'
' Parameter :   N/A
' Returns   :   N/A
' -------------------------------------------------------------------------------------------------------
' Revision History
'
'========================================================================================================

'----------Incident----------
Dim Days0_1_INC(4) As Integer
Dim Days2_3_INC(4) As Integer
Dim Days4_5_INC(4) As Integer
Dim Days6_7_INC(4) As Integer
Dim Days8_14_INC(4) As Integer
Dim Days15_30_INC(4) As Integer
Dim Days31_60_INC(4) As Integer
Dim Days61_90_INC(4) As Integer
Dim Days_GT_90_INC(4) As Integer
Dim Days_Sum_INC(4) As Integer

'--------Service Request--------
Dim Days0_1_SRQ(4) As Integer
Dim Days2_3_SRQ(4) As Integer
Dim Days4_5_SRQ(4) As Integer
Dim Days6_7_SRQ(4) As Integer
Dim Days8_14_SRQ(4) As Integer
Dim Days15_30_SRQ(4) As Integer
Dim Days31_60_SRQ(4) As Integer
Dim Days61_90_SRQ(4) As Integer
Dim Days_GT_90_SRQ(4) As Integer
Dim Days_Sum_SRQ(4) As Integer

'--------Problem Statement--------
Dim Days0_1_PRB(4) As Integer
Dim Days2_3_PRB(4) As Integer
Dim Days4_5_PRB(4) As Integer
Dim Days6_7_PRB(4) As Integer
Dim Days8_14_PRB(4) As Integer
Dim Days15_30_PRB(4) As Integer
Dim Days31_60_PRB(4) As Integer
Dim Days61_90_PRB(4) As Integer
Dim Days_GT_90_PRB(4) As Integer
Dim Days_Sum_PRB(4) As Integer

Dim today As Variant
Dim DAlro As Long
Dim i As Long
Dim ticket_type As String
Dim priority As Long
Dim age_of_tkt As Variant

WS_CSS.Activate
today = DateOfreport

WS_DA.Activate
DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

'cleaning Aging column in mainData before calculating aging for team
WS_DA.Range("S2:S" & DAlro).Clear

'Aging calculation logic
For i = 2 To DAlro

    'checking whether Actualfinish date is not empty and should be greater than date of report
    If Cells(i, 25).Value <> "" Or Cells(i, 25).Value >= CLng(DateOfreport) Then
        'checking whether Actual start date is not empty or not
        If Cells(i, 24).Value <> "" Then
                'if Actual start date is empty then taking dufference between DateOfreport and Actualstartdate
                Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 24).Value
        Else
            'if Actual start date is empty then taking the difference between DateOfreport and Creation date
            If Cells(i, 24).Value = "" Then
                    Cells(i, 19).Value = CLng(DateOfreport) - Cells(i, 23).Value
            End If
        End If
    Else
        'if Actualfinish date empty aging is empty
        Cells(i, 19).Value = ""
    End If
Next i

'------------Counting aging based on team, ticket type and priority-----------------
For i = 2 To DAlro
                 
        ticket_type = Cells(i, 1).Value
        priority = Cells(i, 12).Value
        age_of_tkt = Cells(i, 19).Value

        Select Case ticket_type
            Case "INC":
                Select Case priority
                    Case 1:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(0) = Days0_1_INC(0) + 1
                                    'sum of aging count for the priority 1 of incident tickets
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(0) = Days2_3_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(0) = Days4_5_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(0) = Days6_7_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(0) = Days8_14_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(0) = Days15_30_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(0) = Days31_60_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(0) = Days61_90_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(0) = Days_GT_90_INC(0) + 1
                                    Days_Sum_INC(0) = Days_Sum_INC(0) + 1
                                End If
                            End If
                    Case 2:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(1) = Days0_1_INC(1) + 1
                                    'sum of aging count for the priority 2 of incident tickets
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(1) = Days2_3_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(1) = Days4_5_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(1) = Days6_7_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(1) = Days8_14_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(1) = Days15_30_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(1) = Days31_60_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(1) = Days61_90_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(1) = Days_GT_90_INC(1) + 1
                                    Days_Sum_INC(1) = Days_Sum_INC(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(2) = Days0_1_INC(2) + 1
                                    'sum of aging count for the priority 3 of incident tickets
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(2) = Days2_3_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(2) = Days4_5_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(2) = Days6_7_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(2) = Days8_14_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(2) = Days15_30_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(2) = Days31_60_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(2) = Days61_90_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(2) = Days_GT_90_INC(2) + 1
                                    Days_Sum_INC(2) = Days_Sum_INC(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_INC(3) = Days0_1_INC(3) + 1
                                    'sum of aging count for the priority 4 and 5 of incident tickets
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_INC(3) = Days2_3_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_INC(3) = Days4_5_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_INC(3) = Days6_7_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_INC(3) = Days8_14_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_INC(3) = Days15_30_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_INC(3) = Days31_60_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_INC(3) = Days61_90_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_INC(3) = Days_GT_90_INC(3) + 1
                                    Days_Sum_INC(3) = Days_Sum_INC(3) + 1
                                End If
                            End If
                                
                End Select
            Case "SRQ":
                Select Case priority
                    Case 1:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(0) = Days0_1_SRQ(0) + 1
                                    'sum of aging count for the priority 1 of service requests tickets
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(0) = Days2_3_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(0) = Days4_5_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(0) = Days6_7_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(0) = Days8_14_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(0) = Days15_30_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(0) = Days31_60_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(0) = Days61_90_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(0) = Days_GT_90_SRQ(0) + 1
                                    Days_Sum_SRQ(0) = Days_Sum_SRQ(0) + 1
                                End If
                            End If
                    Case 2:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(1) = Days0_1_SRQ(1) + 1
                                    'sum of aging count for the priority 2 of service requests tickets
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(1) = Days2_3_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(1) = Days4_5_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(1) = Days6_7_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(1) = Days8_14_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(1) = Days15_30_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(1) = Days31_60_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(1) = Days61_90_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(1) = Days_GT_90_SRQ(1) + 1
                                    Days_Sum_SRQ(1) = Days_Sum_SRQ(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(2) = Days0_1_SRQ(2) + 1
                                    'sum of aging count for the priority 3 of service requests tickets
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(2) = Days2_3_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(2) = Days4_5_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(2) = Days6_7_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(2) = Days8_14_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(2) = Days15_30_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(2) = Days31_60_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(2) = Days61_90_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(2) = Days_GT_90_SRQ(2) + 1
                                    Days_Sum_SRQ(2) = Days_Sum_SRQ(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_SRQ(3) = Days0_1_SRQ(3) + 1
                                    'sum of aging count for the priority 4 and 5 of service requests tickets
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_SRQ(3) = Days2_3_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_SRQ(3) = Days4_5_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_SRQ(3) = Days6_7_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_SRQ(3) = Days8_14_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_SRQ(3) = Days15_30_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_SRQ(3) = Days31_60_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_SRQ(3) = Days61_90_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_SRQ(3) = Days_GT_90_SRQ(3) + 1
                                    Days_Sum_SRQ(3) = Days_Sum_SRQ(3) + 1
                                End If
                            End If
                End Select
            Case "PRB":
                Select Case priority
                    Case 1:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(0) = Days0_1_PRB(0) + 1
                                    'sum of aging count for the priority 1 of problem tickets
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(0) = Days2_3_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(0) = Days4_5_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(0) = Days6_7_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(0) = Days8_14_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(0) = Days15_30_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(0) = Days31_60_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(0) = Days61_90_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(0) = Days_GT_90_PRB(0) + 1
                                    Days_Sum_PRB(0) = Days_Sum_PRB(0) + 1
                                End If
                            End If
                    Case 2:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(1) = Days0_1_PRB(1) + 1
                                    'sum of aging count for the priority 2 of problem tickets
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(1) = Days2_3_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(1) = Days4_5_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(1) = Days6_7_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(1) = Days8_14_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(1) = Days15_30_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(1) = Days31_60_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(1) = Days61_90_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(1) = Days_GT_90_PRB(1) + 1
                                    Days_Sum_PRB(1) = Days_Sum_PRB(1) + 1
                                End If
                            End If
                            
                    Case 3:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(2) = Days0_1_PRB(2) + 1
                                    'sum of aging count for the priority 3 of problem tickets
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(2) = Days2_3_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(2) = Days4_5_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(2) = Days6_7_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(2) = Days8_14_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(2) = Days15_30_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(2) = Days31_60_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(2) = Days61_90_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(2) = Days_GT_90_PRB(2) + 1
                                    Days_Sum_PRB(2) = Days_Sum_PRB(2) + 1
                                End If
                            End If
                    
                    Case 4 And 5:
                            If CStr(age_of_tkt) <> "" And age_of_tkt > 0 Then
                                If age_of_tkt >= 0 And age_of_tkt <= 1 Then
                                    Days0_1_PRB(3) = Days0_1_PRB(3) + 1
                                    'sum of aging count for the priority 4 and 5 of problem tickets
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 2 And age_of_tkt <= 3 Then
                                    Days2_3_PRB(3) = Days2_3_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 4 And age_of_tkt <= 5 Then
                                    Days4_5_PRB(3) = Days4_5_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 6 And age_of_tkt <= 7 Then
                                    Days6_7_PRB(3) = Days6_7_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 8 And age_of_tkt <= 14 Then
                                    Days8_14_PRB(3) = Days8_14_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 15 And age_of_tkt <= 30 Then
                                    Days15_30_PRB(3) = Days15_30_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 31 And age_of_tkt <= 60 Then
                                    Days31_60_PRB(3) = Days31_60_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt >= 61 And age_of_tkt <= 90 Then
                                    Days61_90_PRB(3) = Days61_90_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                ElseIf age_of_tkt > 90 Then
                                    Days_GT_90_PRB(3) = Days_GT_90_PRB(3) + 1
                                    Days_Sum_PRB(3) = Days_Sum_PRB(3) + 1
                                End If
                            End If
                End Select
        End Select
Next i

    
'------------Counting aging based on team, ticket type and priority-----------------

WS_CSS.Select
'values insertion to the respective cells in the dashboard
'----------Incident----------
Range("D14:G14").Value = Days0_1_INC
Range("D15:G15").Value = Days2_3_INC
Range("D16:G16").Value = Days4_5_INC
Range("D17:G17").Value = Days6_7_INC
Range("D18:G18").Value = Days8_14_INC
Range("D19:G19").Value = Days15_30_INC
Range("D20:G20").Value = Days31_60_INC
Range("D21:G21").Value = Days61_90_INC
Range("D22:G22").Value = Days_GT_90_INC
Range("D23:G23").Value = Days_Sum_INC

'--------Service Request--------
Range("I14:L14").Value = Days0_1_SRQ
Range("I15:L15").Value = Days2_3_SRQ
Range("I16:L16").Value = Days4_5_SRQ
Range("I17:L17").Value = Days6_7_SRQ
Range("I18:L18").Value = Days8_14_SRQ
Range("I19:L19").Value = Days15_30_SRQ
Range("I20:L20").Value = Days31_60_SRQ
Range("I21:L21").Value = Days61_90_SRQ
Range("I22:L22").Value = Days_GT_90_SRQ
Range("I23:L23").Value = Days_Sum_SRQ

'--------Problem Statement--------
Range("N14:Q14").Value = Days0_1_PRB
Range("N15:Q15").Value = Days2_3_PRB
Range("N16:Q16").Value = Days4_5_PRB
Range("N17:Q17").Value = Days6_7_PRB
Range("N18:Q18").Value = Days8_14_PRB
Range("N19:Q19").Value = Days15_30_PRB
Range("N20:Q20").Value = Days31_60_PRB
Range("N21:Q21").Value = Days61_90_PRB
Range("N22:Q22").Value = Days_GT_90_PRB
Range("N23:Q23").Value = Days_Sum_PRB

'Total count of aging(Incidents) for all priority based on days segragation
Cells(14, 8).Formula = "=Sum(D14:G14)"
Cells(15, 8).Formula = "=Sum(D15:G15)"
Cells(16, 8).Formula = "=Sum(D16:G16)"
Cells(17, 8).Formula = "=Sum(D17:G17)"
Cells(18, 8).Formula = "=Sum(D18:G18)"
Cells(19, 8).Formula = "=Sum(D19:G19)"
Cells(20, 8).Formula = "=Sum(D20:G20)"
Cells(21, 8).Formula = "=Sum(D21:G21)"
Cells(22, 8).Formula = "=Sum(D22:G22)"

'Total count of aging(service requests) for all priority based on days segragation
Cells(14, 13).Formula = "=Sum(I14:L14)"
Cells(15, 13).Formula = "=Sum(I15:L15)"
Cells(16, 13).Formula = "=Sum(I16:L16)"
Cells(17, 13).Formula = "=Sum(I17:L17)"
Cells(18, 13).Formula = "=Sum(I18:L18)"
Cells(19, 13).Formula = "=Sum(I19:L19)"
Cells(20, 13).Formula = "=Sum(I20:L20)"
Cells(21, 13).Formula = "=Sum(I21:L21)"
Cells(22, 13).Formula = "=Sum(I22:L22)"

'Total count of aging(problem statements) for all priority based on days segragation
Cells(14, 18).Formula = "=Sum(N14:Q14)"
Cells(15, 18).Formula = "=Sum(N15:Q15)"
Cells(16, 18).Formula = "=Sum(N16:Q16)"
Cells(17, 18).Formula = "=Sum(N17:Q17)"
Cells(18, 18).Formula = "=Sum(N18:Q18)"
Cells(19, 18).Formula = "=Sum(N19:Q19)"
Cells(20, 18).Formula = "=Sum(N20:Q20)"
Cells(21, 18).Formula = "=Sum(N21:Q21)"
Cells(22, 18).Formula = "=Sum(N22:Q22)"

'-----Sum of Total column------------
'----------Incident------------
Cells(23, 8).Formula = "=sum(H14:H22)"

'--------Service Request--------
Cells(23, 13).Formula = "=sum(M14:M22)"

'--------Problem Statement--------
Cells(23, 18).Formula = "=sum(R14:R22)"

End Sub
