Attribute VB_Name = "Aging"
Sub agingCount(ByVal team As String)
'========================================================================================================
' Main Data for Staging
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

WS_CSS.Activate
today = CLng(Cells(5, 2).Value)

WS_DA.Activate
DAlro = WS_DA.Cells(WS_DA.Rows.Count, "A").End(xlUp).Row

For i = 2 To DAlro
    If Cells(i, 24).Value = "" Then
        Cells(i, 24).Value = Cells(i, 23).Value
    End If
    Cells(i, 19).Value = today - Cells(i, 24).Value
Next i

Sheets("MainData").Range("A1:AA1").Select
Selection.AutoFilter

With Selection
    
     .AutoFilter Field:=19, Criteria1:=">=" & CLng(quarters(c - 1, 1))
     '------------------ Filtering Data for TEAM ----------------------
     .AutoFilter Field:=8, Criteria1:=team

End With


'values insertion to the respective cells in the dashboard
'----------Incident----------
Range("D14:H14").Value = Days0_1_INC
Range("D15:H15").Value = Days2_3_INC
Range("D16:H16").Value = Days4_5_INC
Range("D17:H17").Value = Days6_7_INC
Range("D18:H18").Value = Days8_14_INC
Range("D19:H19").Value = Days15_30_INC
Range("D20:H20").Value = Days31_60_INC
Range("D21:H21").Value = Days61_90_INC
Range("D22:H22").Value = Days_GT_90_INC
Range("D23:H23").Value = Days_Sum_INC

'--------Service Request--------
Range("I14:M14").Value = Days0_1_SRQ
Range("I15:M15").Value = Days2_3_SRQ
Range("I16:M16").Value = Days4_5_SRQ
Range("I17:M17").Value = Days6_7_SRQ
Range("I18:M18").Value = Days8_14_SRQ
Range("I19:M19").Value = Days15_30_SRQ
Range("I20:M20").Value = Days31_60_SRQ
Range("I21:M21").Value = Days61_90_SRQ
Range("I22:M22").Value = Days_GT_90_SRQ
Range("I23:M23").Value = Days_Sum_SRQ

'--------Problem Statement--------
Range("N14:R14").Value = Days0_1_PRB
Range("N15:R15").Value = Days2_3_PRB
Range("N16:R16").Value = Days4_5_PRB
Range("N17:R17").Value = Days6_7_PRB
Range("N18:R18").Value = Days8_14_PRB
Range("N19:R19").Value = Days15_30_PRB
Range("N20:R20").Value = Days31_60_PRB
Range("N21:R21").Value = Days61_90_PRB
Range("N22:R22").Value = Days_GT_90_PRB
Range("N23:R23").Value = Days_Sum_PRB

End Sub
