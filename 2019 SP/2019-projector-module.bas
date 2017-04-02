Attribute VB_Name = "Module1"
Option Explicit
Type Riding
    Region As Integer
    RidingName As String * 55
    PartyVotes(5) As Long
    TotalVotes As Long
End Type
Type Region
    StartProv As Integer
    EndProv As Integer
    RegionName As String * 25
    PartyVotes(5) As Long
    TotalVotes As Long
End Type
Type Parliament
    Ridings(337) As Riding
End Type
Global P As Parliament
Global PartyNames$(5)
Global Regions(1 To 9) As Region
Global LastRegion%
Sub InitializeRidingsAndRegions()
    GetRidings
    Dim PrairieRegions%
    Dim FileCount%
    Dim RidingNo%
    Dim PartyCount%
    Dim TRegion%
    Dim Discard$
    Open "PartyNames.txt" For Input As #1
    For FileCount% = 1 To 6
        Input #1, PartyNames(FileCount - 1)
    Next
    Close #1
    Do
        PrairieRegions = Int(InputBox("Number of Divsions in the Prairies"))
        If PrairieRegions < 1 Or PrairieRegions > 3 Then MsgBox ("Please Enter a number between 1 and 3 based on the number of regions in the polling data")
    Loop While PrairieRegions < 1 Or PrairieRegions > 3
    Open "ProvincesMinMax.txt" For Input As #1
    
    For FileCount% = 2 To PrairieRegions
        Line Input #1, Discard
        Line Input #1, Discard
    Next
    LastRegion% = 4 + PrairieRegions
    For FileCount = 1 To LastRegion%
        Input #1, Regions(FileCount).RegionName
    Next
    For FileCount = 1 To LastRegion%
        Input #1, Regions(FileCount).StartProv
        Input #1, Regions(FileCount).EndProv
    Next
    Close #1
    For RidingNo% = 0 To 334
        TRegion = ReturnRegion(P.Ridings(RidingNo))
        For PartyCount = 0 To 5
            Regions(TRegion).PartyVotes(PartyCount) = Regions(TRegion).PartyVotes(PartyCount) + P.Ridings(RidingNo).PartyVotes(PartyCount)
        Next
        Regions(TRegion).TotalVotes = Regions(TRegion).TotalVotes + P.Ridings(RidingNo).TotalVotes
    Next
End Sub
Function ReturnRegion%(x As Riding)
    Dim RegionCount%
    For RegionCount% = 1 To LastRegion
        If x.Region >= Regions(RegionCount).StartProv And x.Region <= Regions(RegionCount).EndProv Then
            ReturnRegion = RegionCount
            Exit Function
        End If
    Next
End Function
Sub GetRidings()
    Open "LastElection.dat" For Random As #1 Len = 28730
    Get #1, 1, P
    Close #1
End Sub
Sub InitializeRidings()
    Dim TRidingName$
    Dim DiscardFields%
    Dim NewRiding As Boolean
    Dim PartyCount%
    Dim RidingCount% '
    RidingCount = -1
    Dim RidingCode$
    Dim Discard$
    Dim TVotes&
    Dim NameAndParty$
    Open "table_tableau12.CSV" For Input As #1
    Open "PartyNames2.txt" For Input As #2
    For PartyCount = 0 To 4
        Input #2, PartyNames$(PartyCount)
    Next
    Line Input #1, Discard$
    Do
        Input #1, Discard$
        Input #1, TRidingName
        Input #1, RidingCode
        Input #1, NameAndParty
        If RidingCount = -1 Then
            NewRiding = True
        ElseIf Left$(TRidingName, 55) <> Trim(P.Ridings(RidingCount).RidingName) Then
            NewRiding = True
        End If
        If NewRiding Then
            RidingCount% = RidingCount% + 1
            P.Ridings(RidingCount).Region = Val(Left$(RidingCode, 2))
            P.Ridings(RidingCount).RidingName = TRidingName
        End If
        If RidingCount > 245 And RidingCount < 250 Then MsgBox TRidingName + "**" + P.Ridings(RidingCount).RidingName
        NewRiding = False
        For DiscardFields = 5 To 6
            Input #1, Discard
        Next
        Input #1, TVotes
        For PartyCount = 0 To 4
            If Right$(NameAndParty, Len(PartyNames$(PartyCount))) = PartyNames$(PartyCount) Then Exit For
        Next
        P.Ridings(RidingCount).PartyVotes(PartyCount) = P.Ridings(RidingCount).PartyVotes(PartyCount) + TVotes
        P.Ridings(RidingCount).TotalVotes = P.Ridings(RidingCount).TotalVotes + TVotes
        Line Input #1, Discard
    Loop Until EOF(1)
    Close #1
    Open "LastElection.dat" For Random As #1 Len = 28730
    Put #1, 1, P
    Close #1
End Sub
