Attribute VB_Name = "Module1"
Option Explicit
Type Riding
    Region As Integer
    RidingName As String * 50
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
    PrairieRegions = Int(InputBox("Number of Divsions in the Prairies"))
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
    Open "LastElection.dat" For Random As #1 Len = 27040
    Get #1, 1, P
    Close #1
End Sub
Sub InitializeRidings()
    Dim RPCode%
    Dim TRidingName$
    Dim NewRiding As Boolean
    Dim PartyCount%
    Dim RidingCount% '
    RidingCount = -1
    Dim DiscardFields%
    Dim Discard$
    Dim TVotes&
    Open "TRANSP~1.CSV" For Input As #1
    Do
        Input #1, RPCode
        Input #1,Discard$
        Input #1, TRidingName
        If RidingCount = -1 Then
            NewRiding = True
        ElseIf Left$(TRidingName, 50) <> Trim(P.Ridings(RidingCount).RidingName) Then
            NewRiding = True
        End If
        If NewRiding Then
            RidingCount% = RidingCount% + 1
            P.Ridings(RidingCount).Region = RPCode%
            P.Ridings(RidingCount).RidingName = TRidingName
        End If
        NewRiding = False
        For DiscardFields = 4 To 13
            Input #1, Discard
        Next
        For PartyCount = 0 To 5
            Input #1, TVotes
            P.Ridings(RidingCount%).PartyVotes(PartyCount) = P.Ridings(RidingCount).PartyVotes(PartyCount) + TVotes
            P.Ridings(RidingCount%).TotalVotes = P.Ridings(RidingCount).TotalVotes + TVotes
        Next
    Loop Until EOF(1)
    Close #1
    Open "LastElection.dat" For Random As #1 Len = 27040
    Put #1, 1, P
    Close #1
End Sub
