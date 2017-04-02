VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Regional Seat Count Projector"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check2 
      Caption         =   "&Riding by Riding(not to be taken seriously)"
      Height          =   735
      Left            =   120
      TabIndex        =   51
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   34
      Left            =   0
      TabIndex        =   46
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   33
      Left            =   0
      TabIndex        =   45
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   32
      Left            =   0
      TabIndex        =   44
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   31
      Left            =   0
      TabIndex        =   43
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   30
      Left            =   0
      TabIndex        =   42
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   29
      Left            =   0
      TabIndex        =   41
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   28
      Left            =   0
      TabIndex        =   40
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   27
      Left            =   0
      TabIndex        =   39
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   26
      Left            =   0
      TabIndex        =   38
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   25
      Left            =   0
      TabIndex        =   37
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   24
      Left            =   0
      TabIndex        =   36
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   3720
      TabIndex        =   34
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Compute Regional Totals"
      Height          =   615
      Left            =   6480
      TabIndex        =   33
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "&Proportional Vote Swing"
      Height          =   495
      Left            =   1680
      TabIndex        =   32
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   23
      Left            =   0
      TabIndex        =   31
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   22
      Left            =   0
      TabIndex        =   30
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   21
      Left            =   0
      TabIndex        =   29
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   20
      Left            =   0
      TabIndex        =   28
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   19
      Left            =   0
      TabIndex        =   27
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   18
      Left            =   0
      TabIndex        =   26
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   17
      Left            =   0
      TabIndex        =   25
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   16
      Left            =   0
      TabIndex        =   24
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   15
      Left            =   0
      TabIndex        =   23
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   14
      Left            =   0
      TabIndex        =   22
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   13
      Left            =   0
      TabIndex        =   21
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   12
      Left            =   0
      TabIndex        =   20
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   11
      Left            =   0
      TabIndex        =   19
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   10
      Left            =   0
      TabIndex        =   18
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   9
      Left            =   0
      TabIndex        =   17
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   8
      Left            =   0
      TabIndex        =   16
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   7
      Left            =   0
      TabIndex        =   15
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   6
      Left            =   0
      TabIndex        =   14
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   5
      Left            =   0
      TabIndex        =   13
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   12
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   11
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   10
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.TextBox txtVoteCount 
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Text            =   "txtVoteCount"
      Top             =   0
      Width           =   735
   End
   Begin VB.Label lblRegion 
      Caption         =   "lblRegion"
      Height          =   495
      Index           =   6
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblRegion 
      Caption         =   "lblRegion"
      Height          =   495
      Index           =   5
      Left            =   0
      TabIndex        =   49
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblRegion 
      Caption         =   "lblRegion"
      Height          =   495
      Index           =   4
      Left            =   0
      TabIndex        =   48
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblPartyName 
      Caption         =   "lblPartyName"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   47
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "Title:"
      Height          =   255
      Left            =   3000
      TabIndex        =   35
      Top             =   3240
      Width           =   375
   End
   Begin VB.Label lblRegion 
      Caption         =   "lblRegion"
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblRegion 
      Caption         =   "lblRegion"
      Height          =   495
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblRegion 
      Caption         =   "lblRegion"
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblRegion 
      Caption         =   "lblRegion"
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label lblPartyName 
      Caption         =   "lblPartyName"
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label lblPartyName 
      Caption         =   "lblPartyName"
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblPartyName 
      Caption         =   "lblPartyName"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblPartyName 
      Caption         =   "lblPartyName"
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   0
      Top             =   1560
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Dim TRiding As Riding
    Dim PartyCount%
    Dim HighVote&
    Dim HighVoteParty%
    Dim CurrVote&
    Dim RidingCount%
    Dim RegionCount%
    Dim SeatTotals%(1 To 7, 5)
    Dim VoteAdjustment!(1 To 7, 5)
    Dim ProportionalSwing As Boolean
    Dim TextBoxNo%
    Dim CurrRegion%
    
    ProportionalSwing = Check1
    For RegionCount% = 1 To LastRegion
        For PartyCount = 0 To 4
            TextBoxNo = (RegionCount - 1) * 5 + PartyCount
            If ProportionalSwing Then
                If PartyCount <> 3 Or RegionCount = 2 Then
                    VoteAdjustment(RegionCount, PartyCount) = Val(txtVoteCount(TextBoxNo)) / (Regions(RegionCount).PartyVotes(PartyCount) / Regions(RegionCount).TotalVotes * 100)
                Else
                    VoteAdjustment(RegionCount, PartyCount) = 1
                End If
            Else
                VoteAdjustment(RegionCount, PartyCount) = Val(txtVoteCount(TextBoxNo)) / 100 - Regions(RegionCount).PartyVotes(PartyCount) / Regions(RegionCount).TotalVotes
            End If
        Next
    Next
    If Check2 Then Open Text2 + ".csv" For Output As #3

    For RidingCount = 0 To 334
        HighVote = 0
        HighVoteParty = 0
        TRiding = P.Ridings(RidingCount)
        CurrRegion = ReturnRegion(TRiding)
        For PartyCount = 0 To 4
            If ProportionalSwing Then
                CurrVote = TRiding.PartyVotes(PartyCount) * VoteAdjustment(CurrRegion, PartyCount)
            Else
                CurrVote = TRiding.PartyVotes(PartyCount) + VoteAdjustment(CurrRegion, PartyCount) * TRiding.TotalVotes
            End If
            If CurrVote > HighVote Then HighVoteParty = PartyCount: HighVote = CurrVote
        Next
        SeatTotals(CurrRegion, HighVoteParty) = SeatTotals(CurrRegion, HighVoteParty) + 1
        If Check2 Then
            Print #3, TRiding.RidingName; ","; PartyNames$(HighVoteParty)
        End If
    Next
    If Check2 Then Close #3
    Open "projections.htm" For Append As #2
    Print #2, "<Center>Seat Projections:"; Text2; "<br>"
    Print #2, Date$; " - "; Time$; "<br></center>"
    Print #2, "<Table width=100%><td>"
    For RegionCount = 1 To LastRegion
        Print #2, "<td>"; Trim(Regions(RegionCount).RegionName);
    Next
    For PartyCount = 0 To 4
        Print #2, "<tr>"
        Print #2, "<td>"; PartyNames(PartyCount);
        For RegionCount = 1 To LastRegion
            Print #2, "<td>"; SeatTotals(RegionCount, PartyCount);
        Next
    Next
    Print #2, "<tr></table>"
    Close #2
End Sub

Private Sub Form_Load()
    InitializeRidingsAndRegions
    Dim PartyCount%
    Dim RegionCount%
    Dim TextBoxNo%
    For PartyCount = 0 To 4
        lblPartyName(PartyCount).Caption = PartyNames(PartyCount)
        lblPartyName(PartyCount).Top = 500 + PartyCount * 400
    Next
    
    For RegionCount = 1 To LastRegion
        lblRegion(RegionCount - 1).Left = 1000 + 1000 * RegionCount
        lblRegion(RegionCount - 1).Caption = Regions(RegionCount).RegionName
        For PartyCount = 0 To 4
            TextBoxNo = (RegionCount - 1) * 5 + PartyCount
            txtVoteCount(TextBoxNo).Top = 500 + 400 * PartyCount
            txtVoteCount(TextBoxNo).Left = 1000 + 1000 * RegionCount
            txtVoteCount(TextBoxNo).Text = Int(Regions(RegionCount).PartyVotes(PartyCount) / Regions(RegionCount).TotalVotes * 1000) / 10
            If PartyCount = 3 And RegionCount <> 2 Then txtVoteCount(TextBoxNo).Visible = False
        Next
    Next
    For TextBoxNo = TextBoxNo + 1 To 34
        txtVoteCount(TextBoxNo).Visible = False
    Next
    For RegionCount = LastRegion To 6
        lblRegion(RegionCount).Visible = False
    Next
End Sub
