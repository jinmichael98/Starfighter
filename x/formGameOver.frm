VERSION 5.00
Begin VB.Form formGameOver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Game Over - Starfighter"
   ClientHeight    =   7200
   ClientLeft      =   4560
   ClientTop       =   8775
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   7200
   ScaleWidth      =   12000
   Begin VB.PictureBox picBg 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   14400
      Left            =   0
      Picture         =   "formGameOver.frx":0000
      ScaleHeight     =   14400
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   -7200
      Width           =   12000
      Begin VB.Timer tmr100 
         Interval        =   100
         Left            =   11400
         Top             =   9999
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press any key to skip and show stats..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   0
         TabIndex        =   3
         Top             =   13320
         Visible         =   0   'False
         Width           =   12015
      End
      Begin VB.Label lblResult 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "G A M E O V E R"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   0
         TabIndex        =   2
         Top             =   7560
         Width           =   12015
      End
      Begin VB.Label lblOutcome 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   10200
         Width           =   12015
      End
      Begin VB.Label lblStats 
         BackColor       =   &H00000000&
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   1695
         Index           =   0
         Left            =   1005
         TabIndex        =   4
         Top             =   9960
         Visible         =   0   'False
         Width           =   4995
      End
      Begin VB.Label lblStats 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         Height          =   1695
         Index           =   1
         Left            =   6000
         TabIndex        =   5
         Top             =   9960
         Visible         =   0   'False
         Width           =   4995
      End
   End
End
Attribute VB_Name = "formGameOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strOutcomes(4, 4) As String

Dim intPhase As Integer
Dim intWaitingTime As Integer
Dim intCurrentStr As Integer
Dim intLetter As Integer


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case intPhase
        Case 0
            lblPrompt.Caption = "Press any key to go to main menu..."
            lblOutcome.Visible = False
            tmr100.Enabled = False
            intPhase = 1
            
            For i = 0 To intPlayerMode - 1
                With pShip(i)
                    lblStats(i).Visible = True
                    lblStats(i).Caption = "Player " & i + 1 & vbNewLine & vbNewLine _
                    & "Enemies Destroyed: " & .Score & vbNewLine _
                    & "Damage Taken: " & ((1000 - .Health) / 1000) * 100 & "%"
                    
                End With
                
            Next i
            
        
        Case 1
            Call Shell(App.Path & "\" & App.EXEName & ".exe", 1)
            End
            
        End Select
    

End Sub

Private Sub Form_Load()

    If intOutcome > 0 Then
        lblResult.Caption = "V I C T O R Y"
        
    End If

    
    'LOSE
    strOutcomes(0, 0) = "The enemy forces "
    strOutcomes(0, 1) = "have breached into "
    strOutcomes(0, 2) = "Earth's atmosphere. "
    strOutcomes(0, 3) = "Who knows what "
    strOutcomes(0, 4) = "might happen next... "
    
    'WIN EZ
    strOutcomes(1, 0) = "Captain, "
    strOutcomes(1, 1) = "the enemies have "
    strOutcomes(1, 2) = "halted their advance. "
    strOutcomes(1, 3) = "It looks like the battles have stopped "
    strOutcomes(1, 4) = "        for now... "
    
    'WIN NORMAL
    strOutcomes(2, 0) = "Great job, "
    strOutcomes(2, 1) = "Captain! "
    strOutcomes(2, 2) = "Thanks to your efforts, "
    strOutcomes(2, 3) = "humanity is spared "
    strOutcomes(2, 4) = "     for another day... "
    
    'WIN HARD
    strOutcomes(3, 0) = "Unbelievable! "
    strOutcomes(3, 1) = "Your resistance "
    strOutcomes(3, 2) = "has severely crippled their advance. "
    strOutcomes(3, 3) = "You, my captain, "
    strOutcomes(3, 4) = "shall land home a hero... "
    
    'WIN IMPOSSIBLE
    strOutcomes(4, 0) = "Impossible! "
    strOutcomes(4, 1) = "Your resistance has "
    strOutcomes(4, 2) = "single-handedly changed "
    strOutcomes(4, 3) = "the course of the war. "
    strOutcomes(4, 4) = "     Now's our turn... "
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    intPhase = 0
    
    If intPlayerMode = 1 Then
        lblStats(0).Left = Me.ScaleWidth / 2 - 2500
    
    End If
    
End Sub

Private Sub tmr100_Timer()

    If intPhase = 0 Then
        If lblOutcome.Caption <> strOutcomes(intOutcome, intCurrentStr) Then 'TYPE IN
            lblOutcome.Caption = Left(strOutcomes(intOutcome, intCurrentStr), intLetter)
            intLetter = intLetter + 1
                
        ElseIf intCurrentStr < 4 Then 'UBound(strOutcomes, formGame.intOutcome) Then
            lblOutcome.Caption = ""
            intCurrentStr = intCurrentStr + 1
            intLetter = 0
                    
        End If

        If intWaitingTime >= 10 Then
            lblPrompt.Visible = True
            Me.KeyPreview = True
            'MsgBox 1
        
        Else: intWaitingTime = intWaitingTime + 1
        
        End If
        
    End If

End Sub
