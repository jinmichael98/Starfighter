VERSION 5.00
Begin VB.Form formMainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starfighter"
   ClientHeight    =   7200
   ClientLeft      =   14445
   ClientTop       =   2070
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
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
      Picture         =   "formMainMenu.frx":0000
      ScaleHeight     =   14400
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   -7200
      Width           =   12000
      Begin VB.ListBox lstCannons 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1140
         ItemData        =   "formMainMenu.frx":1F41
         Left            =   5000
         List            =   "formMainMenu.frx":1F51
         TabIndex        =   6
         Top             =   10200
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Timer tmr100 
         Interval        =   100
         Left            =   120
         Top             =   12360
      End
      Begin VB.Timer tmrVar 
         Left            =   120
         Top             =   12840
      End
      Begin VB.ListBox lstMenu 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1140
         ItemData        =   "formMainMenu.frx":1FA1
         Left            =   5000
         List            =   "formMainMenu.frx":1FB1
         TabIndex        =   4
         Top             =   10200
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblBriefing 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   3855
         Left            =   1005
         TabIndex        =   7
         Top             =   9120
         Visible         =   0   'False
         Width           =   10005
      End
      Begin VB.Image imgCannon 
         BorderStyle     =   1  'Fixed Single
         Height          =   2895
         Index           =   3
         Left            =   1440
         Picture         =   "formMainMenu.frx":1FFF
         Stretch         =   -1  'True
         Top             =   9360
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.Image imgCannon 
         BorderStyle     =   1  'Fixed Single
         Height          =   2895
         Index           =   2
         Left            =   1440
         Picture         =   "formMainMenu.frx":11A0D
         Stretch         =   -1  'True
         Top             =   9360
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.Image imgCannon 
         BorderStyle     =   1  'Fixed Single
         Height          =   2895
         Index           =   1
         Left            =   1440
         Picture         =   "formMainMenu.frx":34B16
         Stretch         =   -1  'True
         Top             =   9360
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.Image imgCannon 
         BorderStyle     =   1  'Fixed Single
         Height          =   2895
         Index           =   0
         Left            =   1440
         Picture         =   "formMainMenu.frx":5FA0D
         Stretch         =   -1  'True
         Top             =   9360
         Visible         =   0   'False
         Width           =   3450
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DESCRIPTION:                  AVERAGE DAMAGE WITH A HIGH RATE OF FIRE"""
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   7080
         TabIndex        =   5
         Top             =   9600
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "S T A R F I G H T E R"
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
         TabIndex        =   3
         Top             =   7560
         Visible         =   0   'False
         Width           =   12015
      End
      Begin VB.Label lblIntro 
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
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   10200
         Width           =   12000
      End
      Begin VB.Label lblPrompt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Press any key to skip..."
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
         TabIndex        =   1
         Top             =   13200
         Width           =   12015
      End
      Begin VB.Image imgCockpit 
         Height          =   7995
         Left            =   0
         Picture         =   "formMainMenu.frx":87532
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   12000
      End
   End
End
Attribute VB_Name = "formMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intPhase As Integer
Dim intCurrentStr As Integer
Dim intLetter As Integer
Dim intListSelect As Integer


Dim boolStart As Boolean, boolUpward As Boolean

Dim strLevels(0 To 3) As String
Dim strIntro(0 To 7) As String
Dim strDesc(0 To 3) As String


Private Function fnSearchList(List As ListBox) As Integer

    For i = 0 To List.ListCount - 1
        If List.Selected(i) = True Then
            fnSearchList = i
            Exit For
        
        End If
        
    Next i
                    
End Function

Private Sub subUpdateList(Lst As ListBox, Arr As Variant)

    For i = 0 To UBound(Arr)
        Lst.List(i) = Arr(i)
        
    Next i

End Sub

Private Sub subLoadGame()

    Unload Me
    'Set formMainMenu = Nothing
    
    Load formGame
    formGame.Show

End Sub

Private Sub subUpdtP2() 'UPDATE DECRIPTION AND IMAGE FOR PHASE 2

    intListSelect = fnSearchList(lstCannons)
    
    For i = 0 To lstCannons.ListCount - 1
        If i = intListSelect Then
            imgCannon(i).Visible = True
            
            lblDesc.Caption = strDesc(i)
        
        Else
            imgCannon(i).Visible = False
            
        End If
    
    Next i

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case intPhase
        Case 0 'PHASE 0: INTRO
            lblIntro.Visible = False
            lblPrompt.Visible = False
            lblTitle.Visible = True
            lstMenu.Visible = True
            
            intPhase = 1
            
        Case 1 'PHASE 1: MAIN MENU
            If KeyCode = vbKeyUp Or KeyCode = vbKeyW Then 'CHANGE SELECTED ITEM
                intListSelect = fnSearchList(lstMenu)
                
                If intListSelect > 0 Then
                    lstMenu.Selected(intListSelect - 1) = True
                
                End If
                    
            ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyS Then 'CHANGE SELECTED ITEM
                intListSelect = fnSearchList(lstMenu)
                
                If intListSelect = 0 Then
                    lstMenu.Selected(1) = True
                            
                ElseIf intListSelect <> 0 And intListSelect < lstMenu.ListCount - 1 Then
                    lstMenu.Selected(intListSelect + 1) = True
                                
                End If
                
            ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then 'ENTER/SPACE KEY
                intListSelect = fnSearchList(lstMenu)
                Select Case intListSelect
                    Case Is <= 1 'SOLO OR COOP GAME
                        intPlayerMode = intListSelect + 1
                        intPhase = 2
                        
                        imgCannon(0).Visible = True
                        lstMenu.Visible = False
                        lstCannons.Visible = True
                        lblDesc.Visible = True
                        
                        If intPlayerMode = 1 Then
                            ReDim pShip(0) As Player
                            pShip(0).Health = 1000
                            pShip(0).Energy = 100
                            
                            lblPrompt.Caption = "Please equip a cannon for your ship."
                        
                        ElseIf intPlayerMode = 2 Then
                            ReDim pShip(1) As Player
                            pShip(0).Health = 1000
                            pShip(1).Health = 1000
                            pShip(0).Energy = 100
                            pShip(1).Energy = 100
                            
                            lblPrompt.Caption = "Player 1, please equip a cannon for your ship."
                            
                        
                        End If
                        
                        lblPrompt.Visible = True
                        
                        lstCannons.Selected(0) = True
                    
                    Case 2 'HELP
                        lstMenu.Visible = False
                        lblPrompt.Caption = "Press any key to come back to main menu..."
                        lblPrompt.Visible = True
                        lblBriefing.Visible = True
                        lblBriefing.Caption = "CONTROLS" & vbNewLine & vbNewLine _
                        & "PLAYER 1 KEYS:" & vbNewLine _
                        & "WASD KEYS FOR MOVEMENT, SPACE KEY TO FIRE" & vbNewLine & vbNewLine _
                        & "PLAYER 2 KEYS:" & vbNewLine _
                        & "ARROW KEYS FOR MOVEMENT, ENTER/RETURN KEY TO FIRE" & vbNewLine & vbNewLine _
                        & "HELPFUL TIPS" & vbNewLine & vbNewLine _
                        & "   -THE ENTIRE GAME CAN BE PLAYED ONLY USING A KEYBOARD" & vbNewLine _
                        & "   -DO NOT CRASH INTO YOUR PARTNER IN COOP!" & vbNewLine _
                        & "   -THE TRI-SHOT CANNON CAN ONE-SHOT ANY ENEMY WHEN ALL THE THREE MISSILES CONNECT" & vbNewLine _
                        & "   -YOU CAN DESTROY ENEMIES BEYOND THE SCREEN" & vbNewLine _
                        & "   -ENEMIES HAVE A MUCH LOWER RATE OF FIRE, TRY TO BAIT OUT THEIR SHOTS BY GETTING NEAR AND DODGING THE SHOT THEN COMING BACK IN"
                        
                        intPhase = 11

                    Case 3 'QUIT
                        End
                        
                End Select
                
            End If
                     
        Case 2, 3
            If KeyCode = vbKeyUp Or KeyCode = vbKeyW Then 'CHANGE SELECTED ITEM THEN UPDATE DESCRIPTION AND IMAGE
                intListSelect = fnSearchList(lstCannons)
                
                If intListSelect > 0 Then
                    lstCannons.Selected(intListSelect - 1) = True
                
                End If
                
                Call subUpdtP2
                    
            ElseIf KeyCode = vbKeyDown Or KeyCode = vbKeyS Then 'CHANGE SELECTED ITEM THEN UPDATE DESCRIPTION AND IMAGE
                intListSelect = fnSearchList(lstCannons)
                
                If intListSelect = 0 Then
                    lstCannons.Selected(1) = True
                            
                ElseIf intListSelect <> 0 And intListSelect < lstCannons.ListCount - 1 Then
                    lstCannons.Selected(intListSelect + 1) = True
                                
                End If
                
                Call subUpdtP2
                
            ElseIf KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then 'ENTER/SPACE KEY
                intListSelect = fnSearchList(lstCannons)
                
                With pShip(intPhase - 2)
                    .CannonType = intListSelect
                    Select Case .CannonType
                        Case 0 'RAPID CANNON
                            .OnFireCD = 3
                            .FirePower = 40
                            .FiringCost = 8
                            .MissileSpeed = 300
                            
                        Case 1 'TRI SHOTCANNON
                            .OnFireCD = 10
                            .FirePower = 55
                            .FiringCost = 14
                            .MissileSpeed = 500
                                
                        Case 2 'SNIPE CANNON
                            .OnFireCD = 30
                            .FirePower = 9000
                            .FiringCost = 60
                            .MissileSpeed = 1200
                                
                    End Select
                    
                End With
                
                If intPlayerMode = 1 Then 'GAME STARTING, 1PLAYER
                    'pShip(0).CannonType = intListSelect
                    If pShip(0).CannonType <> 3 Then
                        intPhase = 4
                        lblDesc.Visible = False
                        imgCannon(pShip(0).CannonType).Visible = False
                        lstCannons.Visible = False
                        lblPrompt.Visible = True
                        lblPrompt.Caption = "Please select a difficulty level."
                        
                        lstMenu.Visible = True
                        Call subUpdateList(lstMenu, strLevels)
                        
                    ElseIf pShip(0).CannonType = 3 Then
                        lblPrompt.Caption = "Don't be silly, please equip a REAL cannon for your ship."
                        
                    End If
                
                ElseIf intPlayerMode = 2 And intPhase = 2 And pShip(0).CannonType <> 3 Then 'SELECT CANNON FOR PLAYER2
                    'pShip(0).CannonType = intListSelect
                    If pShip(0).CannonType <> 3 Then
                        intPhase = 3
                        lblPrompt.Caption = "Player 2, please equip a cannon for your ship."
                        
                    ElseIf pShip(0).CannonType = 3 Then
                        lblPrompt.Caption = "Player 1, please equip a REAL cannon for your ship."
                    
                    End If
                    
                
                ElseIf intPlayerMode = 2 And intPhase = 3 Then 'GAME STARTING, 2PLAYERS
                    If pShip(1).CannonType <> 3 Then
                        'pShip(0).CannonType = intListSelect
                        intPhase = 4
                        lblDesc.Visible = False
                        imgCannon(pShip(1).CannonType).Visible = False
                        lstCannons.Visible = False
                        lblPrompt.Visible = True
                        lblPrompt.Caption = "Please select a difficulty level."
                        
                        lstMenu.Visible = True
                        Call subUpdateList(lstMenu, strLevels)
                    
                    ElseIf pShip(1).CannonType = 3 Then
                        lblPrompt.Caption = "Player 2, please equip a REAL cannon for your ship."
                    
                    End If
                    
                    
                End If
                    
                
                'Call subLoadGame
                
            End If
        
        Case 4
            
            Select Case KeyCode
                Case vbKeyW, vbKeyUp
                    intListSelect = fnSearchList(lstMenu)
                    
                    
                    If intListSelect > 0 Then
                        lstMenu.Selected(intListSelect - 1) = True
                        
                        
                    End If
                
                Case vbKeyS, vbKeyDown
                    intListSelect = fnSearchList(lstMenu)
                    
                    If intListSelect < 3 Then
                        lstMenu.Selected(intListSelect + 1) = True
                        
                    End If
                
                Case vbKeySpace, vbKeyReturn
                    intListSelect = fnSearchList(lstMenu)
                    intDifficulty = intListSelect
                    intGoal = 20 + (intDifficulty * 25)
                    
                    Select Case intDifficulty
                        Case 0
                            intShields = 10
            
                        Case 1
                            intShields = 5
            
                        Case 2
                            intShields = 3
            
        
                    End Select
                    
                    lblPrompt.Caption = "Press any key to start the game"
                    lblBriefing.Visible = True
                    lstMenu.Visible = False
                    
                    lblBriefing.Caption = "MISSION OBJECTIVES: " & vbNewLine & vbNewLine _
                    & "   -WEAKEN THE ENEMY FORCES BY TAKING DOWN AT LEAST " & intGoal & " ENEMY SHIPS" & vbNewLine _
                    & "   -DO NOT LET MORE THAN " & intShields & " ENEMY SHIPS PASS AND ADVANCE INTO EARTH"
                    
                    If intPlayerMode = 2 Then
                        lblBriefing.Caption = lblBriefing.Caption & vbNewLine & "   -PROTECT YOUR PARTNER"
                    End If
                    
                    intPhase = 99
                    
            
            End Select
            
        Case 11
            lstMenu.Visible = True
            lblPrompt.Visible = False
            lblBriefing.Visible = False
            
            intPhase = 1
            
        Case 99
            Call subLoadGame
    
    End Select

End Sub

Private Sub Form_Load()

    'INTRO TEXT STRINGS
    tmrVar.Interval = fnRndInt(500, 1000)
    strIntro(0) = "The intergalactic wars"
    strIntro(1) = "are reaching unthinkable"
    strIntro(2) = "levels of chaos."
    strIntro(3) = "You are one of the last"
    strIntro(4) = "hopes we have left."
    strIntro(5) = "Defend Earth against these"
    strIntro(6) = "swarms of intruders."
    strIntro(7) = "Fight for humanity!"
    
    'CANNON DESCRIPTION TEXT STRINGS
    strDesc(0) = "DESCRIPTION:" & vbNewLine & "LOW DAMAGE WITH A VERY HIGH RATE OF FIRE, CAN DRAIN OUT OF ENERGY QUICKLY IF USED INEFFICIENTLY"
    strDesc(1) = "DESCRIPTION:" & vbNewLine & "FIRES THREE BOLTS PER SHOT THAT ARE SPREAD IN A 45° ANGLE. HAS AN AVERAGE RATE OF FIRE AND DEALS AVERAGE DAMAGE PER BOLT, GREAT AT CLOSE RANGE AND CLEARING ENEMY SWARMS, ALTHOUGH INEFFICIENT USE CAN CAUSE ENERGY PROBLEMS"
    strDesc(2) = "DESCRIPTION:" & vbNewLine & "WITH AN EXTREMELY FAST MISSILE SPEED, THIS CANNON'S BOLTS WILL PIERCE THROUGH EVERYTHING THAT IT HITS. HAS A DOWNSIDE OF A VERY SLOW RATE OF FIRE "
    strDesc(3) = "DESCRIPTION:" & vbNewLine & "THE ULTIMATE WEAPON OF COURAGE AND STUPIDITY"
    
    'DIFFICULY DESCRIPTION TEXT STRINGS
    strLevels(0) = "EASY"
    strLevels(1) = "NORMAL"
    strLevels(2) = "HARD"
    strLevels(3) = "IMPOSSIBLE"
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
        
    lstMenu.Selected(0) = True
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Set formMainMenu = Nothing
    
End Sub

Private Sub lstMenu_Click()
    
    intListSelect = fnSearchList(lstMenu)
    
End Sub

Private Sub lstCannons_Click()

    If intPhase = 2 Then 'CHANGE NOTHING
        intListSelect = fnSearchList(lstCannons)
        Call subUpdtP2
    End If
    

End Sub

Private Sub tmr100_Timer() 'TIMER FOR INTRO PHASE

    If intPhase = 0 Then
        If lblIntro.Caption <> strIntro(intCurrentStr) Then 'TYPE IN INTRO
            lblIntro.Caption = Left(strIntro(intCurrentStr), intLetter)
            intLetter = intLetter + 1
                
        ElseIf intCurrentStr < UBound(strIntro) Then
            lblIntro.Caption = ""
            intCurrentStr = intCurrentStr + 1
            intLetter = 0
                    
        End If
           
        If boolUpward Then
            imgCockpit.Top = imgCockpit.Top - 10
            imgCockpit.Left = imgCockpit.Left + 1
            'lblPrompt.Top = lblPrompt.Top + 1
        
        Else
            imgCockpit.Top = imgCockpit.Top + 10
            imgCockpit.Left = imgCockpit.Left - 1
            'picBg.Top = picBg.Top + 1
            'lblPrompt.Top = lblPrompt.Top - 1
            
        End If

    End If

End Sub

Private Sub tmrVar_Timer()
    
    If boolUpward Then
        tmrVar.Interval = fnRndInt(500, 2000)
        boolUpward = False
        
    Else
        boolUpward = True
        
    End If
    
End Sub
