VERSION 5.00
Begin VB.Form formGame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Starfighter"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   Begin VB.PictureBox picBg 
      BorderStyle     =   0  'None
      Height          =   18000
      Left            =   0
      MousePointer    =   2  'Cross
      Picture         =   "formMain.frx":0000
      ScaleHeight     =   18000
      ScaleWidth      =   12000
      TabIndex        =   0
      Top             =   -9000
      Width           =   12000
      Begin VB.Timer tmrVar 
         Interval        =   5000
         Left            =   11160
         Top             =   17280
      End
      Begin VB.Timer tmr10 
         Interval        =   10
         Left            =   11160
         Top             =   16800
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   9
         Left            =   4560
         Picture         =   "formMain.frx":24A6
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   8
         Left            =   4080
         Picture         =   "formMain.frx":26CF
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   7
         Left            =   3600
         Picture         =   "formMain.frx":28F8
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   6
         Left            =   3120
         Picture         =   "formMain.frx":2B21
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   5
         Left            =   2640
         Picture         =   "formMain.frx":2D4A
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   4
         Left            =   2160
         Picture         =   "formMain.frx":2F73
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   3
         Left            =   1680
         Picture         =   "formMain.frx":319C
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   2
         Left            =   1200
         Picture         =   "formMain.frx":33C5
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   1
         Left            =   720
         Picture         =   "formMain.frx":35EE
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgShield 
         Height          =   450
         Index           =   0
         Left            =   240
         Picture         =   "formMain.frx":3817
         Top             =   9240
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.Image imgEnemyBolt 
         Height          =   555
         Index           =   0
         Left            =   20000
         Picture         =   "formMain.frx":3A40
         Top             =   11880
         Width           =   135
      End
      Begin VB.Image imgHPBar 
         Height          =   420
         Index           =   1
         Left            =   8280
         Picture         =   "formMain.frx":3AEE
         Stretch         =   -1  'True
         Top             =   17400
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Image imgEnergyBar 
         Height          =   210
         Index           =   1
         Left            =   8280
         Picture         =   "formMain.frx":7A7C
         Stretch         =   -1  'True
         Top             =   17760
         Visible         =   0   'False
         Width           =   3600
      End
      Begin VB.Image imgPlayer2BoltL 
         Height          =   495
         Index           =   0
         Left            =   360
         Picture         =   "formMain.frx":B490
         Stretch         =   -1  'True
         Top             =   15360
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgPlayer2BoltR 
         Height          =   495
         Index           =   0
         Left            =   1560
         Picture         =   "formMain.frx":BA5F
         Stretch         =   -1  'True
         Top             =   15360
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgPlayer2Bolt 
         Height          =   495
         Index           =   0
         Left            =   1080
         Picture         =   "formMain.frx":C036
         Stretch         =   -1  'True
         Top             =   15360
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgPlayer 
         Height          =   705
         Index           =   1
         Left            =   6120
         Picture         =   "formMain.frx":C346
         Stretch         =   -1  'True
         Top             =   16440
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image imgEnemyHPBar 
         Height          =   135
         Index           =   0
         Left            =   6840
         Picture         =   "formMain.frx":CD62
         Stretch         =   -1  'True
         Top             =   13800
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Image imgPlayer1BoltR 
         Height          =   495
         Index           =   0
         Left            =   1440
         Picture         =   "formMain.frx":12874
         Stretch         =   -1  'True
         Top             =   16560
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgPlayer1BoltL 
         Height          =   495
         Index           =   0
         Left            =   480
         Picture         =   "formMain.frx":12E1B
         Stretch         =   -1  'True
         Top             =   16560
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   8
         Left            =   8280
         Picture         =   "formMain.frx":133C3
         Stretch         =   -1  'True
         Top             =   9960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   7
         Left            =   7320
         Picture         =   "formMain.frx":2AC09
         Stretch         =   -1  'True
         Top             =   9960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   6
         Left            =   6360
         Picture         =   "formMain.frx":3C127
         Stretch         =   -1  'True
         Top             =   9960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   5
         Left            =   5520
         Picture         =   "formMain.frx":4EC31
         Stretch         =   -1  'True
         Top             =   10005
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   4
         Left            =   4440
         Picture         =   "formMain.frx":63433
         Stretch         =   -1  'True
         Top             =   9960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   3
         Left            =   3360
         Picture         =   "formMain.frx":7A8F9
         Stretch         =   -1  'True
         Top             =   9960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   2
         Left            =   2280
         Picture         =   "formMain.frx":8DF2C
         Stretch         =   -1  'True
         Top             =   9960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   1
         Left            =   1440
         Picture         =   "formMain.frx":A2265
         Stretch         =   -1  'True
         Top             =   10005
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgExplosion 
         Height          =   900
         Index           =   0
         Left            =   600
         Picture         =   "formMain.frx":B8A75
         Stretch         =   -1  'True
         Top             =   9960
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Image imgEnemyType 
         Height          =   1200
         Index           =   2
         Left            =   9600
         Picture         =   "formMain.frx":D02B5
         Stretch         =   -1  'True
         Top             =   14040
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.Image imgEnemyType 
         Height          =   915
         Index           =   1
         Left            =   8280
         Picture         =   "formMain.frx":D0CEC
         Stretch         =   -1  'True
         Top             =   14040
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image imgEnemyType 
         Height          =   750
         Index           =   0
         Left            =   7080
         Picture         =   "formMain.frx":D1842
         Stretch         =   -1  'True
         Top             =   14040
         Visible         =   0   'False
         Width           =   650
      End
      Begin VB.Image imgEnemy 
         Height          =   945
         Index           =   0
         Left            =   3120
         Picture         =   "formMain.frx":D22FB
         Stretch         =   -1  'True
         Top             =   8055
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image imgPlayer 
         Height          =   705
         Index           =   0
         Left            =   5040
         Picture         =   "formMain.frx":D2DB4
         Stretch         =   -1  'True
         Top             =   16440
         Width           =   930
      End
      Begin VB.Image imgPlayerMove 
         Height          =   705
         Index           =   0
         Left            =   480
         Picture         =   "formMain.frx":D37CD
         Stretch         =   -1  'True
         Top             =   13440
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image imgPlayerMove 
         Height          =   705
         Index           =   1
         Left            =   1440
         Picture         =   "formMain.frx":D423B
         Stretch         =   -1  'True
         Top             =   13440
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image imgPlayerMove 
         Height          =   705
         Index           =   2
         Left            =   2400
         Picture         =   "formMain.frx":D4C54
         Stretch         =   -1  'True
         Top             =   13440
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Image imgPlayer1Bolt 
         Height          =   495
         Index           =   0
         Left            =   1080
         Picture         =   "formMain.frx":D563C
         Stretch         =   -1  'True
         Top             =   16560
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Image imgHPBar 
         Height          =   420
         Index           =   0
         Left            =   120
         Picture         =   "formMain.frx":D5958
         Stretch         =   -1  'True
         Top             =   17400
         Width           =   3600
      End
      Begin VB.Image imgEnergyBar 
         Height          =   210
         Index           =   0
         Left            =   120
         Picture         =   "formMain.frx":D98E6
         Stretch         =   -1  'True
         Top             =   17760
         Width           =   3600
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "status"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   19.5
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
         Top             =   10800
         Width           =   12000
      End
   End
End
Attribute VB_Name = "formGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intBaseSpeed As Integer

Dim intTotal As Integer

Dim boolBoost As Boolean

Dim nmeShip() As Enemy

Private Function fnCheckCollision(objA As Image, objB As Image) As Boolean
    
    'CHECKS FOR COLLISION OF TOP SIDE OF A, THEN BOTTOM, THEN LEFT, AND RIGHT
    If ((objA.Top >= objB.Top And objA.Top <= objB.Top + objB.Height) Or _
    (objA.Top + objA.Height >= objB.Top And objA.Top + objA.Height <= objB.Top + objB.Height)) _
    And _
    ((objA.Left >= objB.Left And objA.Left <= objB.Left + objB.Width) Or _
    (objA.Left + objA.Width >= objB.Left And objA.Left + objA.Width <= objB.Left + objB.Width)) Then
        fnCheckCollision = True
        
    Else
        fnCheckCollision = False
        
    End If
    
End Function

Private Sub subGameOver(Win As Boolean)

    If Win = False Then
        intOutcome = 0
        
        
    Else
        intOutcome = 1 + intDifficulty
    End If
    
    tmr10.Enabled = False
    tmrVar.Enabled = False
        
    Me.Hide
            
    formGameOver.Show
    

End Sub

Private Sub subCenterPos(objA As Image, objB As Image, OnTop As Boolean)

    objA.Left = objB.Left + objB.Width / 2 - objA.Width / 2
    On Error GoTo 0
            
    If OnTop Then
        objA.Top = (objB.Top - objB.Height / 2) - objA.Height
    
    Else
        objA.Top = (objB.Top + objB.Height / 2) + objB.Height
    
    End If
    
End Sub

Private Sub subNextFrame(Img As Image, imgArray As Variant, CurrentIndex As Integer)

    If CurrentIndex < imgArray.UBound Then
        Img.Picture = imgArray(CurrentIndex + 1).Picture
        
    ElseIf CurrentIndex = imgArray.UBound Then
        Img.Visible = False
        
    End If

End Sub


Private Sub subLoadNew(imgArray As Variant, imgSample As Variant, SampleIndex As Integer)

    Load imgArray(imgArray.UBound + 1)
         
    With imgArray(imgArray.UBound)
        .Picture = imgSample(SampleIndex).Picture
        .Visible = True
        
    End With
    
End Sub

Private Sub subFire(Shooter As Image, Projectile As Image, i As Integer, IsPlayer As Boolean)
    
    'imgProjectile = imgPlayer1Bolt(imgPlayer1Bolt.UBound)
    If IsPlayer Then
        pShip(i).FireCD = pShip(i).OnFireCD
        pShip(i).Energy = pShip(i).Energy - pShip(i).FiringCost
        Call subCenterPos(Projectile, Shooter, True)
        
    Else
        nmeShip(i).FireCD = nmeShip(i).OnFireCD
        Call subCenterPos(Projectile, Shooter, False)
    
    End If

    
End Sub

Private Sub subReposition(YDiff As Integer)

    picBg.Top = picBg.Top + YDiff
    
    For i = 0 To UBound(pShip)
        imgPlayer(i).Top = imgPlayer(i).Top - YDiff
        imgEnergyBar(i).Top = imgEnergyBar(i).Top - YDiff
        imgHPBar(i).Top = imgHPBar(i).Top - YDiff
    
    Next i
    
    lblStatus.Top = lblStatus.Top - YDiff
    
    For i = 0 To intShields - 1
        imgShield(i).Top = imgShield(i).Top - YDiff
    Next i
    
    For i = 0 To imgPlayer1Bolt.UBound
        imgPlayer1Bolt(i).Top = imgPlayer1Bolt(i).Top - YDiff
        
    Next i
    
    For i = 0 To imgEnemy.UBound
        imgEnemy(i).Top = imgEnemy(i).Top - YDiff
        imgEnemyHPBar(i).Top = imgEnemyHPBar(i).Top - YDiff
    Next i
    
    For i = 0 To imgEnemyBolt.UBound
        imgEnemyBolt(i).Top = imgEnemyBolt(i).Top - YDiff
        On Error Resume Next
    
    Next i
          
End Sub

'PROGRAM STARTS
Private Sub Form_Load()
    
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    With picBg
        .Height = Me.ScaleHeight * 2
        .Width = Me.ScaleWidth
        .Top = -Me.ScaleHeight
    
    End With
    
    lblStatus.Caption = ""
            
    ReDim nmeShip(0) As Enemy
    
    nmeShip(0).Destroyed = True
    With nmeShip(0)
        Randomize
        .EnemyType = fnRndInt(0, 2)
        .Health = 50 + (50 * .EnemyType)
        .MaxHealth = .Health
        .Speed = 100 - (45 * .EnemyType)
        .EnemyIndex = UBound(nmeShip)
        
        Call subCenterPos(imgEnemyHPBar(0), imgEnemy(0), True)
    End With
    
    If intPlayerMode = 1 Then
        imgPlayer(0).Left = Me.ScaleWidth / 2 - imgPlayer(0).Width / 2
        
    ElseIf intPlayerMode = 2 Then
        imgPlayer(0).Left = Me.ScaleWidth / 4 - imgPlayer(0).Width / 2
        imgPlayer(1).Left = 3 * (Me.ScaleWidth / 4) - imgPlayer(0).Width / 2
        
        imgPlayer(1).Visible = True
        imgHPBar(1).Visible = True
        imgEnergyBar(1).Visible = True
    
        
    End If
    
    For i = 0 To intShields - 1
        imgShield(i).Visible = True
    Next i
    
    'Call fnCheckCollision(imgPlayer, imgEnemy(0))
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyW 'SPEED UP
            imgPlayer(0).Picture = imgPlayerMove(1).Picture
            If pShip(0).YSpeed <= 10 And pShip(0).Energy >= 25 Then
                pShip(0).YSpeed = 25
                pShip(0).Energy = pShip(0).Energy - 25
            
            End If
            
        Case vbKeyS 'SLOW DOWN
            imgPlayer(0).Picture = imgPlayerMove(1).Picture
            If pShip(0).YSpeed > 10 Then
                pShip(0).YSpeed = 10 'pShip(0).YSpeed - 5

            End If
            
        Case vbKeyA 'TURN LEFT
            imgPlayer(0).Picture = imgPlayerMove(0).Picture
            pShip(0).XSpeed = -intBaseSpeed * 6
        
        Case vbKeyD 'TURN RIGHT
            imgPlayer(0).Picture = imgPlayerMove(2).Picture
            pShip(0).XSpeed = intBaseSpeed * 6
            
        Case vbKeySpace 'FIRE
            If pShip(0).FireCD <= 0 Then
                If (pShip(0).Energy >= pShip(0).FiringCost And pShip(0).CannonType <> 1) Or pShip(0).Energy >= pShip(0).FiringCost * 3 Then
                    Call subLoadNew(imgPlayer1Bolt, imgPlayer1Bolt, imgPlayer1Bolt.LBound)
                    Call subFire(imgPlayer(0), imgPlayer1Bolt(imgPlayer1Bolt.UBound), 0, True)
                    
                    If pShip(0).CannonType = 1 Then
                        Call subLoadNew(imgPlayer1BoltL, imgPlayer1BoltL, 0)
                        Call subLoadNew(imgPlayer1BoltR, imgPlayer1BoltR, 0)
                        
                        Call subFire(imgPlayer(0), imgPlayer1BoltL(imgPlayer1BoltL.UBound), 0, True)
                        Call subFire(imgPlayer(0), imgPlayer1BoltR(imgPlayer1BoltR.UBound), 0, True)
                        
                        imgPlayer1BoltL(imgPlayer1BoltL.UBound).Left = imgPlayer1Bolt(imgPlayer1Bolt.UBound).Left - imgPlayer1BoltL(imgPlayer1BoltL.UBound).Width
                        imgPlayer1BoltR(imgPlayer1BoltR.UBound).Left = imgPlayer1Bolt(imgPlayer1Bolt.UBound).Left + imgPlayer1BoltR(imgPlayer1BoltR.UBound).Width
                        
                    End If
                    
                Else
                    lblStatus.Caption = "You need more energy to fire!"
                    
                    
                End If
                
            End If
            
        Case vbKeyEscape
            tmr10.Enabled = False
            tmr10.Enabled = False
        
    End Select
    
    If intPlayerMode = 2 Then
        
        Select Case KeyCode
        
            Case vbKeyUp 'SPEED UP
                'imgPlayer(0).Picture = imgPlayerMove(1).Picture
                If pShip(1).YSpeed <= 10 And pShip(1).Energy >= 25 Then
                    pShip(1).YSpeed = 25
                    pShip(1).Energy = pShip(1).Energy - 25
                
                End If
                
            Case vbKeyDown 'SLOW DOWN
                'imgPlayer(0).Picture = imgPlayerMove(1).Picture
                If pShip(1).YSpeed > 10 Then
                    pShip(1).YSpeed = 10 'pShip(1).YSpeed - 5
    
                End If
                
            Case vbKeyLeft 'TURN LEFT
                'imgPlayer(1).Picture = imgPlayerMove(0).Picture
                pShip(1).XSpeed = -pShip(1).YSpeed * 6
            
            Case vbKeyRight 'TURN RIGHT
                imgPlayer(0).Picture = imgPlayerMove(2).Picture
                pShip(1).XSpeed = pShip(1).YSpeed * 6
                
            Case vbKeyReturn 'FIRE
                If pShip(1).FireCD <= 0 Then
                    If (pShip(1).Energy >= pShip(1).FiringCost And pShip(1).CannonType <> 1) Or pShip(1).Energy >= pShip(1).FiringCost * 3 Then
                        Call subLoadNew(imgPlayer2Bolt, imgPlayer2Bolt, imgPlayer2Bolt.LBound)
                        Call subFire(imgPlayer(1), imgPlayer2Bolt(imgPlayer2Bolt.UBound), 1, True)
                        
                        If pShip(1).CannonType = 1 Then
                            Call subLoadNew(imgPlayer2BoltL, imgPlayer2BoltL, 0)
                            Call subLoadNew(imgPlayer2BoltR, imgPlayer2BoltR, 0)
                            
                            Call subFire(imgPlayer(1), imgPlayer2BoltL(imgPlayer2BoltL.UBound), 1, True)
                            Call subFire(imgPlayer(1), imgPlayer2BoltR(imgPlayer2BoltR.UBound), 1, True)
                            
                            imgPlayer2BoltL(imgPlayer2BoltL.UBound).Left = imgPlayer2Bolt(imgPlayer2Bolt.UBound).Left - imgPlayer2BoltL(imgPlayer2BoltL.UBound).Width
                            imgPlayer2BoltR(imgPlayer2BoltR.UBound).Left = imgPlayer2Bolt(imgPlayer2Bolt.UBound).Left + imgPlayer2BoltR(imgPlayer2BoltR.UBound).Width
                            
                        End If
                        
                    Else
                        lblStatus.Caption = "You need more energy to fire!"
                        
                        
                    End If
                    
                End If
                
            Case vbKeyDelete
                tmr10.Enabled = False
                tmr10.Enabled = False
            
        End Select
        
    End If


End Sub


Private Sub tmr10_Timer() 'ticks every 0.01 of a second

    If intPlayerMode = 2 Then
        If pShip(0).YSpeed <= pShip(1).YSpeed Then
            intBaseSpeed = pShip(0).YSpeed
            
        Else
            intBaseSpeed = pShip(1).YSpeed
        
        End If
        
        If (intBaseSpeed = 0 And pShip(0).YSpeed > 0) Or (intBaseSpeed = 0 And pShip(1).YSpeed > 0) Then
            lblStatus.Caption = "Waiting for other player to accelerate"
            
        Else: lblStatus.Caption = ""
        
        End If
    
    Else: intBaseSpeed = pShip(0).YSpeed
    
    End If

    For i = 0 To UBound(pShip)
        imgPlayer(i).Top = imgPlayer(i).Top - intBaseSpeed
        imgPlayer(i).Left = imgPlayer(i).Left + pShip(i).XSpeed
        
        If imgPlayer(i).Left < 0 Then 'PLAYER MOVEMENT LIMITED TO SCREEN
            imgPlayer(i).Left = 0
        
        ElseIf imgPlayer(i).Left + imgPlayer(i).Width > Me.ScaleWidth Then
            imgPlayer(i).Left = Me.ScaleWidth - imgPlayer(i).Width
    
        End If
        
        If fnCheckCollision(imgPlayer(i), imgPlayer(0)) And fnCheckCollision(imgPlayer(i), imgPlayer(1)) And intPlayerMode = 2 Then 'CHECK PLAYER-PLAYER COLLISION
            pShip(0).Health = 0
            pShip(1).Health = 0
            Call subGameOver(False)
        
        End If
        
        For n = 0 To imgEnemyBolt.UBound 'CHECK PLAYER-BOLT COLLISION
            If fnCheckCollision(imgEnemyBolt(n), imgPlayer(i)) = True Then
                pShip(i).Health = pShip(i).Health - 25
                imgEnemyBolt(n).Left = -5000
                On Error Resume Next
                
            End If
            
        Next n
        
        
        imgEnergyBar(i).Top = imgEnergyBar(i).Top - intBaseSpeed
        imgHPBar(i).Top = imgHPBar(i).Top - intBaseSpeed
        
        
        If pShip(i).FireCD > 0 Then 'UPDATE FIRING COOLDOWNS
            pShip(i).FireCD = pShip(i).FireCD - 1
        
        End If
        
        If pShip(i).Energy < 100 Then 'REGEN AND UPDATE ENERGY BAR
            pShip(i).Energy = pShip(i).Energy + 1
            imgEnergyBar(i).Width = 3600 * (pShip(i).Energy / 100)
        
            If imgEnergyBar(i).Width >= 300 And imgEnergyBar(i).Width <= 1000 Then
                lblStatus.Caption = "Low energy!"
            
            ElseIf imgEnergyBar(i).Width > 1500 Then
                lblStatus.Caption = ""
            
            End If
                
        End If
            
        If pShip(0).Health < 0 Then
            Call subGameOver(False)
            
        Else
            imgHPBar(0).Width = 3600 * (pShip(0).Health / 1000) 'UPDATE HEALTH BAR
            
        End If
            
        If intPlayerMode = 2 Then
            If pShip(1).Health < 0 Then
                Call subGameOver(False)
                
            Else
                imgHPBar(1).Width = 3600 * (pShip(1).Health / 1000)
            
            End If
            
        
            
        End If
        
    Next i
    
    lblStatus.Top = lblStatus.Top - intBaseSpeed 'pShip(0).YSpeed
    
    For i = 0 To intShields - 1
        imgShield(i).Top = imgShield(i).Top - intBaseSpeed
        
    Next i
    
    picBg.Top = picBg.Top + intBaseSpeed 'pShip(0).YSpeed
    
    
    If picBg.Top >= 0 Then 'REPOSITION WHEN BACKGROUND RESETS
      Call subReposition(-Me.ScaleHeight + picBg.Top)
        
    End If
    
    For i = 0 To UBound(nmeShip)
        With nmeShip(i)
            
            imgEnemy(i).Top = imgEnemy(i).Top + (.Speed - intBaseSpeed)
            imgEnemyHPBar(i).Top = imgEnemyHPBar(i).Top + (.Speed - intBaseSpeed)
            .FireCD = .FireCD - 1
                    
            If .Destroyed = False Then
                'CHECK IF ENEMY SHIP SHOULD FIRE AT PLAYER(S)
                If (Abs(imgEnemy(i).Left - imgPlayer(0).Left) <= 1000 Or Abs(imgEnemy(i).Left - imgPlayer(1).Left) <= 1000) And .FireCD <= 0 And imgPlayer(0).Top - imgEnemy(i).Top <= 6200 Then
                    Call subLoadNew(imgEnemyBolt, imgEnemyBolt, 0)
                    Call subFire(imgEnemy(i), imgEnemyBolt(imgEnemyBolt.UBound), Int(i), False)
                    
                End If
                
                If ((imgEnemy(i).Top > imgPlayer(0).Top + 100 And intPlayerMode = 1) Or (imgEnemy(i).Top > imgPlayer(0).Top + 100 And imgEnemy(i).Top > imgPlayer(1).Top + 100)) And .Passed = False Then 'DID ENEMY GET PAST PLAYER(S)??
                    intShields = intShields - 1
                    .Passed = True
                    If intShields < 0 Then
                        Call subGameOver(False)
                        
                    Else
                        imgShield(intShields).Visible = False
                            
                    End If
                    
                End If
                
                If fnCheckCollision(imgPlayer(0), imgEnemy(i)) = True And .Destroyed = False Then 'CHECK FOR PLAYER1-ENEMY COLLISION
                    .Destroyed = True
                    pShip(0).Health = pShip(0).Health - nmeShip(i).MaxHealth
                    pShip(0).Score = pShip(0).Score + 1
                    imgEnemyHPBar(i).Visible = False
                    
                ElseIf fnCheckCollision(imgPlayer(1), imgEnemy(i)) = True And .Destroyed = False And intPlayerMode = 2 Then 'CHECK FOR PLAYER2-ENEMY COLLISION
                    .Destroyed = True
                    pShip(1).Health = pShip(1).Health - nmeShip(i).MaxHealth
                    pShip(1).Score = pShip(1).Score + 1
                    imgEnemyHPBar(i).Visible = False
                        
                        
                Else
                    For n = 0 To imgPlayer1Bolt.UBound 'CHECK FOR BOLT-ENEMY COLLISION, P1
                        If (fnCheckCollision(imgPlayer1Bolt(n), imgEnemy(i)) = True And imgPlayer1Bolt(n).Visible And imgEnemy(i).Visible = True) Then 'MID BOLT
                            .Health = .Health - pShip(0).FirePower
                            If pShip(0).CannonType <> 2 Then
                                imgPlayer1Bolt(n).Left = -5000
                                
                            End If
                                
                            If .Health < 0 Then
                                .Destroyed = True
                                imgEnemyHPBar(i).Visible = False
                                
                                pShip(0).Score = pShip(0).Score + 1
                                    
                            Else
                                imgEnemyHPBar(i).Visible = True
                                
                            End If
                            
                        End If
                                
                        If pShip(0).CannonType = 1 Then
                            If fnCheckCollision(imgPlayer1BoltL(n), imgEnemy(i)) = True And imgPlayer1BoltL(n).Visible And imgEnemy(i).Visible = True Then 'LBOLT
                                .Health = .Health - pShip(0).FirePower
                                imgPlayer1BoltL(n).Top = -5000
                                    
                                If .Health < 0 Then
                                    .Destroyed = True
                                    imgEnemyHPBar(i).Visible = False
                                    
                                    pShip(0).Score = pShip(0).Score + 1
                                        
                                Else
                                    imgEnemyHPBar(i).Visible = True
                                    
                                End If
                                
                                
                            End If
                                
                            If fnCheckCollision(imgPlayer1BoltR(n), imgEnemy(i)) = True And imgPlayer1BoltR(n).Visible And imgEnemy(i).Visible = True Then 'RBOLT
                                .Health = .Health - pShip(0).FirePower
                                imgPlayer1BoltR(n).Top = -5000
                                                                  
                                If .Health < 0 Then
                                    .Destroyed = True
                                    imgEnemyHPBar(i).Visible = False
                                    
                                    pShip(0).Score = pShip(0).Score + 1
                                    
                                Else
                                    imgEnemyHPBar(i).Visible = True
                                    
                                End If
                                    
                            End If
                        
                        End If
                            
                    Next n
                    
                    If intPlayerMode = 2 Then
                        
                        For n = 0 To imgPlayer2Bolt.UBound 'CHECK FOR BOLT-ENEMY COLLISION, P2
                            If (fnCheckCollision(imgPlayer2Bolt(n), imgEnemy(i)) = True And imgPlayer2Bolt(n).Visible And imgEnemy(i).Visible = True) Then 'MID BOLT
                                .Health = .Health - pShip(1).FirePower
                                
                                If pShip(1).CannonType <> 2 Then
                                    imgPlayer2Bolt(n).Left = -5000
                                
                                End If
                                    
                                If .Health < 0 Then
                                    .Destroyed = True
                                    imgEnemyHPBar(i).Visible = False
                                    
                                    pShip(1).Score = pShip(1).Score + 1
                                        
                                Else
                                    imgEnemyHPBar(i).Visible = True
                                    
                                End If
                                
                            End If
                                    
                            If pShip(1).CannonType = 1 Then
                                If fnCheckCollision(imgPlayer2BoltL(n), imgEnemy(i)) = True And imgPlayer2BoltL(n).Visible And imgEnemy(i).Visible = True Then 'LBOLT
                                    .Health = .Health - pShip(1).FirePower
                                    imgPlayer2BoltL(n).Left = -5000
                                        
                                    If .Health < 0 Then
                                        .Destroyed = True
                                        imgEnemyHPBar(i).Visible = False
                                        
                                        pShip(1).Score = pShip(1).Score + 1
                                            
                                    Else
                                        imgEnemyHPBar(i).Visible = True
                                        
                                    End If
                                    
                                    
                                End If
                                    
                                If fnCheckCollision(imgPlayer2BoltR(n), imgEnemy(i)) = True And imgPlayer2BoltR(n).Visible And imgEnemy(i).Visible = True Then 'RBOLT
                                    .Health = .Health - pShip(1).FirePower
                                    imgPlayer2BoltR(n).Left = -5000
                                                                      
                                    If .Health < 0 Then
                                        .Destroyed = True
                                        imgEnemyHPBar(i).Visible = False
                                        
                                        pShip(1).Score = pShip(1).Score + 1
                                        
                                    Else
                                        imgEnemyHPBar(i).Visible = True
                                        
                                    End If
                                        
                                End If
                            
                            End If
                                
                        Next n
                        
                    End If
                    
                    If .Health > 0 Then
                        imgEnemyHPBar(i).Width = 1200 * (.Health / .MaxHealth)
                                    
                    End If
                        
                End If '!!!

            
            Else
                Call subNextFrame(imgEnemy(i), imgExplosion, .ExplosionCounter)
                .ExplosionCounter = .ExplosionCounter + 1
                
            
            End If
            
        End With
            
    Next i
        
    For i = 0 To imgPlayer1Bolt.UBound 'MOVE P1 MISSILES
        imgPlayer1Bolt(i).Top = imgPlayer1Bolt(i).Top - pShip(0).MissileSpeed
        
        If pShip(0).CannonType = 1 Then
            imgPlayer1BoltL(i).Top = imgPlayer1BoltL(i).Top - pShip(0).MissileSpeed
            imgPlayer1BoltR(i).Top = imgPlayer1BoltR(i).Top - pShip(0).MissileSpeed
            imgPlayer1BoltL(i).Left = imgPlayer1BoltL(i).Left - (pShip(0).MissileSpeed / 2)
            imgPlayer1BoltR(i).Left = imgPlayer1BoltR(i).Left + (pShip(0).MissileSpeed / 2)
    
        End If
            
    Next i
    
    If intPlayerMode = 2 Then
        For i = 0 To imgPlayer2Bolt.UBound 'MOVE P2 MISSILES
            imgPlayer2Bolt(i).Top = imgPlayer2Bolt(i).Top - pShip(1).MissileSpeed
            
            If pShip(1).CannonType = 1 Then
                imgPlayer2BoltL(i).Top = imgPlayer2BoltL(i).Top - pShip(1).MissileSpeed
                imgPlayer2BoltR(i).Top = imgPlayer2BoltR(i).Top - pShip(1).MissileSpeed
                imgPlayer2BoltL(i).Left = imgPlayer2BoltL(i).Left - (pShip(1).MissileSpeed / 2)
                imgPlayer2BoltR(i).Left = imgPlayer2BoltR(i).Left + (pShip(1).MissileSpeed / 2)
        
            End If
                
        Next i
    
    End If
    
    For i = 0 To imgEnemyBolt.UBound
        imgEnemyBolt(i).Top = imgEnemyBolt(i).Top + 200
        On Error Resume Next
        
    Next i
    
    If pShip(0).Score = intGoal Then
        Call subGameOver(True)
        
    ElseIf intPlayerMode = 2 Then
        If pShip(0).Score + pShip(1).Score = intGoal Then
            Call subGameOver(True)
        End If
        
    End If

End Sub

Private Sub tmrVar_Timer()

    If pShip(0).YSpeed > 0 Then
        ReDim Preserve nmeShip(UBound(nmeShip) + 1) As Enemy
        With nmeShip(UBound(nmeShip))
            Randomize
            .EnemyType = fnRndInt(0, 2)
            .Health = 50 + (50 * .EnemyType)
            .MaxHealth = .Health
            .Speed = 100 - (45 * .EnemyType)
            .OnFireCD = 50
            .EnemyIndex = UBound(nmeShip)
            
            Call subLoadNew(imgEnemy, imgEnemyType, .EnemyType)
            Call subLoadNew(imgEnemyHPBar, imgEnemyHPBar, 0)
                        
            imgEnemy(.EnemyIndex).Height = imgEnemyType(.EnemyType).Height
            imgEnemy(.EnemyIndex).Width = imgEnemyType(.EnemyType).Width
        
        End With
        
        
        With imgEnemy(imgEnemy.UBound)
            .Top = -picBg.Top - .Height
            .Left = fnRndInt(0, Me.ScaleWidth - .Width)
        End With
        
        
        Call subCenterPos(imgEnemyHPBar(imgEnemyHPBar.UBound), imgEnemy(imgEnemy.UBound), True)
        imgEnemyHPBar(imgEnemyHPBar.UBound).Visible = False
        
        tmrVar.Interval = 2000 / intPlayerMode
        
    End If

End Sub
