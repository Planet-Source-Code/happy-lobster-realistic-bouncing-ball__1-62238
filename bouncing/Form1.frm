VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Drag and release the ball"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9120
   LinkTopic       =   "Form1"
   ScaleHeight     =   428
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   608
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CheckBox chkTrails 
      BackColor       =   &H00000000&
      Caption         =   "Show Trails"
      ForeColor       =   &H000080FF&
      Height          =   195
      Left            =   4920
      TabIndex        =   2
      Top             =   240
      Width           =   1275
   End
   Begin VB.Timer tmrDrag 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7080
      Top             =   240
   End
   Begin VB.Timer tmrAnimate 
      Interval        =   10
      Left            =   6540
      Top             =   240
   End
   Begin VB.PictureBox picBounce 
      BorderStyle     =   0  'None
      Height          =   1110
      Left            =   1920
      MousePointer    =   15  'Size All
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   74
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   0
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Realistic Bouncing Ball"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   915
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5475
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Based on:
'Title: Flickerless, Smooth Animation using pure VB with No OCXs, DLLs or ASM!
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=13701&lngWId=1
' by Doug Puckett ( dpuckett@thelittleman.com )
'
' Uses modified bounce algorithm

Const AirResistance = 0.995 ' Air resistance constant
Const RollResistance = 0.982   ' Resistance of ground on ball when rolling
Const BounceCoefficient = 0.75 ' Bounce constant

Dim Ymove As Integer
Dim Xmove As Single

Dim YMoveAfterBounce As Single
Dim YLastMoveAtBounce As Single
Dim bRolling As Boolean

Dim iNumberOfBounces As Integer

' Used for dragging ball picture box
Dim iOldX As Integer
Dim iOldY As Integer

Dim iDragLeft As Integer
Dim iDragTop As Integer

' Has the glass sound been played?
Dim bGlassPlayed As Boolean

Dim lTrailColour As Long

Private Sub chkTrails_Click()
    Me.Cls
End Sub

Private Sub Form_Load()
    StopBouncing  ' Clear vars
    StartThrow  ' Start bouncing
End Sub



' Disable animation timer and reser variaibles
Private Sub StopBouncing()
    tmrAnimate.Enabled = False
    tmrDrag.Enabled = True
    
    iNumberOfBounces = 0
    
    bRolling = False
    
    YMoveAfterBounce = -1000
End Sub

' Enable animate timer
Private Sub StartThrow()
    tmrAnimate.Enabled = True
    tmrDrag.Enabled = False
    PlayWAVFile "swhoosh.wav", True
    lTrailColour = RGB(Rnd * 256, Rnd * 256, Rnd * 256)
End Sub

' Store the amount that the picture box has been dragged
Private Sub tmrDrag_Timer()
    Static iLastX As Integer
    Static iLastY As Integer

    Xmove = iDragLeft - iLastX
    Ymove = -(iDragTop - iLastY)

    iLastX = iDragLeft
    iLastY = iDragTop
End Sub


Private Sub picBounce_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        StopBouncing
    
        ' Store drag initial positions
        iDragLeft = picBounce.Left
        iDragTop = picBounce.Top
        
        iOldX = X
        iOldY = Y
     End If
End Sub

Private Sub picBounce_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        tmrDrag.Enabled = True
        
        ' Drag picture box
        picBounce.Move iDragLeft + (X - iOldX), iDragTop + (Y - iOldY)
        
        ' Store current
        iDragLeft = picBounce.Left
        iDragTop = picBounce.Top
    End If
End Sub

' Throw the ball once dragging has finished
Private Sub picBounce_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        StartThrow
    End If
End Sub

Private Sub tmrAnimate_Timer()
    With picBounce
        Dim OldX As Integer ' Used for drawing trails
        Dim OldY As Integer
        
        OldX = .Left
        OldY = .Top
    
        ' Should the ball change direction?
        If .Top >= Me.ScaleHeight - (.Height) Then
            
            ' Set minimum position
            .Top = Me.ScaleHeight - (.Height)
            
            ' Invert y velocity
            Ymove = -Ymove * BounceCoefficient
        
            YMoveAfterBounce = -Ymove
            
            ' Count number of bounces
            If Not bRolling Then
                iNumberOfBounces = iNumberOfBounces + 1
                boing
                Me.Caption = iNumberOfBounces
            End If
            
            ' Workout if the ball is rolling
            If YMoveAfterBounce = YLastMoveAtBounce Then
                bRolling = True
            Else
                ' Restrain max speed
                YLastMoveAtBounce = YMoveAfterBounce
                If Ymove > 50 Then
                    Ymove = 50
                End If
            End If
        End If
        
        
        ' Don't throw the ball too high now! ;)
        If Not bGlassPlayed And .Top < -400 Then
            PlayWAVFile "glass.wav", True
            bGlassPlayed = True
        End If
        If .Top > 0 Then bGlassPlayed = False
        
    
        ' Horizontal ball constraights
        If .Left >= Me.ScaleWidth - .Width Then Xmove = -Xmove: boing
        If .Left <= 0 Then Xmove = -Xmove: boing
        
        ' Update Ball
        .Left = .Left + Xmove
        
        If Not bRolling Then
        
            ' Update ball y position
            .Top = .Top + (Ymove * -0.5)
            
            ' Apply effect of gravity on ball
            If Ymove >= 0 Then
                Ymove = Ymove - 1 ' Ball is ascending
            Else
                If Ymove > YMoveAfterBounce Then ' Ball is descending
                    Ymove = Ymove - 1
                End If
            End If
            
            ' Reduce ball speed due to air resistance
            Xmove = Xmove * AirResistance
            Ymove = Ymove * AirResistance
        Else
            If Abs(Xmove) <= 0.1 Then tmrAnimate.Enabled = False ' Test if ball has stopped moving
            Xmove = (Xmove * RollResistance) * AirResistance ' Reduce speed
        End If
        
        ' Draw trails if required
        Dim iOffset As Integer

        
        If chkTrails.Value = 1 Then
            Me.Line (OldX + .Width * 0.5, OldY + .Height * 0.5)-(.Left + .Width * 0.5, .Top + .Height * 0.5), lTrailColour
        End If
    End With
    
End Sub

Private Sub boing()
    PlayWAVFile "boing.wav", True
End Sub



