VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmPool 
   Appearance      =   0  'Flat
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Billiards"
   ClientHeight    =   8025
   ClientLeft      =   885
   ClientTop       =   2565
   ClientWidth     =   11505
   Icon            =   "frmPool.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   535
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   767
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picFire 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   4927
      Picture         =   "frmPool.frx":08CA
      ScaleHeight     =   750
      ScaleWidth      =   1650
      TabIndex        =   3
      Top             =   6615
      Width           =   1650
   End
   Begin VB.PictureBox picPower 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   225
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "击球力度调整"
      Top             =   6600
      Width           =   3750
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      ForeColor       =   &H0080FF80&
      Height          =   5535
      Left            =   0
      ScaleHeight     =   367
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   765
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11505
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   75
      Left            =   9360
      TabIndex        =   4
      Top             =   7920
      Visible         =   0   'False
      Width           =   435
      URL             =   "\Caromhall.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   767
      _cy             =   132
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00202020&
      Height          =   525
      Left            =   7560
      TabIndex        =   2
      ToolTipText     =   "当前选手"
      Top             =   7260
      Width           =   3360
      WordWrap        =   -1  'True
   End
   Begin VB.Image imgPanel 
      BorderStyle     =   1  'Fixed Single
      Height          =   1620
      Left            =   -360
      Picture         =   "frmPool.frx":15EA
      Stretch         =   -1  'True
      Top             =   360
      Width           =   11460
   End
   Begin VB.Menu mnuGame 
      Caption         =   "游戏(&G)"
      Begin VB.Menu mnuGameNew 
         Caption         =   "新游戏(&N)"
      End
      Begin VB.Menu mnuGameBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuStgs 
      Caption         =   "设置(&S)"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout 
         Caption         =   "暂停(&A)"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSV 
         Caption         =   "游戏音量(&V)"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnuSVMute 
            Caption         =   "静音(&1)"
         End
         Begin VB.Menu mnuSVQuiet 
            Caption         =   "安静(&2)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSVAve 
            Caption         =   "均衡(&3)"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuSVLoud 
            Caption         =   "大声(&4)"
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuTG 
         Caption         =   "切换镜头                   Tab"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnuN8BP 
         Caption         =   "需要指定黑球8的球袋"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTCL 
         Caption         =   "在开出白球后切换镜头"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POWER_INDICATOR
    Left        As Single
    Top         As Single
    Width       As Single
    Height      As Single
    MaxLeft     As Single
    MinLeft     As Single
End Type

Private m_CurrentX          As Single
Private m_CurrentY          As Single

' Cursors:
Private m_LButtonCur        As Integer
Private m_RButtonCur        As Integer
Private m_NoButtonCur       As Integer

Private m_PwrIndctr         As POWER_INDICATOR
Private m_Power             As Single
Private m_MaxPower          As Single

Private m_PicPowerRange     As StdPicture
Private m_PicPwrIndctr      As StdPicture
Private m_PicPwrIndctrMask  As StdPicture

Private Sub Form_Load()
    m_MaxPower = 7
    
    Set m_PicPwrIndctr = LoadPicture(g_AppPath & "\PowerIndicator.bmp")
    Set m_PicPwrIndctrMask = LoadPicture(g_AppPath & "\PowerIndicatorMask.bmp")
    Set m_PicPowerRange = LoadPicture(g_AppPath & "\PowerRange.bmp")
    
    With Me
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
        .ScaleHeight = 750
        .ScaleWidth = 1000
    End With
    
    ' Main picture box:
    Picture1.Height = 600
    
    ' Controll panel:
    With imgPanel
        .Left = 0
        .Top = 600
        .Width = 1000
        .Height = 150
    End With
    
    With WindowsMediaPlayer1
        .URL = g_AppPath & "\Caromhall.mp3"
    End With
    
    ' Power setting picture box:
    With picPower
        .ScaleHeight = 80
        .ScaleWidth = 250
        .Left = 20
        .Top = 620
        .Height = 114
        .Width = 330
        .AutoRedraw = True
        .PaintPicture m_PicPowerRange, 0, 0, .ScaleWidth, .ScaleHeight
        .AutoRedraw = False
    End With
    
    ' Power indicator
    With m_PwrIndctr
        .Height = 34
        .Width = 20
        .MaxLeft = 219
        .MinLeft = 9
        .Left = (.MaxLeft - .MinLeft) / 2 + .MinLeft
        .Top = 14
    End With
    
    ' Initial power value:
    m_Power = (m_PwrIndctr.Left - m_PwrIndctr.MinLeft + m_PwrIndctr.Width / 2) / (m_PwrIndctr.MaxLeft - m_PwrIndctr.MinLeft) * m_MaxPower
           
    ' The "shoot button" label:
    With picFire
        .Left = (Me.ScaleWidth - .Width) / 2
        .Top = 640
    End With
       
    ' The label displaying current player number:
    With lblPlayer
        .FontSize = 24
        .Left = Me.ScaleWidth - .Width - 20
        .Top = 620
    End With
      
    'Sounds:
    mnuSVAve.Checked = True
    mnuSVAve_Click
    
    'Default settings:
    mnuTCL_Click
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace: FireCueBall m_Power
        Case vbKeyEscape: Form_Unload 0
        Case vbKeyTab: ToggleCameras
        Case vbKeyF2: StartNewGame
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopGame
    Set m_PicPwrIndctrMask = Nothing
    Set m_PicPwrIndctr = Nothing
    Set m_PicPowerRange = Nothing
    
    MsgBox "欢迎下次来玩", , "Billiards"
    End
End Sub

Private Sub lblShoot_Click()
    FireCueBall m_Power
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

'--------------------
'     GAME MENU:
'--------------------
Private Sub mnuGameExit_Click()
    Form_Unload 0
End Sub

Private Sub mnuGameNew_Click()
    StartNewGame
End Sub




Private Sub picFire_Click()
    FireCueBall m_Power
End Sub

Private Sub picPower_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With m_PwrIndctr
            .Left = x - .Width / 2
            If .Left < .MinLeft Then .Left = .MinLeft
            If .Left > .MaxLeft Then .Left = .MaxLeft
        End With
        With picPower
            .Cls
            .PaintPicture m_PicPwrIndctrMask, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcPaint
            .PaintPicture m_PicPwrIndctr, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcAnd
        End With
        m_Power = (m_PwrIndctr.Left - m_PwrIndctr.MinLeft + m_PwrIndctr.Width / 2) / (m_PwrIndctr.MaxLeft - m_PwrIndctr.MinLeft) * m_MaxPower
    End If
End Sub

Private Sub picPower_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With m_PwrIndctr
            .Left = x - .Width / 2
            If .Left < .MinLeft Then .Left = .MinLeft
            If .Left > .MaxLeft Then .Left = .MaxLeft
        End With
        With picPower
            .Cls
            .PaintPicture m_PicPwrIndctrMask, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcPaint
            .PaintPicture m_PicPwrIndctr, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcAnd
        End With
        m_Power = (m_PwrIndctr.Left - m_PwrIndctr.MinLeft + m_PwrIndctr.Width / 2) / (m_PwrIndctr.MaxLeft - m_PwrIndctr.MinLeft) * m_MaxPower
    End If
End Sub

Private Sub picPower_Paint()
    ' Power indicator:
    picPower.PaintPicture m_PicPwrIndctrMask, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcPaint
    picPower.PaintPicture m_PicPwrIndctr, m_PwrIndctr.Left, m_PwrIndctr.Top, m_PwrIndctr.Width, m_PwrIndctr.Height, , , , , vbSrcAnd
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    m_CurrentX = x
    m_CurrentY = y
    MouseEventHandler Button, x, y, 0, 0
    
    ' Set the cursor
    If Button = 1 Then
        Picture1.MousePointer = m_LButtonCur
    ElseIf Button = 2 Then
        Picture1.MousePointer = m_RButtonCur
    End If
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim ShiftX As Single
    Dim ShiftY As Single
            
    ShiftX = x - m_CurrentX
    ShiftY = y - m_CurrentY
    m_CurrentX = x: m_CurrentY = y
    If m_CurrentX >= 0 And m_CurrentX <= Picture1.ScaleWidth And m_CurrentY >= 0 And m_CurrentY <= Picture1.ScaleHeight Then
        ShiftX = ShiftX / Picture1.ScaleWidth * 0.2
        ShiftY = ShiftY / Picture1.ScaleHeight * 0.2
        MouseEventHandler Button, x, y, ShiftX, ShiftY
    End If
    
    ' Set the cursor
    If Button = 1 Then
        Picture1.MousePointer = m_LButtonCur
    ElseIf Button = 2 Then
        Picture1.MousePointer = m_RButtonCur
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Picture1.MousePointer = m_NoButtonCur
End Sub

Friend Sub SetCursorType(Optional NoButtonCur As Integer, Optional LButtonCur As Integer, Optional RButtonCur As Integer)
    m_NoButtonCur = NoButtonCur
    m_LButtonCur = LButtonCur
    m_RButtonCur = RButtonCur
    Picture1.MousePointer = m_NoButtonCur
End Sub


Private Sub Picture1_DblClick()
    FireCueBall m_Power
End Sub


Private Sub Form_Resize()
imgPanel.Top = frmPool.ScaleHeight - imgPanel.Height + 25
imgPanel.Width = frmPool.ScaleWidth
picPower.Top = imgPanel.Top + 12
picFire.Top = imgPanel.Top + 16
picFire.Left = (frmPool.ScaleWidth - picFire.Width) / 2
Picture1.Height = frmPool.ScaleHeight - imgPanel.Height + 25
lblPlayer.Top = imgPanel.Top + 16
lblPlayer.Left = frmPool.ScaleWidth - lblPlayer.Width
End Sub


