Attribute VB_Name = "mdlGame"

'核心规则：黑8放在第三排中间
'犯规情况：空杆、界外球（这个游戏不考虑）、错击目标球、母球落袋
'需要注意 错击对方球要复位、但是自己方的球可以不用再次取出
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long


Private Type POOL_PLAYER
     Controller As Integer
     Colour As BALL_COLOURS
End Type

' 状态枚举量
Private Const PC_HUMAN = 1
Private Const PC_CPU = 2

Private Enum GAME_MODE
    GM_BALLS_IN_MOTION = 1
    GM_PLAYERS_AIMING_INIT = 2
    GM_PLAYERS_AIMING = 3
    GM_SHOWING_RESULTS_INIT = 4
    GM_SHOWING_RESULTS = 5
    GM_FREEBALL_INIT = 6
    GM_FREEBALL = 7
    GM_8BALL_INIT = 8
    GM_8BALL = 9
    GM_CHOOSECOLOR = 10
End Enum

Private Enum STICKER_CAPTIONS
    SC_PLAYER1 = 1
    SC_PLAYER2 = 2
    SC_FREEBALL = 3
    SC_EIGHTBALL = 4
    SC_WINS = 5
    SC_BREAKS = 6
    SC_FOUL = 7
    SC_CHOOSECOLOR = 8
    SC_RED = 9
    SC_YELLOW = 10
End Enum

'-----------------------------------------------------------------------------
' Private variables
'-----------------------------------------------------------------------------

'核心目录路径:
Public g_AppPath                As String

'DirectX8 对象:
Private m_DX                    As New DirectX8
Private m_D3DX                  As New D3DX8
Private m_D3D                   As Direct3D8
Private m_D3DDevice             As Direct3DDevice8

'视角
Private m_ViewPort              As D3DVIEWPORT8
Private m_ViewWdth              As Single
Private m_ViewHght              As Single

'球桌
Private m_Table                 As New PoolTable
'球桌参数
Private m_TableWidth            As Single
Private m_TableLength           As Single
Private m_GameAreaMinZ          As Single
Private m_GameAreaMaxZ          As Single
Private m_GameAreaMinX          As Single
Private m_GameAreaMaxX          As Single
'...and its pockets:
Private m_Pockets(1 To 6)       As D3DVECTOR
Private m_PocketRadius          As Single

'球
Private m_PoolBalls             As New PoolBalls
Private m_numPoolBalls          As Long
Private m_BPocketNumber()       As Integer

'摄像视角
Private m_Cameras(1 To 2)       As New PoolCamera
Private m_ActiveCam             As Integer
Private m_CameraToggleEnabled   As Boolean

'碰撞类
Private m_PCC                   As New PoolCollisionController

'特效
Private m_Light0                As D3DLIGHT8

'声音
Private m_DirectSound8          As DirectSound8
Private m_DSBillBillHit         As DirectSoundSecondaryBuffer8
Private m_DSBillTableHit        As DirectSoundSecondaryBuffer8
Private m_DSCueBallLaunched     As DirectSoundSecondaryBuffer8
Private m_DSPocketHit           As DirectSoundSecondaryBuffer8
Private m_VolumeBase            As Single

Private Const m_VolumeMultiplier = 2000!
Private Const m_MinVolume = -3000

'玩家:
Private m_Players(1 To 2)       As POOL_PLAYER
Private m_8BallPocket           As Integer
Private m_CurrentPlayer         As Integer
Private m_AndTheWinnerIs        As Integer
Private m_bPlayersSwap          As Boolean
Private m_bFoul                 As Boolean

'游戏设置:
Private m_bName8BallPocket      As Boolean
Private m_bToggleAfterLaunch    As Boolean

'Billboard:
Private m_Arrow                 As New PoolBillboard

Private Const m_numStickers = 10
Private m_Sprite                As D3DXSprite
Private m_Stickers _
        (1 To m_numStickers)    As New ScreenSticker
    
'当前模式:
Private m_CurrentGM             As GAME_MODE

'Boolean variable
'that remains true while
'the application is running:
Private m_bAppRunning           As Boolean

'A boolean indicating that
'a new game has been started:
Private m_bNewGame              As Boolean

'Time step and g_PI:
Public Const g_dt               As Single = 0.025
Public Const g_dTicks           As Long = g_dt * 1000
Public Const g_PI               As Single = 3.1415927
Private m_StartTick             As Long

'Helpers and iterators:
Private i As Long, j As Long, k As Long
Private mMtrx1 As D3DMATRIX
Private vVctr1 As D3DVECTOR
Private vVctr2 As D3DVECTOR
Private vVctr3 As D3DVECTOR

'-------------------------------
' Name: StopGame()
' Desc: Stops the game (duh...)
'-------------------------------
Sub StopGame()
    m_bAppRunning = False
End Sub

'---------------------------------------------
' Name: GetLightDir()
' Desc: Retrieves the light direction vector
'---------------------------------------------
Function GetLightDir() As D3DVECTOR4
    With m_Light0
        GetLightDir.x = .Direction.x
        GetLightDir.y = .Direction.y
        GetLightDir.z = .Direction.z
        If .Type = D3DLIGHT_DIRECTIONAL Then GetLightDir.w = 0 Else GetLightDir.w = 1
    End With
End Function

'---------------------------------------------------
' Name: FireCueBall()
' Desc: Sets the cue-ball's initial velocity vector
'---------------------------------------------------
Sub FireCueBall(ByVal Velocity As Single)
    Dim vDir As D3DVECTOR
    
    D3DXVec3Subtract vDir, m_PoolBalls.BallPosition(0), m_Cameras(1).Position
    vDir.y = 0
    D3DXVec3Normalize vDir, vDir
    If m_CurrentGM = GM_PLAYERS_AIMING Then
        m_PoolBalls.FireCueBall ScaleVec3(vDir, Velocity)
        m_CurrentGM = GM_BALLS_IN_MOTION
        ' Play cue-ball launch sound:
        PlaySound m_DSCueBallLaunched, Velocity / 5
    End If
               
    If m_bToggleAfterLaunch And m_ActiveCam = 1 Then ToggleCameras
End Sub

'------------------------------------------------------
' Name: MouseEventHandler()
' Desc: Acts as a middle-man between the user
'       and objects (mainly cameras) in the game.
'------------------------------------------------------
Sub MouseEventHandler(ByVal Button As Integer, _
                      ByVal MouseX As Single, ByVal MouseY As Single, _
                      ByVal MouseMoveX As Single, ByVal MouseMoveY As Single)
          
    Dim vCueBall    As D3DVECTOR        ' Position vector of the cue-ball.
    Dim vCamPos     As D3DVECTOR        ' Position vector of the active camera.
    Dim vDist       As D3DVECTOR        ' A vector linking the camera with the cue-ball.
    Dim DistSq      As Single           ' Squared module of vDist vector.
    Dim mProj       As D3DMATRIX        ' Projection transformation matrix of the active camera (used for "unprojection")
    Dim mView       As D3DMATRIX        ' View transformation matrix of the active camera (same as sbove)
    Dim mWorld      As D3DMATRIX        ' World transformation matrix of the active camera (again same as above)
    Dim Lambda      As Single           ' A utility variable used for scaling vectors.
    
    Select Case m_CurrentGM
        Case GM_BALLS_IN_MOTION:
            '----------------------------------------------------------------------
            ' If the balls are moving the camera should rotate around itselft
            ' and translate within a plane parallel to the table.
            '----------------------------------------------------------------------
            
            Dim vShift      As D3DVECTOR    ' A vector representing a linear translation
                                            ' of the active cam along a specified line or plane.
            Dim mCamGen     As D3DMATRIX    ' Camera generator matrix.

            ' Camera #2 is fixed, so there is no point in responding for user input.
            If m_ActiveCam > 1 Then Exit Sub
            If Button = 1 Then
                ' When the left mouse button is pressed, the camera will rotate around itself.
                MouseMoveY = MouseMoveY * 3 * g_PI
                MouseMoveX = MouseMoveX * 3 * g_PI
                vCamPos = m_Cameras(m_ActiveCam).Position
                m_Cameras(m_ActiveCam).Pivot vCamPos, MouseMoveY, MouseMoveX
            ElseIf Button = 2 Then
                'When the right mouse button is pressed, the camera will translate.
                MouseMoveY = MouseMoveY * 2
                MouseMoveX = -MouseMoveX * 2
                mCamGen = m_Cameras(m_ActiveCam).Generators
                vShift = vec3(-mCamGen.m11 * MouseMoveX * 2 - mCamGen.m31 * MouseMoveY * 2, _
                              0, _
                              -mCamGen.m13 * MouseMoveX * 2 - mCamGen.m33 * MouseMoveY * 2)
                m_Cameras(m_ActiveCam).Move vShift
            End If
        Case GM_PLAYERS_AIMING
            '--------------------------------------------------------------------------
            ' If a human player is aiming the camera should rotate around the cue-ball
            ' and translate along a line that runs from the cue-ball to the camera
            ' to aid aiming. However, if the player is CPU driven, the camera should
            ' move in exactly the same way as it does, when the ball sare in motion.
            '--------------------------------------------------------------------------
            
            If m_ActiveCam <> 1 Then Exit Sub
            If Button = 1 Then
                MouseMoveY = MouseMoveY * 2 * g_PI
                MouseMoveX = MouseMoveX * 2 * g_PI
                If m_Players(m_CurrentPlayer).Controller = PC_CPU Then
                    vCamPos = m_Cameras(m_ActiveCam).Position
                    m_Cameras(m_ActiveCam).Pivot vCamPos, MouseMoveY, MouseMoveX
                Else
                    vCueBall = m_PoolBalls.BallPosition(0)
                    m_Cameras(m_ActiveCam).Pivot vCueBall, MouseMoveY, MouseMoveX
                End If
            ElseIf Button = 2 Then
                MouseMoveY = MouseMoveY * 4
                MouseMoveX = -MouseMoveX * 4
                mCamGen = m_Cameras(m_ActiveCam).Generators
                If m_Players(m_CurrentPlayer).Controller = PC_CPU Then
                    vShift = vec3(-mCamGen.m11 * MouseMoveX - mCamGen.m31 * MouseMoveY, _
                                  -mCamGen.m12 * MouseMoveX, _
                                  -mCamGen.m13 * MouseMoveX - mCamGen.m33 * MouseMoveY)
                    m_Cameras(m_ActiveCam).Move vShift
                Else
                    vCueBall = m_PoolBalls.BallPosition(0)
                    vCamPos = m_Cameras(m_ActiveCam).Position
                    D3DXVec3Subtract vDist, vCamPos, vCueBall
                    DistSq = D3DXVec3Dot(vDist, vDist)
                    If (DistSq < 0.04 And MouseMoveY < 0) Or (DistSq > 9 And MouseMoveY > 0) Then
                        m_Cameras(m_ActiveCam).StopMovement True, True, True
                    Else
                        vShift = vec3(-mCamGen.m31 * MouseMoveY, -mCamGen.m32 * MouseMoveY, -mCamGen.m33 * MouseMoveY)
                        m_Cameras(m_ActiveCam).Move vShift
                    End If
                End If
            End If
        Case GM_FREEBALL
            '----------------------------------------------------------------------------------------
            ' Free-ball is relativelly simple. There are no problems with the cameras
            ' as the only active camera allowed in this game mode is Camera #2, which is fixed.
            ' Mouse cursor position has a different role now - it indicates the spot on the table's
            ' surface where the user whants to place the cue-ball.
            '----------------------------------------------------------------------------------------
            
            Dim CueBallCleared  As Boolean      ' "True" if the new cue-ball's position (set by the user)
                                                ' is within the game area.
            If Button = 0 Then
                ' Button = 0 means that no button is pressed. This in turn means that the user is moving the
                ' cue-ball around. Make sure that it stays within the table and does not "hit" any other ball.
                ' Get the camera's position:
                vCamPos = m_Cameras(m_ActiveCam).Position
                ' Unproject the point specified by MouseX and MouseY:
                vVctr1 = vec3(MouseX, MouseY, 1)    'NOTE: z = 1 as we want the point to be placed at the back of the viewport.
                mView = m_Cameras(m_ActiveCam).ViewMatrix
                mProj = m_Cameras(m_ActiveCam).ProjectionMatrix
                D3DXMatrixIdentity mWorld
                D3DXVec3Unproject vVctr1, vVctr1, m_ViewPort, mProj, mView, mWorld
                ' vVctr1 is now the point (MouseX, MouseY, 1) in 3D space.
                ' Now all we need to do is run a line through this point and the camera position vector
                ' and find the spot where this line intersects with the table plane
                ' (i.e. when y = 0, though not exactly - explained later)
                D3DXVec3Subtract vVctr2, vVctr1, vCamPos
                Lambda = (0.05 - vCamPos.y) / vVctr2.y      ' y = 0.05 instead of y = 0 as 0.05 is the balls' radius.
                D3DXVec3Add vCueBall, vCamPos, ScaleVec3(vVctr2, Lambda)    ' Cue-ball's new position.
                ' Just to be sure that the cue-ball is on exactly the same level as other balls:
                vCueBall.y = 0.05 ' Without this line, the cue-ball tends to end up slightly below, or above y = 0.05.
                                  ' This has devastating consequences for collision responce.
                ' Now, check whether the new cue-ball position fits within the game area.
                If vCueBall.x > m_GameAreaMinX + 0.05 And vCueBall.x < m_GameAreaMaxX - 0.05 _
                And vCueBall.z > m_GameAreaMinZ + 0.05 And vCueBall.z < m_GameAreaMaxZ - 0.05 Then
                    ' If it does, set the CueBallCleared variable to true:
                    CueBallCleared = True
                    ' Now, check for any overlapping between the cue-ball and other balls:
                    For i = 1 To m_numPoolBalls - 1
                        D3DXVec3Subtract vVctr1, m_PoolBalls.BallPosition(i), vCueBall
                        ' If a ball overlapps with the cue-ball set CueBallCleared to false
                        If vVctr1.x * vVctr1.x + vVctr1.z * vVctr1.z < 0.1 * 0.1 Then CueBallCleared = False
                    Next i
                    ' The new cue-ball position can be applied to the cue-ball only if the ball is cleared:
                    If CueBallCleared Then m_PoolBalls.BallPosition(0) = vCueBall
                End If
            ElseIf Button = 1 Then
                ' When the left mouse button is pressed the user
                ' whants to place the cue-ball in the spot specified by MouseX and MouseY variables.
                ' Make sure that this point is within the game area.
                vCueBall = m_PoolBalls.BallPosition(0)
                If vCueBall.x > m_GameAreaMinX + 0.05 And vCueBall.x < m_GameAreaMaxX - 0.05 _
                And vCueBall.z > m_GameAreaMinZ + 0.05 And vCueBall.z < m_GameAreaMaxZ - 0.05 Then
                    ' If the spot fits within the table, change the current game mode to enable
                    ' aiming:
                    m_CurrentGM = GM_PLAYERS_AIMING_INIT
                End If
            End If
        Case GM_8BALL:
            '-----------------------------------------------------
            ' In this mode the user is expected to point, with
            ' the mouse cursor, at the pocket, they want to shoot
            ' the eight-ball into.
            '-----------------------------------------------------
            
            Dim v8BallTarget As D3DVECTOR    ' The result of unprojecting the mouse cursor coordinates
                                             ' onto the table's surface.
            If Button = 1 Then 'a pocket has been chosen.
                ' Get the camera's position.
                vCamPos = m_Cameras(m_ActiveCam).Position
                ' Unproject the coordinates specified by MouseX and MouseY:
                vVctr1 = vec3(MouseX, MouseY, 1)
                mView = m_Cameras(m_ActiveCam).ViewMatrix
                mProj = m_Cameras(m_ActiveCam).ProjectionMatrix
                D3DXMatrixIdentity mWorld
                D3DXVec3Unproject vVctr1, vVctr1, m_ViewPort, mProj, mView, mWorld
                D3DXVec3Subtract vVctr2, vVctr1, vCamPos
                Lambda = 0 - vCamPos.y / vVctr2.y      ' This time we use y = 0.
                D3DXVec3Add v8BallTarget, vCamPos, ScaleVec3(vVctr2, Lambda)
                ' If a distance between a pocket and the v8BallTarget point is smaller
                ' than the radius of a pocket then we've found the pocket indicated by
                ' the user.
                For i = 1 To 6
                    D3DXVec3Subtract vDist, m_Pockets(i), v8BallTarget
                    ' We will use the squared module of vDist to save on CPU time.
                    If (vDist.x * vDist.x + vDist.z * vDist.z) < 0.01 Then      'NOTE: 0.01 is 0.1 squared.
                        m_8BallPocket = i
                        m_Arrow.BasePoint = vec3(m_Pockets(i).x, 0.1, m_Pockets(i).z)
                        ' We found the right pocket, thus set the current game mode to aiming mode.
                        m_CurrentGM = GM_PLAYERS_AIMING_INIT
                    End If
                Next i
            End If
        Case GM_CHOOSECOLOR:
            If Button = 0 Then
                Call m_Stickers(SC_RED).UnderCursor(MouseX, MouseY)
                Call m_Stickers(SC_YELLOW).UnderCursor(MouseX, MouseY)
            ElseIf Button = 1 Then
                If m_Stickers(SC_RED).UnderCursor(MouseX, MouseY) Then
                    m_Players(m_CurrentPlayer).Colour = BC_RED
                    m_Players(3 - m_CurrentPlayer).Colour = BC_YELLOW
                    m_CurrentGM = GM_PLAYERS_AIMING_INIT
                ElseIf m_Stickers(SC_YELLOW).UnderCursor(MouseX, MouseY) Then
                    m_Players(m_CurrentPlayer).Colour = BC_YELLOW
                    m_Players(3 - m_CurrentPlayer).Colour = BC_RED
                    m_CurrentGM = GM_PLAYERS_AIMING_INIT
                End If
            End If
    End Select

End Sub

'---------------------------------------------
'主程序入口
'---------------------------------------------
Sub Main()
    '目录
    g_AppPath = App.Path
    
    
    ' 视角:
    frmPool.Picture1.ScaleMode = vbPixels
    With m_ViewPort
        .Height = frmPool.Picture1.ScaleHeight
        .Width = frmPool.Picture1.ScaleWidth - 1
        .MaxZ = 1
        .MinZ = 0
        .x = 0
        .y = 0
    End With
       
    On Error GoTo ErrorHandler
       
    ' Initialize D3D and D3DDevice
    InitD3D frmPool.Picture1.hWnd
    
    ' Clear the renderring surface and the z-buffer,
    ' then set the viewport:
    With m_D3DDevice
        .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0&, 1, 0
        .SetViewport m_ViewPort
    End With
    
    ' Setup lights:
    SetupLights
    
    ' Setup sounds:
    SetupSounds
        
    '初始化视角
    m_Cameras(1).Setup vec3(0, 0.5, -2), vec3(0, 0.1, -1), vec3(0, 1, 0), 10, 0.01, m_ViewPort.Height / m_ViewPort.Width
    ' The second camera will hang from the roof, pointing downwards:
    m_Cameras(2).Setup vec3(0, 3, 0), vec3(0, 0, 0), vec3(1, 0, 0), 10, 0.01, m_ViewPort.Height / m_ViewPort.Width
    ' Initially the active camera is No 2:
    m_ActiveCam = 2
    ' Disable camera toggle:
    m_CameraToggleEnabled = False
        
    ' Create the table:
    m_TableWidth = 2: m_TableLength = 4
    With m_Table
        .Create m_TableWidth, m_TableLength, m_D3DDevice, m_D3DX
        ' After creating the table "ask" it for the position and radius of its pockets.
        ' We will need these data for selecting the pocket for the eight-ball
        ' at the end of the game.
        .GetPockets m_Pockets, m_PocketRadius
        ' We also need to know, how large the game area really is.
        .GetGameArea m_GameAreaMinZ, m_GameAreaMaxZ, m_GameAreaMinX, m_GameAreaMaxX
    End With
                
    ' Create the balls:
    With m_PoolBalls
        .Create m_D3DDevice, m_D3DX
        ' After creating the balls, we can "ask" them for their number...
        m_numPoolBalls = .NumBalls
    End With
    ' ...so that we can set the dimensions of the array holding indexes of pockets,
    ' into which the balls fell:
    ReDim m_BPocketNumber(m_numPoolBalls - 1)
    
    ' Create the collision controller:
    m_PCC.Setup m_PoolBalls, m_Table
    
    ' Create the billboard with an arrow indicating the pocket
    ' for the 8'th ball:
    m_Arrow.Setup g_AppPath & "\Arrow.dds", 0.2, 0.2, m_D3DDevice, m_D3DX
    
    ' Create the D3DXSprite object to enable renderring screen stickers.
    Set m_Sprite = m_D3DX.CreateSprite(m_D3DDevice)
    ' Create the stickers:
    m_Stickers(SC_PLAYER1).Setup g_AppPath & "\Messages\Player1.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 305) / 2, (m_ViewPort.Height - 84) / 2, 305, 84, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_PLAYER2).Setup g_AppPath & "\Messages\Player2.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 316) / 2, (m_ViewPort.Height - 84) / 2, 316, 84, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_FREEBALL).Setup g_AppPath & "\Messages\FreeBall.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 335) / 2, (m_ViewPort.Height + 90) / 2, 335, 65, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_EIGHTBALL).Setup g_AppPath & "\Messages\EightBall.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 355) / 2, (m_ViewPort.Height + 90) / 2, 355, 83, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_WINS).Setup g_AppPath & "\Messages\Wins.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 176) / 2, (m_ViewPort.Height + 90) / 2, 176, 64, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_BREAKS).Setup g_AppPath & "\Messages\Breaks.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 260) / 2, (m_ViewPort.Height + 90) / 2, 260, 65, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_FOUL).Setup g_AppPath & "\Messages\Foul.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 161) / 2, (m_ViewPort.Height - 65) / 2, 161, 65, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_CHOOSECOLOR).Setup g_AppPath & "\Messages\ChooseColour.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 387) / 2, (m_ViewPort.Height - 43) / 2 + 60, 387, 43, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_RED).Setup g_AppPath & "\Messages\Red.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 77 - 200) / 2, (m_ViewPort.Height - 33) / 2 + 110, 77, 33, &HFFFFFFFF, m_D3DDevice, m_D3DX
    m_Stickers(SC_YELLOW).Setup g_AppPath & "\Messages\Yellow.dds", D3DFMT_A8R8G8B8, (m_ViewPort.Width - 131 + 200) / 2, (m_ViewPort.Height - 33) / 2 + 110, 131, 33, &HFFFFFFFF, m_D3DDevice, m_D3DX
        
    ' Start a new game
    StartNewGame
    
    ' Show the form:
    frmPool.Show
              
   
   '  Main game loop:
    Do
        Select Case m_CurrentGM
            Case GM_BALLS_IN_MOTION:
                UpdateBalls
            Case GM_PLAYERS_AIMING_INIT:
                PlayersTakingAimInit m_Players(m_CurrentPlayer).Controller
            Case GM_PLAYERS_AIMING:
                PlayersTakingAim m_Players(m_CurrentPlayer).Controller
            Case GM_FREEBALL_INIT:
                FreeBallInit
            Case GM_FREEBALL:
                FreeBall
            Case GM_8BALL_INIT:
                Set8BallTargetInit
            Case GM_8BALL:
                Set8BallTarget
            Case GM_SHOWING_RESULTS_INIT:
                ShowResultsInit
            Case GM_SHOWING_RESULTS:
                ShowResults
            Case GM_CHOOSECOLOR:
                ChooseColor
        End Select
    Loop While m_bAppRunning
    
    Cleanup
    End
  
ErrorHandler:
    MsgBox Err.Description & ". Bailing out...", vbCritical, " Error"
    Cleanup
    End
End Sub

'-------------------------------
' 开始新游戏
'-------------------------------
Sub StartNewGame()
    With m_Players(1): .Controller = PC_HUMAN: .Colour = BC_NONE: End With
    With m_Players(2): .Controller = PC_HUMAN: .Colour = BC_NONE: End With
    m_CurrentPlayer = 1
    m_CurrentGM = GM_PLAYERS_AIMING_INIT
    m_bAppRunning = True
    m_Arrow.Visible = False
    m_ActiveCam = 2
    m_CameraToggleEnabled = True
    m_bNewGame = True
    m_bFoul = False
    m_bPlayersSwap = False
    m_AndTheWinnerIs = 0
        
    ' Show the startup stickers:
    m_Stickers(SC_PLAYER1).Visible = True
    m_Stickers(SC_BREAKS).Visible = True
    ' Hide the others:
    For i = 1 To m_numStickers
        If i <> SC_PLAYER1 And i <> SC_BREAKS Then m_Stickers(i).Visible = False
    Next i
    
    ' Set the correct caption and font size on the lblPlayer label:
    With frmPool.lblPlayer
        .FontSize = 20
        .Caption = "玩家 " & m_CurrentPlayer & vbCrLf & "尚未选择颜色"
    End With
    
    ' Set the cursors for GM_PLAYERS_AIMING_INIT mode:
    frmPool.SetCursorType vbArrow, vbSizePointer, vbSizeNS
    
    ' 初始化球位置
    m_PoolBalls.InitialPositions
    
    ' Reset the 8'th ball's target pocket:
    m_8BallPocket = 0
    
End Sub

'-----------------------------------------------------------------------------
' Name: CleanUp()
' Desc: Clears all objects
'-----------------------------------------------------------------------------
Sub Cleanup()
    Set m_Table = Nothing
    Set m_PoolBalls = Nothing
    Set m_PCC = Nothing
    Set m_Arrow = Nothing
    Erase m_Cameras, m_BPocketNumber, m_Stickers
    
    Set m_DirectSound8 = Nothing
    Set m_DSBillBillHit = Nothing
    Set m_DSBillTableHit = Nothing
    Set m_DSCueBallLaunched = Nothing
    Set m_DSPocketHit = Nothing
    Set m_Sprite = Nothing
    Set m_D3DDevice = Nothing
    Set m_D3D = Nothing
    Set m_D3DX = Nothing
    Set m_DX = Nothing
    
    Set frmPool = Nothing
End Sub

'  游戏设置

Public Sub SetSoundVolumeBase(ByVal VolBase As Single)
    m_VolumeBase = VolBase
End Sub

Public Sub ToggleCameras()
    If m_CameraToggleEnabled Then m_ActiveCam = 3 - m_ActiveCam
End Sub

Public Sub NameEightBallPocket(bVal As Boolean)
    m_bName8BallPocket = bVal
End Sub

Public Sub ToggleCamAfterLaunch(bVal As Boolean)
    m_bToggleAfterLaunch = bVal
End Sub



'         回合循环



'----------------------------------------------------
' 更新场上球的位置 绘制动画
'----------------------------------------------------
Private Sub UpdateBalls()
    
    m_bNewGame = False
    Do While m_PoolBalls.AnyBallInMotion
        m_StartTick = GetTickCount
        DoEvents
        If m_bNewGame Then Exit Sub
        With m_PCC
            .RunCollisions
            ' 检查是否需要播放声音
            If .BillBillCollisionMomentum > 0 Then PlaySound m_DSBillBillHit, .BillBillCollisionMomentum * 2
            If .BillTableCollisionMomentum > 0 Then PlaySound m_DSBillTableHit, .BillTableCollisionMomentum
            If .PocketHitDetected Then PlaySound m_DSPocketHit, 0.5
        End With
        m_PoolBalls.NextFrame
        m_Cameras(m_ActiveCam).Update
        Render
        WaitUntill m_StartTick + g_dTicks
    Loop
    
    ' 所有球都进后判断谁胜
    ResolveLastStrike
End Sub


'--------------------------------------------
' Loop name: FreeBallInit()
' Loop desc:
'--------------------------------------------
Private Sub FreeBallInit()
    Dim iFrames     As Long
    Dim numFrames   As Long
                     
    ' Cursors for camera's transition:
    frmPool.SetCursorType vbArrow, vbArrow, vbArrow
            
    ' If there was a foul in the previous strike, show a message informing about it:
    If m_bFoul Then
        m_Stickers(SC_FOUL).Visible = True
        For iFrames = 0 To 20
            m_StartTick = GetTickCount
            DoEvents
            If m_bNewGame Then Exit Sub
        
            Render

            WaitUntill m_StartTick + g_dTicks
        Next iFrames
    End If
    m_bFoul = False
    m_Stickers(SC_FOUL).Visible = False
                       
    ' Show relevant messages:
    m_Stickers(m_CurrentPlayer).Visible = True
    m_Stickers(SC_FREEBALL).Visible = True
                       
    ' Cameras:
    If m_ActiveCam = 1 Then 'move the camera into the position of camera #2.
        m_Cameras(1).Transit vec3(0, 3, 0), vec3(0, 0, 0), vec3(1, 0, 0), 100
        ' The transition process requires some time:
        numFrames = 100
    Else
        ' In this case the camera remains on its place, thus we don't
        ' need to wait too long:
        numFrames = 30
    End If
    
    m_bNewGame = False
    For iFrames = 0 To numFrames
        m_StartTick = GetTickCount
        DoEvents
        If m_bNewGame Then Exit Sub
        
        m_Cameras(m_ActiveCam).Update
        Render

        WaitUntill m_StartTick + g_dTicks
    Next
    
    ' Set the lblPlayer's caption and font size for FreeBall loop:
    With frmPool.lblPlayer
        .FontSize = 12
        .Caption = "移动鼠标确定母球位置"
    End With
    
    ' Make all stickers disappear:
    For i = 1 To m_numStickers
        m_Stickers(i).Visible = False
    Next i
    
    frmPool.SetCursorType vbCrosshair
    m_CurrentGM = GM_FREEBALL
    m_ActiveCam = 2
End Sub

'--------------------------------------------
' Loop name: FreeBall()
' Loop desc:
'--------------------------------------------
Private Sub FreeBall()
    If m_Players(m_CurrentPlayer).Controller = PC_HUMAN Then m_ActiveCam = 2
    m_CameraToggleEnabled = False
    m_PoolBalls.ReappearCueBall
    Do
        DoEvents
        ' This loop will go on and on until the user places
        ' the cue-ball on the table.
        m_Cameras(m_ActiveCam).Update
        Render
    Loop While m_CurrentGM = GM_FREEBALL
    
    ' If the only balls left on the table are the cue-ball
    ' and the 8'th ball, then initiate a special loop, that
    ' enables the player to specify the pocket for the 8'th ball:
    If m_Players(m_CurrentPlayer).Colour <> BC_NONE Then
        Dim b8Ball As Boolean
        b8Ball = True
        For i = 1 To m_numPoolBalls - 1
            If i <> 8 Then
                If m_PoolBalls.InTheGame(i) And m_PoolBalls.Colour(i) = m_Players(m_CurrentPlayer).Colour Then
                    b8Ball = False
                End If
            End If
        Next i
        If b8Ball Then m_Players(m_CurrentPlayer).Colour = BC_BLACK
    End If
    
    ' Change the game's mode, from GM_FREEBALL to
    ' the aiming mode or the 8'th ball pocket designation mode.
    If b8Ball And m_bName8BallPocket Then
        m_CurrentGM = GM_8BALL_INIT
    Else
        m_CurrentGM = GM_PLAYERS_AIMING_INIT
    End If
    
End Sub


'--------------------------------------------
' Loop name: PlayersTakingAimInit()
' Loop desc: Sets the scene before calling
'            PlayersTakingAim.
'--------------------------------------------
Private Sub PlayersTakingAimInit(ByVal pc As Integer)
    Dim vLookAt     As D3DVECTOR
    Dim vCamPos     As D3DVECTOR
    Dim iFrames     As Long

    m_bNewGame = False
    
    ' If there was a foul in the previous strike, display an appropriate message:
    If m_bFoul Then
        m_Stickers(SC_FOUL).Visible = True
        For iFrames = 0 To 20
            m_StartTick = GetTickCount
            DoEvents
            If m_bNewGame Then Exit Sub
        
            Render

            WaitUntill m_StartTick + g_dTicks
        Next iFrames
    End If
    
    m_bFoul = False
    m_Stickers(SC_FOUL).Visible = False
    
    ' Show relevant messages:
    m_Stickers(m_CurrentPlayer).Visible = True
    If Not m_bName8BallPocket And m_Players(m_CurrentPlayer).Colour = BC_BLACK Then m_Stickers(SC_EIGHTBALL).Visible = True
    
    ' Set the correct caption and font size for the lblPlayer label:
    With frmPool.lblPlayer
        .FontSize = 20
        .Caption = "玩家 " & m_CurrentPlayer
        Select Case m_Players(m_CurrentPlayer).Colour
            Case BC_RED:    .Caption = .Caption & vbCrLf & "请打红球"
            Case BC_YELLOW: .Caption = .Caption & vbCrLf & "请打黄球"
            Case BC_BLACK:  .Caption = .Caption & vbCrLf & "请打黑球"
            Case BC_NONE:   .Caption = .Caption & vbCrLf & "请选择色球"
        End Select
    End With
       
    If pc = PC_CPU Then
        '
        ' Code for computer player aiming goes here
        '
        frmPool.SetCursorType vbArrow, vbSizePointer, vbSizePointer
    Else
        ' The first camera should be placed "behind" the cue-ball:
        vLookAt = m_PoolBalls.BallPosition(0)
        D3DXVec3Normalize vVctr1, vLookAt   ' A unit length vector, that points from the cue-ball to the table's centre.
        D3DXVec3Add vCamPos, vLookAt, ScaleVec3(vVctr1, 0.5)    ' Move 0.5 metre along the vVctr1 starting at the cue-ball,
                                                                ' to reach the point, where the camera will be placed.
        ' Set the camera's elevation above the table to 0.33m.
        vCamPos.y = 0.33
        ' The camera's line of sight shouldn't go through the middle of the cue-ball,
        ' but rather above it:
        vLookAt.y = vLookAt.y + 0.075
            
        ' Cursors for camera's transition:
        frmPool.SetCursorType vbArrow, vbArrow, vbArrow
                    
        ' Camera's transition loop:
        If m_ActiveCam = 2 Then m_Cameras(1).ChangeView vec3(0, 3, 0), vec3(0, 0, 0), vec3(1, 0, 0): m_ActiveCam = 1
        m_Cameras(1).Transit vCamPos, vLookAt, vec3(0, 1, 0), 100
        m_CameraToggleEnabled = False
        For iFrames = 0 To 100
            m_StartTick = GetTickCount
            DoEvents
            If m_bNewGame Then Exit Sub
            
            m_Cameras(1).Update
            Render
    
            WaitUntill m_StartTick + g_dTicks
        Next iFrames
        m_CameraToggleEnabled = True
        ' Cursors for aiming:
        frmPool.SetCursorType vbArrow, vbSizePointer, vbSizeNS
    End If
       
    ' Make all stickers disappear:
    For i = 1 To m_numStickers
        m_Stickers(i).Visible = False
    Next i
    
    m_CameraToggleEnabled = True
    m_CurrentGM = GM_PLAYERS_AIMING
End Sub

'-----------------------------------
' Loop name: PlayersTakingAim()
' Loop desc:
'-----------------------------------
Private Sub PlayersTakingAim(ByVal pc As Integer)
            
    If pc = PC_CPU Then
        '
        ' Code for computer player aiming goes here
        '
    Else
        Do
            m_StartTick = GetTickCount
            DoEvents
            
            ' Do nothing. Just wait for user's input.
            m_Cameras(m_ActiveCam).Update
            Render
            If m_ActiveCam = 1 Then frmPool.Picture1.Line (frmPool.Picture1.ScaleWidth / 2, 0)-(frmPool.Picture1.ScaleWidth / 2, frmPool.Picture1.ScaleHeight)
    
            WaitUntill m_StartTick + g_dTicks
        Loop While m_CurrentGM = GM_PLAYERS_AIMING
    End If
    
    frmPool.SetCursorType vbArrow, vbSizePointer, vbSizePointer
End Sub

'-----------------------------------------------------------
' Loop name: Set8BallTargetInit()
' Loop desc: Prepares the scene for the Set8BallTarget loop
'-----------------------------------------------------------
Private Sub Set8BallTargetInit()
    Dim iFrames As Long
    Dim numFrames As Long
    
    m_bNewGame = False
    
    ' If there was a foul in the previous strike, show a message informing about it:
    If m_bFoul Then
        m_Stickers(SC_FOUL).Visible = True
        For iFrames = 0 To 20
            m_StartTick = GetTickCount
            DoEvents
            If m_bNewGame Then Exit Sub
        
            Render

            WaitUntill m_StartTick + g_dTicks
        Next iFrames
    End If
    m_bFoul = False
    m_Stickers(SC_FOUL).Visible = False
       
    If m_ActiveCam = 1 Then 'move the camera into the position of camera #2.
        m_Cameras(1).Transit vec3(0, 3, 0), vec3(0, 0, 0), vec3(1, 0, 0), 100
        ' The transition process requires some time:
        numFrames = 100
    Else
        ' In this case the camera remains on its place, thus we don't
        ' need to wait too long:
        numFrames = 30
    End If
    
    ' Set the lblPlayer's caption and font size for Set8BallTarget loop:
    With frmPool.lblPlayer
        .FontSize = 12
        .Caption = "左键点击你想射入的决胜球（黑球）的球袋。"
    End With
    
    ' Show relevant messages:
    m_Stickers(m_CurrentPlayer).Visible = True
    m_Stickers(SC_EIGHTBALL).Visible = True
    
    ' Dissable camera toggling for the transition loop:
    m_CameraToggleEnabled = False
    For iFrames = 0 To numFrames
        m_StartTick = GetTickCount
        DoEvents
        If m_bNewGame Then Exit Sub
        
        m_Cameras(m_ActiveCam).Update
        Render

        WaitUntill m_StartTick + g_dTicks
    Next iFrames
    ' After completing the loop we can enable camera toggling:
    m_CameraToggleEnabled = True
       
    ' Make all stickers disappear:
    For i = 1 To m_numStickers
        m_Stickers(i).Visible = False
    Next i
    
    ' The Set8BallTarget loop requires that the player
    ' sees the table from the top:
    m_ActiveCam = 2
    m_CameraToggleEnabled = False
    ' Set the new game mode:
    m_CurrentGM = GM_8BALL
    ' Change the cursors on the form.
    frmPool.SetCursorType vbArrow, vbCrosshair, vbArrow

End Sub

'---------------------------------------------------------------------
' 设置关键球目标进球的球袋
'---------------------------------------------------------------------
Private Sub Set8BallTarget()
    If m_Players(m_CurrentPlayer).Controller = PC_HUMAN Then
        Do
            DoEvents
            m_Cameras(m_ActiveCam).Update
            Render
        Loop While m_CurrentGM = GM_8BALL
    Else
        
    End If
    m_CameraToggleEnabled = True
    m_Arrow.Visible = True
End Sub


'--------------------------------------------
' 准备输出结束游戏的信息
'--------------------------------------------
Private Sub ShowResultsInit()
    Dim iFrames As Long
    
    m_Arrow.Visible = False
    frmPool.SetCursorType
            
    ' 上一次出手犯规，那么打印相应信息
    If m_bFoul Then
        m_Stickers(SC_FOUL).Visible = True
        For iFrames = 0 To 20
            m_StartTick = GetTickCount
            DoEvents
            If m_bNewGame Then Exit Sub
        
            Render

            WaitUntill m_StartTick + g_dTicks
        Next iFrames
        m_Stickers(SC_FOUL).Visible = False
    End If
        
    ' 为ShowResults循环设置lblPlayer的标题和字体大小:
    With frmPool.lblPlayer
        .FontSize = 16
        .Caption = "F2 开始新游戏, ESC 退出游戏."
    End With
    
    ' 显示信息
    m_Stickers(m_AndTheWinnerIs).Visible = True
    m_Stickers(SC_WINS).Visible = True
    
    m_CameraToggleEnabled = False
    m_bNewGame = False
    If m_ActiveCam = 1 Then m_Cameras(1).Transit vec3(0, 3, 0), vec3(0, 0, 0), vec3(1, 0, 0), 100
    For i = 0 To 100
        m_StartTick = GetTickCount
        DoEvents
        If m_bNewGame Then Exit Sub
        
        m_Cameras(m_ActiveCam).Update
        Render

        WaitUntill m_StartTick + g_dTicks
    Next i
            
    ' Set the current game mode to GM_SHOWING_RESULTS
    m_CurrentGM = GM_SHOWING_RESULTS
End Sub

'--------------------------------------
'显示结果
'--------------------------------------
Private Sub ShowResults()
    m_bNewGame = False
    Do
        DoEvents
    Loop Until m_bNewGame
End Sub

'--------------------------------------
' 选择颜色循环
'--------------------------------------
Private Sub ChooseColor()
    Dim Color1 As D3DCOLORVALUE
    Dim Color2 As D3DCOLORVALUE
    
    If m_Players(m_CurrentPlayer).Controller = PC_HUMAN Then
        frmPool.SetCursorType
        Color1 = MakeD3DCOLORVALUE(1, 1, 1, 1)
        Color2 = MakeD3DCOLORVALUE(1, 1, 1, 0)
        m_Stickers(m_CurrentPlayer).Visible = True
        m_Stickers(SC_CHOOSECOLOR).Visible = True
        With m_Stickers(SC_RED)
            .Visible = True
            .SetupColorFlashing Color1, Color2, 10!
        End With
        With m_Stickers(SC_YELLOW)
            .Visible = True
            .SetupColorFlashing Color1, Color2, 10!
        End With
        Do
            DoEvents
            Render
            ' This loop will go on and on until the
            ' user picks a color.
        Loop While m_CurrentGM = GM_CHOOSECOLOR
    End If
    
    m_Stickers(SC_CHOOSECOLOR).Visible = False
    m_Stickers(SC_RED).Visible = False
    m_Stickers(SC_YELLOW).Visible = False
    
    m_CameraToggleEnabled = True
End Sub
    


' 核心代码


'-----------------------------------------------------------------------------
' 等待函数
'-----------------------------------------------------------------------------
Private Sub WaitUntill(ByVal LastTick As Single)
    Do
    Loop While GetTickCount < LastTick
End Sub


'-----------------------------------------------------------------------------
' 绘制帧
'-----------------------------------------------------------------------------
Private Sub Render()
    Dim mView           As D3DMATRIX
    Dim mProj           As D3DMATRIX
    Dim GSTex           As Direct3DTexture8
    Dim GSRect          As RECT
    Dim vGSRotCntr      As D3DVECTOR2
    Dim GSRot           As Single
    Dim vGSScale        As D3DVECTOR2
    Dim vGSTranslate    As D3DVECTOR2
    Dim GSColor         As Long

    
    If m_D3DDevice Is Nothing Then Exit Sub

    mView = m_Cameras(m_ActiveCam).ViewMatrix
    mProj = m_Cameras(m_ActiveCam).ProjectionMatrix
    m_D3DDevice.SetTransform D3DTS_VIEW, mView
    m_D3DDevice.SetTransform D3DTS_PROJECTION, mProj
    
    ' Clear the backbuffer to a black color, clear the z buffer to 1
    m_D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, &H0&, 1, 0
       
    m_D3DDevice.BeginScene
        ' The table...
        m_Table.Render
        ' ...and the balls:
        m_PoolBalls.Render
        ' Sprites:
        m_Sprite.Begin
            For j = 1 To m_numStickers
                If m_Stickers(j).Visible Then m_Stickers(j).Draw m_Sprite
            Next j
        m_Sprite.End
       ' The arrow billboard:
        If m_ActiveCam = 1 Then vVctr1 = vec3(0, 1, 0) Else vVctr1 = m_Arrow.BasePoint
        m_Arrow.Render m_Cameras(m_ActiveCam).Position, vVctr1
    m_D3DDevice.EndScene
            
    ' 将backbuffer内容显示到front buffer(屏幕)
    m_D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
    
End Sub

'----------------------------------------------------
'判断易手、获胜、输局 处理
'----------------------------------------------------
Private Sub ResolveLastStrike()
    Dim bPlayerChoosesColour    As Boolean
    Dim bPlayerScored           As Boolean
    Dim bPlayerWins             As Boolean
    Dim b8BallIn                As Boolean
    Dim bOnly8BallLeft          As Boolean
    Dim NewColour               As BALL_COLOURS
    Dim CurrentColour           As BALL_COLOURS
    
    m_bFoul = False ' 默认没有犯规 没有易手
    m_bPlayersSwap = False
    If m_Arrow.Visible Then m_Arrow.Visible = False
        
           
    '----------------------------
    ' 首先判断犯规情况
    '----------------------------
    ' 1) 母球进洞
         If m_PoolBalls.FellInPocketNumber(0) > 0 Then m_bFoul = True
    
    ' 2) 如果选择与被击中不同
         If m_Players(m_CurrentPlayer).Colour <> BC_NONE And Not m_bFoul Then
            If m_PCC.FirstBallTouchedColour <> m_Players(m_CurrentPlayer).Colour Then m_bFoul = True
         End If
         m_PCC.ClearFirstBallTouchedColour  'We won't be needing this any more.
    
    ' 3) 黑球—关键球进洞判断
         If m_PoolBalls.FellInPocketNumber(8) > 0 Then
            b8BallIn = True
            ' 有可能是最后胜利的标志、也有可能是自己其余七球未进造成犯规
            If m_bName8BallPocket Then
                If m_PoolBalls.FellInPocketNumber(8) <> m_8BallPocket Then
                    m_bFoul = True
                Else
                    bPlayerWins = True
                End If
            End If
         End If
    
    ' 4) 只要进的球有不同颜色 即视为犯规
         CurrentColour = BC_NONE: bPlayerChoosesColour = False  'Just to be sure...
         For i = 1 To m_numPoolBalls - 1
             If m_PoolBalls.FellInPocketNumber(i) > 0 Then 'the ball fell into a pocket.
                 bPlayerScored = True
                 NewColour = m_PoolBalls.Colour(i)
                 If m_Players(m_CurrentPlayer).Colour = BC_NONE Then
                     If NewColour = BC_BLACK Then m_bFoul = True: Exit For
                     If NewColour <> CurrentColour And (CurrentColour = BC_RED Or CurrentColour = BC_YELLOW) Then
                         bPlayerChoosesColour = True
                     Else
                         CurrentColour = NewColour
                     End If
                 ElseIf NewColour <> m_Players(m_CurrentPlayer).Colour Then m_bFoul = True
                 End If
             End If
         Next i
    '   如果还未选择色球 并且此杆全部击中所有球的颜色都相同那么就可以确定玩家的目标色球颜色了
        If Not bPlayerChoosesColour And CurrentColour <> BC_NONE Then
            m_Players(m_CurrentPlayer).Colour = CurrentColour
            m_Players(3 - m_CurrentPlayer).Colour = -1 * CurrentColour
        End If
        m_PoolBalls.ClearPocketNumbers
    
    '---------------------------------------------------
    ' We now have enough information to tell
    ' whether the player commited a foul and if so
    ' then what consequences will he have to suffer.
    '---------------------------------------------------
        If m_bFoul Then
            If b8BallIn Then
                ' 如果提前将关键球打入球袋 直接输局
                m_AndTheWinnerIs = 3 - m_CurrentPlayer
                m_CurrentGM = GM_SHOWING_RESULTS_INIT
            ElseIf m_Players(m_CurrentPlayer).Colour = BC_BLACK Then
                ' 如果在关键球模式下犯规，也算输
                m_AndTheWinnerIs = 3 - m_CurrentPlayer
                m_CurrentGM = GM_SHOWING_RESULTS_INIT
            Else
                m_CurrentGM = GM_FREEBALL_INIT
            End If
            '交换手
            m_bPlayersSwap = True: m_CurrentPlayer = 3 - m_CurrentPlayer
            Exit Sub
        End If
    
    '--------------------------------------------
    ' 正常事务逻辑
    '--------------------------------------------
    '1) 正常击球
        If Not bPlayerScored Then m_bPlayersSwap = True: m_CurrentPlayer = 3 - m_CurrentPlayer ' 1 2->a b
    
    '2)  ' 判断当该玩家的色球是否全部打完，进入关键球模式
        If m_Players(m_CurrentPlayer).Colour <> BC_NONE Then
            bOnly8BallLeft = True
            For i = 1 To m_numPoolBalls - 1
                If i <> 8 Then
                    If m_PoolBalls.InTheGame(i) And m_PoolBalls.Colour(i) = m_Players(m_CurrentPlayer).Colour Then
                        bOnly8BallLeft = False
                    End If
                End If
            Next i
            If bOnly8BallLeft Then m_Players(m_CurrentPlayer).Colour = BC_BLACK
        End If
    
    '-----------------------------------------
    ' 关键球模式 触发当且仅当场上只有一个黑球未进
    '-----------------------------------------
        If m_Players(m_CurrentPlayer).Colour = BC_BLACK Then
            If m_bName8BallPocket Then
                If b8BallIn And bPlayerWins Then
                    m_AndTheWinnerIs = m_CurrentPlayer
                    m_CurrentGM = GM_SHOWING_RESULTS_INIT
                ElseIf b8BallIn And Not bPlayerWins Then
                    m_AndTheWinnerIs = 3 - m_CurrentPlayer
                    m_CurrentGM = GM_SHOWING_RESULTS_INIT
                Else
                    m_CurrentGM = GM_8BALL_INIT
                End If
            Else
                If b8BallIn Then
                    m_AndTheWinnerIs = m_CurrentPlayer
                    m_CurrentGM = GM_SHOWING_RESULTS_INIT
                Else
                    m_CurrentGM = GM_PLAYERS_AIMING_INIT
                End If
            End If
        ElseIf bPlayerChoosesColour Then
            m_CurrentGM = GM_CHOOSECOLOR
        Else
            m_CurrentGM = GM_PLAYERS_AIMING_INIT
        End If
            
End Sub

' 初始化布局 此部分未修改
'----------------------------------------------------
' Name: InitD3D()
' Desc: Attempts to create Direct3DDevice8 object.
'----------------------------------------------------
Private Sub InitD3D(hWnd As Long)
    On Local Error GoTo RaiseError
    
    ' Create the D3D object:
    Set m_D3D = m_DX.Direct3DCreate()
    If m_D3D Is Nothing Then GoTo RaiseError
    
    ' Get The current Display Mode format
    Dim mode As D3DDISPLAYMODE
    m_D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, mode
         
    ' Set up the structure used to create the D3DDevice.
    Dim d3dpp As D3DPRESENT_PARAMETERS
    With d3dpp
        .Windowed = 1
        .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
        .BackBufferFormat = mode.Format
        .BackBufferCount = 1
        .EnableAutoDepthStencil = 1
        .AutoDepthStencilFormat = D3DFMT_D16
    End With

    ' Create the D3DDevice
    ' If you do not have hardware 3d acceleration. Enable the reference rasterizer
    ' using the DirectX control panel and change D3DDEVTYPE_HAL to D3DDEVTYPE_REF
    
    Set m_D3DDevice = m_D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, _
                                      D3DCREATE_SOFTWARE_VERTEXPROCESSING, d3dpp)
    If m_D3DDevice Is Nothing Then GoTo RaiseError
    
    ' Device state would normally be set here
    ' Turn off culling, so we see the front and back of the triangle
    With m_D3DDevice
        .SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
        .SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
        .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_DIFFUSE
        .SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
        .SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
                                      
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_DITHERENABLE, 1
        .SetRenderState D3DRS_FOGENABLE, 0
        .SetRenderState D3DRS_FILLMODE, 0
        .SetRenderState D3DRS_LASTPIXEL, 1
        .SetRenderState D3DRS_ZENABLE, 1
    End With

    Exit Sub
    
RaiseError:
    Err.Raise vbObjectError + 1, , "Unable to create device."
End Sub

' 部分特效 此处未修改
'--------------------------------------------------
' Name: SetupLights()
' Desc: Introduces a source of directional light
'       to the scene.
'--------------------------------------------------
Private Sub SetupLights()
    'We have one lamp above the table:
    With m_Light0
        .Type = D3DLIGHT_DIRECTIONAL
        .diffuse = MakeD3DCOLORVALUE(1, 1, 1, 1)
        .specular = MakeD3DCOLORVALUE(1, 1, 1, 1)
        .Direction = vec3(1, -3, 1)
    End With
    
    With m_D3DDevice
        .SetLight 0, m_Light0
        .LightEnable 0, 1
        .SetRenderState D3DRS_LIGHTING, 1
        .SetRenderState D3DRS_SPECULARENABLE, 1
        .SetRenderState D3DRS_AMBIENT, D3DColorRGBA(100, 100, 100, 0)
    End With
        
End Sub

'--------------------------------------------------
' 设置声音 四种 碰撞到球、击打长板、进库、发射
'--------------------------------------------------
Private Sub SetupSounds()
    
    ' Create a default DirectSound object:
    Set m_DirectSound8 = m_DX.DirectSoundCreate(vbNullString)
    ' Set the cooperation level:
    m_DirectSound8.SetCooperativeLevel frmPool.hWnd, DSSCL_PRIORITY
    
    Dim dsBufDesc As DSBUFFERDESC
    dsBufDesc.lFlags = DSBCAPS_CTRLVOLUME
    '导入四种文件，出错继续运行
    On Error Resume Next
    Set m_DSBillBillHit = m_DirectSound8.CreateSoundBufferFromFile(g_AppPath & "\BBH.wav", dsBufDesc)
    Set m_DSBillTableHit = m_DirectSound8.CreateSoundBufferFromFile(g_AppPath & "\BTH.wav", dsBufDesc)
    Set m_DSPocketHit = m_DirectSound8.CreateSoundBufferFromFile(g_AppPath & "\PH.wav", dsBufDesc)
    Set m_DSCueBallLaunched = m_DirectSound8.CreateSoundBufferFromFile(g_AppPath & "\CBL.wav", dsBufDesc)
       
End Sub

'--------------------------------------------------
' 播放音乐 以指定音量
'--------------------------------------------------
Private Sub PlaySound(ByVal dsBuffer As DirectSoundSecondaryBuffer8, ByVal Volume As Single)
    Dim FinalVolume As Long
    
    If Not (dsBuffer Is Nothing) And m_VolumeBase >= m_MinVolume Then
        FinalVolume = CLng(m_VolumeBase + Volume * m_VolumeMultiplier)
        If FinalVolume > 0 Then FinalVolume = 0
        With dsBuffer
            If Not .GetStatus = DSBSTATUS_PLAYING Then
                .SetVolume FinalVolume
                '.Stop
                '.SetCurrentPosition 0
                .Play DSBPLAY_DEFAULT
            End If
        End With
    End If
End Sub

'===================================
'  public 构造函数
'  by val 不会改变传值
'===================================

Public Function vec3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
    vec3.x = x: vec3.y = y: vec3.z = z
End Function

Public Function MakeD3DCOLORVALUE(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single) As D3DCOLORVALUE
    With MakeD3DCOLORVALUE: .r = r: .g = g: .b = b: .a = a: End With
End Function

Public Function DotProduct(v1 As D3DVECTOR, v2 As D3DVECTOR) As Single
    DotProduct = v1.x * v2.x + v1.y * v2.y + v1.z * v2.z
End Function

Public Function ScaleVec3(v1 As D3DVECTOR, s1 As Single) As D3DVECTOR
    With ScaleVec3: .x = v1.x * s1: .y = v1.y * s1: .z = v1.z * s1: End With
End Function
