VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6435
   ScaleWidth      =   7710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdQuit 
      Caption         =   "Quit"
      Height          =   615
      Left            =   5880
      TabIndex        =   1
      Top             =   5760
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   5535
      Left            =   120
      ScaleHeight     =   5475
      ScaleWidth      =   7395
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Rain for 3rd Person View
'
'If you use anything I made plz give me some credit :)
'Specially the dds file and 3ds which I worked for a whole hour :)
'And if u could, email me if you used it on ur game. so i can take a peek of what I helped create :)
'
'
'By Pavel
'Credits to James And Sury
'Mail: shah_pavel@hotmail.com
'http://www.geocities.com/shah_pavel/



' Force explicit declarations
Option Explicit

' We declare TrueVision8.
Private TV8 As TVEngine

' We declare the tank as an mesh8 object
Private Tank As TVMesh

' We declare the camera
Private Camera As TVCamera

' We declare the landscape
Private Land As TVLandscape

' We declare the texture factory
Private TextureFactory As TVTextureFactory

' We the declare the scene
Private Scene As TVScene

' We declare the input engine.
Private InputEngine As TVInputEngine


' The loop.
Private DoLoop As Boolean

' We need a position for the tank
Private TankPosition As D3DVECTOR

' We need angles for the tank.
Private sngAngleX As Single
Private sngAngleY As Single

' We could have done this in many ways, but we added some smoothing to
' the movement se we need to declare two additional variables.
Private sngWalk As Single
Private sngStrafe As Single

' We declare here the brake factor needed to make the tank break.
Private sngBrake As Single

'<<<<<<<<<<<<<<<<<<<<<Initialize Rain>>>>>>>>>>>>>>>>>>>>>>>>
Private Rain As TVMesh 'the 3ds cylinder
Private V As Single 'For the scrolling
'<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>

Private Sub cmdQuit_Click()

    ' We have clicked on the "Quit" button, so we change the DoLoop.
    DoLoop = False

End Sub

Private Sub Form_Load()

    ' We have to create the TV8 object before anything else.
    Set TV8 = New TVEngine

    ' Set search directory for textures, objects, ...
    TV8.SetSearchDirectory App.Path

    ' We initialize TV8 in the picture box of the form.
    TV8.Init3DWindowedMode Picture1.hWnd

    ' We want to see the FPS.
    TV8.DisplayFPS = True
    
    ' Set the AngleSystem to degree
    TV8.SetAngleSystem TV_ANGLE_DEGREE

    ' We create the input object.
    Set InputEngine = New TVInputEngine

    ' We create the scene (the world).
    Set Scene = New TVScene

    ' We create the camera object
    Set Camera = New TVCamera

    ' New : We are not going to use the sky boxes to save some weight
    ' on the tutorial zipfiles. Instead, we will use a background color.
    Scene.SetSceneBackGround 0, 0.3, 0.9

    ' We need to create a new texture factory
    Set TextureFactory = New TVTextureFactory

    ' The land generation
    Set Land = New TVLandscape
        
    ' Generate the height of the land from the grayscale of the image.
    Land.GenerateHugeTerrain "Track.jpg", TV_PRECISION_LOW, 8, 8, -700, -1024, True
    
    ' Then, we load the land texture.
    TextureFactory.LoadTexture "sand.jpg", "LandTexture"
    
    ' We assign a texture to that land.
    Land.SetTexture GetTex("LandTexture")
    
    ' New : we ceate a mesh by loading an external model (no, not a
    ' teapot). We are going to load the tank model. Let's start by
    ' creating the object that will hold our model.
    Set Tank = New TVMesh
    
    ' Now, load our model.
    Set Tank = Scene.CreateMeshBuilder
    Tank.Load3DSMesh "tank.3ds", False
    
    ' Load the texture and assign it.
    TextureFactory.LoadTexture "tank.bmp", "TankTexture", , , TV_COLORKEY_NO
    Tank.SetTexture GetTex("TankTexture")
    
    ' Let set the initial position of the tank
    TankPosition.x = 50
    TankPosition.z = 50
    TankPosition.y = Land.GetHeight(TankPosition.x, TankPosition.z)
    
    ' We set the camera vectors to be on the tank. To have a cool zoom in
    ' effect when we start the game, we set the camera really far away
    ' from the tank.
    Camera.SetPosition TankPosition.x, TankPosition.y + 200, TankPosition.z
    
    ' We set the initial values of movement
    sngWalk = 0
    sngStrafe = 0
    
  
'<<<<<<<<<<<<<<<<<<<<<<<<LOADING THE RAIN>>>>>>>>>>>>>>>>>>>>>>>>>
'<<<<<<<<<<<<<<<<<<<<<Load Rain.3DS>>>>>>>>>>>>>>>>>>>>>>>>>
    
    Set Rain = Scene.CreateMeshBuilder
    Rain.Load3DSMesh "rain2.3ds"
    TextureFactory.LoadTexture "Texture1.dds", "RainTex", , , TV_COLORKEY_BLACK
    
    
Rain.SetTexture GetTex("RainTex")
Rain.SetColor RGBA(1, 1, 1, 0.6)   '<<<<<<<<<<<<<<<<<  Change the last value for the transperency of the rain.
Rain.SetTextureModEnable True 'we use this line so that we can scroll the texture in the loop



    '<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
    
    ' We pop the form over everything else.
    Form1.Show
        
    ' We start the main loop.
    DoLoop = True
    Main_Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ' The user asked to quit but clicked on the 'X' button at up right.
    DoLoop = False
    
    ' And ask to quit.
    Main_Quit

End Sub

Private Sub Main_Loop()

    ' The main loop
    Do
        ' Let us the capacity to use buttons of the form.
        DoEvents
        
        ' We check the input
        Check_Input


        
        
        
        ' We check and update the movement
        Check_Movement

        ' Clear the the last frame.
        TV8.Clear
        
        ' Render the mesh



        ' New : we have to render the landscape.
        Land.Render True
 Tank.Render


'Note. The Rain has to be the last thing u render
'<<<<<<<<<<<<<<<<<<<<<<<Render and Scroll>>>>>>>>>>>>>>>>>>>>
V = V - TV8.AccurateTimeElapsed * 0.003
Rain.ScaleMesh 1, 1, 1
Rain.SetTextureModTranslationScale 0, V
Rain.Render
Rain.ScaleMesh 0.8, 0.8, 0.8
Rain.Render
Rain.ScaleMesh 0.5, 0.5, 0.5
Rain.Render
Rain.SetPosition TankPosition.x, TankPosition.y, TankPosition.z
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>

 
        ' We display everything that we have rendered
        TV8.RenderToScreen
    
    'We loop all of this until the DoLoop isn't True.
    Loop Until DoLoop = False
    
    ' We ask to quit.
    Main_Quit

End Sub

Private Sub Check_Input()
        
        ' Check if we pressed the UP arrow key, if so, then we are
        ' walking forward.

    
        If InputEngine.IsKeyPressed(TV_KEY_UP) = True Then
            
            sngWalk = 1
        
        End If
        
        ' Check if we pressed the DOWN arrow key, if so, then we are
        ' pressing the breaks.
        If InputEngine.IsKeyPressed(TV_KEY_DOWN) = True Then
            
            sngBrake = 0.002
        
        Else
        
            ' We are not pressing the brakes, let the tank float.
            sngBrake = 0.0005
        
        End If

        ' Check if we pressed the LEFT arrow key, if so, then strafe
        ' on the left.
        If InputEngine.IsKeyPressed(TV_KEY_LEFT) = True Then
            
            sngStrafe = 1
                
        ' If we are not strafing left, maybe we want to strafe to the
        ' right, using the RIGHT arrow? If so, set strafe to negative.
        ElseIf InputEngine.IsKeyPressed(TV_KEY_RIGHT) = True Then
        
            sngStrafe = -1
        
        End If

        ' Now, for the mouse input...
        Dim tmpMouseX As Long, tmpMouseY As Long

        ' Get the movement of the mouse.
        InputEngine.GetMouseState tmpMouseX, tmpMouseY

        ' Update the tank angle.
        sngAngleY = sngAngleY - (tmpMouseX / 100)

End Sub

Private Sub Check_Movement()

        ' Okay, now for the smothing of the movement... Update
        ' the forward and backward (walk) movement.
        Select Case sngWalk
        Case Is > 0
            sngWalk = sngWalk - sngBrake * TV8.TimeElapsed
            If sngWalk < 0 Then sngWalk = 0
        End Select
        
        ' Now, we update the left and right (strafe) movement.
        Select Case sngStrafe
        Case Is > 0
            sngStrafe = sngStrafe - 0.001 * TV8.TimeElapsed
            If sngStrafe < 0 Then sngStrafe = 0
        Case Is < 0
            sngStrafe = sngStrafe + 0.001 * TV8.TimeElapsed
            If sngStrafe > 0 Then sngStrafe = 0
        End Select
        
        ' Update the vectors using the angles and positions. But this
        ' time, we don't update the camera position but the tank
        ' vector position.
        TankPosition.x = TankPosition.x + (Cos(sngAngleY) * sngWalk / 5 * TV8.TimeElapsed) + (Cos(sngAngleY + 3.141596 / 2) * sngStrafe / 5 * TV8.TimeElapsed)
        TankPosition.z = TankPosition.z + (Sin(sngAngleY) * sngWalk / 5 * TV8.TimeElapsed) + (Sin(sngAngleY + 3.141596 / 2) * sngStrafe / 5 * TV8.TimeElapsed)
        TankPosition.y = Land.GetHeight(TankPosition.x, TankPosition.z) + 10
        
        ' From the tank position vector, we update the mesh position.
        Tank.SetPosition TankPosition.x, TankPosition.y, TankPosition.z
        
        ' From the angle variable, we update the tank rotation
        Tank.SetRotation 0, (sngAngleY * -57.295) + 90, 0
                
        ' With the new values of the tank vector, we update the camera,
        ' by using a dynamic camera that will follow the mesh.
        'Camera.ChaseCamera Tank, Vector(0, 25, -50), Vector(0, 0, 50), 120, True, 150, True
        Camera.ChaseCamera Tank, Vector(0, 45, -160), Vector(0, 0, 50), 20, True, 1050

End Sub

Private Sub Main_Quit()
        
    ' We want to quit the project, so we start by desroyng
    ' the texture factory.
    Set TextureFactory = Nothing
    
    ' We destroy the camera object
    Set Camera = Nothing
    
    ' We destroy the tank object
    Set Tank = Nothing
    
    ' We destroy the land object.
    Set Land = Nothing
    
    ' Don't forget to destroy the inputengine object...
    Set InputEngine = Nothing
    
    ' Then, we destroy the scene object.
    Set Scene = Nothing
    
    ' We finish the frenetic destroy with the TV8 object.
    Set TV8 = Nothing
    
    ' We end the application.
    End

End Sub

