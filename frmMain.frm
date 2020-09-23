VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public m_Texture As cTexture        ' Our own texture


Public Function DrawGLScene() As Boolean
    Static xrot As GLfloat
    Static yrot As GLfloat
    Static zrot As GLfloat

    ' Clear the backbuffer and the depth buffer
    glClear clrColorBufferBit Or clrDepthBufferBit
    ' Reset the modelview matrix
    glLoadIdentity
    
    ' Translate out of the scene
    glTranslatef 0#, 0#, gflZ
    ' Rotate the scene along the x and y axis
    glRotatef xrot, 1#, 0#, 0#
    glRotatef yrot, 0#, 1#, 0#

    ' Use our texture
    m_Texture.useTexture

    ' Build a cube using quads
    glBegin GL_QUADS
        ' Front Face
        glNormal3f 0#, 0#, 1#                              'Set the surface normal
        glTexCoord2f 0#, 0#: glVertex3f -1#, -1#, 1#       'Bottom Left Of The Texture and Quad
        glTexCoord2f 1#, 0#: glVertex3f 1#, -1#, 1#        'Bottom Right Of The Texture and Quad
        glTexCoord2f 1#, 1#: glVertex3f 1#, 1#, 1#         'Top Right Of The Texture and Quad
        glTexCoord2f 0#, 1#: glVertex3f -1#, 1#, 1#        'Top Left Of The Texture and Quad
        ' Back Face
        glNormal3f 0#, 0#, -1#                             'Set the surface normal
        glTexCoord2f 1#, 0#: glVertex3f -1#, -1#, -1#      'Bottom Right Of The Texture and Quad
        glTexCoord2f 1#, 1#: glVertex3f -1#, 1#, -1#       'Top Right Of The Texture and Quad
        glTexCoord2f 0#, 1#: glVertex3f 1#, 1#, -1#        'Top Left Of The Texture and Quad
        glTexCoord2f 0#, 0#: glVertex3f 1#, -1#, -1#       'Bottom Left Of The Texture and Quad
        ' Top Face
        glNormal3f 0#, 1#, 0#                              'Set the surface normal
        glTexCoord2f 0#, 1#: glVertex3f -1#, 1#, -1#       'Top Left Of The Texture and Quad
        glTexCoord2f 0#, 0#: glVertex3f -1#, 1#, 1#        'Bottom Left Of The Texture and Quad
        glTexCoord2f 1#, 0#: glVertex3f 1#, 1#, 1#         'Bottom Right Of The Texture and Quad
        glTexCoord2f 1#, 1#: glVertex3f 1#, 1#, -1#        'Top Right Of The Texture and Quad
        ' Bottom Face
        glNormal3f 0#, -1#, 0#                             'Set the surface normal
        glTexCoord2f 1#, 1#: glVertex3f -1#, -1#, -1#      'Top Right Of The Texture and Quad
        glTexCoord2f 0#, 1#: glVertex3f 1#, -1#, -1#       'Top Left Of The Texture and Quad
        glTexCoord2f 0#, 0#: glVertex3f 1#, -1#, 1#        'Bottom Left Of The Texture and Quad
        glTexCoord2f 1#, 0#: glVertex3f -1#, -1#, 1#       'Bottom Right Of The Texture and Quad
        ' Right face
        glNormal3f 1#, 0#, 0#                              'Set the surface normal
        glTexCoord2f 1#, 0#: glVertex3f 1#, -1#, -1#       'Bottom Right Of The Texture and Quad
        glTexCoord2f 1#, 1#: glVertex3f 1#, 1#, -1#        'Top Right Of The Texture and Quad
        glTexCoord2f 0#, 1#: glVertex3f 1#, 1#, 1#         'Top Left Of The Texture and Quad
        glTexCoord2f 0#, 0#: glVertex3f 1#, -1#, 1#        'Bottom Left Of The Texture and Quad
        ' Left Face
        glNormal3f -1#, 0#, 0#                             'Set the surface normal
        glTexCoord2f 0#, 0#: glVertex3f -1#, -1#, -1#      'Bottom Left Of The Texture and Quad
        glTexCoord2f 1#, 0#: glVertex3f -1#, -1#, 1#       'Bottom Right Of The Texture and Quad
        glTexCoord2f 1#, 1#: glVertex3f -1#, 1#, 1#        'Top Right Of The Texture and Quad
        glTexCoord2f 0#, 1#: glVertex3f -1#, 1#, -1#       'Top Left Of The Texture and Quad
    glEnd

    xrot = xrot + gflXSpeed
    yrot = yrot + gflYSpeed
    
    DrawGLScene = True
End Function

Public Function InitGL() As Boolean
    Dim aflLightAmbient(4) As GLfloat
    Dim aflLightDiffuse(4) As GLfloat
    Dim aflLightPosition(4) As GLfloat
    
    ' Create new texture
    Set m_Texture = New cTexture
    m_Texture.loadTexture App.Path & "\Data\Crate.tga", FILETYPE_TGA
    
    ' Enable texture mapping
    glEnable glcTexture2D
    ' Smooth shading
    glShadeModel smSmooth
    
    ' Set the clear colour
    glClearColor 0#, 0#, 0#, 0#
    ' Set the clear depth
    glClearDepth 1#
    
    ' Enable Z-buffer
    glEnable glcDepthTest
    ' Set test type
    glDepthFunc cfLEqual
    ' Best perspective correction
    glHint htPerspectiveCorrectionHint, hmNicest
      
    ' Ambient light settings
    aflLightAmbient(0) = 0.5
    aflLightAmbient(1) = 0.5
    aflLightAmbient(2) = 0.5
    aflLightAmbient(3) = 1#
    ' Diffuse light settings
    aflLightDiffuse(0) = 1#
    aflLightDiffuse(1) = 1#
    aflLightDiffuse(2) = 1#
    aflLightDiffuse(3) = 1#
    ' Light position settings
    aflLightPosition(0) = 0#
    aflLightPosition(1) = 0#
    aflLightPosition(2) = 2#
    aflLightPosition(3) = 1#
      
    ' Set the light's ambient and diffuse levels and its position
    glLightfv ltLight1, lpmAmbient, aflLightAmbient(0)
    glLightfv ltLight1, lpmDiffuse, aflLightDiffuse(0)
    glLightfv ltLight1, lpmPosition, aflLightPosition(0)
    
    ' Enable light1
    glEnable glcLight1
    
    InitGL = True
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Set the key to be pressed
    gbKeys(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    ' Set the key to be not pressed
    gbKeys(KeyCode) = False
End Sub

Private Sub Form_Load()
    Dim bFullscreen As Boolean
    Dim frm As frmMain
    Dim bLightSwitched As Boolean
    Dim bFilterSwitched As Boolean
    Dim bLightOn As Boolean
    Dim giCurrFilter As Integer

    ' Put us into fullscreen automatically
    bFullscreen = True
    bLightSwitched = False
    bFilterSwitched = False
    bLightOn = False
    gflZ = -5#

    ' Save the current display settings
    SaveDisplaySettings

    ' Show this form
    Me.Show
    ' Attempt to create a compatible GL window and set the display mode
    If (CreateGLWindow(Me, 640, 480, 32, bFullscreen) = False) Then
        Unload Me
    End If
    ' Attempt to set up OpenGL
    If (InitGL() = False) Then
        Unload Me
    End If
  
    ' Loop until the form is unloaded, process windows events every time we're not rendering
    Do While DoEvents()
        ' Render the scene, if it failed or the user has pressed the escape key then exit the program
        If (DrawGLScene() = False) Or (gbKeys(vbKeyEscape)) Then
            Unload Me
        Else
            ' Swap the front and back buffers to display what we've just rendered
            SwapBuffers Me.hDC
      
            ' Toggle lighting
            If (gbKeys(vbKeyL) = True) And (bLightSwitched = False) Then
                bLightOn = Not (bLightOn)
            
                If (bLightOn) Then
                    glEnable glcLighting
                Else
                    glDisable glcLighting
                End If
              
                bLightSwitched = True
            End If
      
            If (gbKeys(vbKeyL) = False) Then
                bLightSwitched = False
            End If
      
            ' Toggle filtering
            If (gbKeys(vbKeyF) = True) And (bFilterSwitched = False) Then
                giCurrFilter = m_Texture.getFilter
                giCurrFilter = giCurrFilter + 1
                If giCurrFilter > 2 Then giCurrFilter = 0
                Select Case giCurrFilter
                    Case 0:
                        m_Texture.setFilter FILTER_NEAREST
                    Case 1:
                        m_Texture.setFilter FILTER_LINEAR
                    Case 2:
                        m_Texture.setFilter FILTER_MIPMAPPED
                End Select
            
                bFilterSwitched = True
            End If
      
            If (gbKeys(vbKeyF) = False) Then
                bFilterSwitched = False
            End If
        
            ' Zoom in and out
            If (gbKeys(vbKeyPageUp) = True) Then
                gflZ = gflZ - 0.02
            End If
            
            If (gbKeys(vbKeyPageDown) = True) Then
                gflZ = gflZ + 0.02
            End If
            
            ' Increase / decrease cube's rotation amount
            If (gbKeys(vbKeyUp) = True) Then
                gflXSpeed = gflXSpeed - 0.01
            End If
            
            If (gbKeys(vbKeyDown) = True) Then
                gflXSpeed = gflXSpeed + 0.01
            End If
            
            If (gbKeys(vbKeyLeft) = True) Then
                gflYSpeed = gflYSpeed - 0.01
            End If
            
            If (gbKeys(vbKeyRight) = True) Then
                gflYSpeed = gflYSpeed + 0.01
            End If
            
            ' Key escape has been pressed, so exit the program!
            If (gbKeys(vbKeyEscape) = True) Then
                Unload Me
            End If
        End If
    Loop
End Sub

Private Sub Form_Resize()
    ' When the user resizes the form, tell OpenGL to update so that it renders to the right place!
    ' Primarily used when in windowed mode
    ReSizeGLScene ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Shut down OpenGL
    KillGLWindow Me
End Sub
