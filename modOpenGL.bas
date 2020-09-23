Attribute VB_Name = "modOpenGL"
Option Explicit

' Variables set by SaveDisplaySettings - used by RestoreDisplayMode
Global giOldWidth As Integer
Global giOldHeight As Integer
Global giOldNumBits As Integer
Global giOldVRefresh As Integer

Private hRC As Long     ' Handle to our Rendering Context


Public Function CreateGLWindow(frm As Form, iWidth As Integer, iHeight As Integer, iNumBits As Integer, bFullscreen As Boolean) As Boolean
    Dim bRet As Boolean                 ' Return value
    Dim pfd As PIXELFORMATDESCRIPTOR    '
    Dim lPixelFormat As Long            '
    Dim lRet As Long                    ' Another return value :)

    ' If we are in fullscreen mode then we need to do different things than if we want to run in windowed mode
    If (bFullscreen = True) Then
        ' Change to fullscreen
        bRet = SetFullscreenMode(iWidth, iHeight, iNumBits)
        ' If the call failed then
        If (bRet = False) Then
            ' Alert the user and carry on in windowed mode
            MsgBox "Could not set fullscreen mode.  App will continue in windowed mode."
        End If
    
        ' Set the caption property to nothing
        frm.Caption = ""
        ' Set the form border style
        frm.BorderStyle = 0
        ' Hide the cursor
        ShowCursor False
        ' Maximise the form
        frm.WindowState = vbMaximized
    Else
        ' Set the caption of the form
        frm.Caption = "The Gilb's TGA loading tutorial"
        ' Give the form a border (Fixed, single)
        frm.BorderStyle = 1
    End If
  
    ' Set the properties of the pixel format descriptor to allow OpenGL to use the window
    pfd.nVersion = 1
    pfd.dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    pfd.iLayerType = PFD_MAIN_PLANE
    pfd.iPixelType = PFD_TYPE_RGBA
    pfd.cColorBits = iNumBits
    pfd.cDepthBits = 16
    pfd.nSize = Len(pfd)
  
    ' Choose the desired pixel format
    lPixelFormat = ChoosePixelFormat(frm.hDC, pfd)
    If (lPixelFormat = 0) Then
        MsgBox "Could not get desired Pixel Format"
        CreateGLWindow = False
        Exit Function
    End If
  
    ' Set the desired pixel format
    lRet = SetPixelFormat(frm.hDC, lPixelFormat, pfd)
    If (lRet = 0) Then
        MsgBox "Could not set desired Pixel Format"
        CreateGLWindow = False
        Exit Function
    End If
  
    ' Create an OpenGL rendering context. The rendering context is what OpenGL uses to render to much like
    ' Windows's DC's.
    hRC = wglCreateContext(frm.hDC)
    If (hRC = 0) Then
        MsgBox "Could not create the rendering context"
        CreateGLWindow = False
        Exit Function
    End If
  
    ' Attach the RC to our DC
    lRet = wglMakeCurrent(frm.hDC, hRC)
    If (lRet = 0) Then
        MsgBox "Could not activate the RC"
        CreateGLWindow = False
        Exit Function
    End If
  
    ' Show the form
    frm.Show
    ' Set it as the foreground window
    SetForegroundWindow frm.hWnd
    ' Set it as the window with the focus
    SetFocus frm.hWnd
  
    ' Set the form's scalemode to pixels
    frm.ScaleMode = 3        'ScaleMode = Pixel
    ' Call Resize GL scene to update the viewport
    ReSizeGLScene frm.ScaleWidth, frm.ScaleHeight
  
    CreateGLWindow = True
End Function

'
' This function is called when we want to shut down OpenGL and restore the display mode
'
Public Sub KillGLWindow(frm As Form)
    Dim lRet As Long

    ' If the caption property of the form is set to nothing then restore the display mode and show the cursor
    If (frm.Caption = "") Then
      RestoreDisplayMode
      ShowCursor True
    End If

    ' If we have a valid handle to an opengl rendering context
    If hRC Then
        ' Attempt to detatch the rendering context from the hdc associated with our window
        lRet = wglMakeCurrent(0, 0)
        If (lRet = 0) Then
            MsgBox "Release Of DC And RC Failed."
        End If
    
        ' Delete the rendering context
        lRet = wglDeleteContext(hRC)
        If (lRet = 0) Then
            MsgBox "Release Rendering Context Failed."
        End If
        
        hRC = 0
    End If
End Sub

Public Sub ReSizeGLScene(ByVal Width As GLsizei, ByVal Height As GLsizei)
    ' Prevent an error if the user attempts to resize the form too small
    If Height = 0 Then
        Height = 1
    End If
      
    ' Set the viewing area
    glViewport 0, 0, Width, Height
      
    ' Change to the projection matrix
    glMatrixMode mmProjection
    ' Reset the projection matrix
    glLoadIdentity
    
    ' Set our viewing frustrum
    gluPerspective 45#, Width / Height, 0.1, 100#
    
    ' Change to the model view
    glMatrixMode mmModelView
    ' Reset the model view
    glLoadIdentity
End Sub

Private Sub RestoreDisplayMode()
    Dim dmDisplaySettings As DEVMODE
    Dim lRet As Long

    ' Fill the structure with the info we saved earlier in savedisplaysettings
    dmDisplaySettings.dmPelsWidth = giOldWidth
    dmDisplaySettings.dmPelsHeight = giOldHeight
    dmDisplaySettings.dmBitsPerPel = giOldNumBits
    dmDisplaySettings.dmDisplayFrequency = giOldVRefresh
    dmDisplaySettings.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL Or DM_DISPLAYFREQUENCY
    dmDisplaySettings.dmSize = Len(dmDisplaySettings)
    
    ' Attempt to restore the display mode
    lRet = ChangeDisplaySettings(dmDisplaySettings, 0)
    If (lRet <> DISP_CHANGE_SUCCESSFUL) Then
        MsgBox "Display restore failed."
    End If
End Sub

Public Sub SaveDisplaySettings()
    Dim hRet As Long
    Dim dm As DEVMODE

    ' Create an information context for the current display
    hRet = CreateIC("DISPLAY", "", "", dm)
    
    ' Save information
    giOldWidth = GetDeviceCaps(hRet, HORZRES)
    giOldHeight = GetDeviceCaps(hRet, VERTRES)
    giOldNumBits = GetDeviceCaps(hRet, BITSPIXEL)
    giOldVRefresh = GetDeviceCaps(hRet, VREFRESH)
    
    ' Delete the IC
    DeleteDC hRet
End Sub

Public Function SetFullscreenMode(iWidth As Integer, iHeight As Integer, iNumBits As Integer) As Boolean
    Dim bSuccess As Boolean
    Dim dmDisplaySettings As DEVMODE
    Dim lRet As Long

    bSuccess = False
  
    ' Fill structure with info on display mode
    dmDisplaySettings.dmPelsWidth = iWidth
    dmDisplaySettings.dmPelsHeight = iHeight
    dmDisplaySettings.dmBitsPerPel = iNumBits
    dmDisplaySettings.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT Or DM_BITSPERPEL
    dmDisplaySettings.dmSize = Len(dmDisplaySettings)
    
    ' Change the display mode
    lRet = ChangeDisplaySettings(dmDisplaySettings, CDS_FULLSCREEN)
    If (lRet <> DISP_CHANGE_SUCCESSFUL) Then
        bSuccess = False
    Else
        bSuccess = True
    End If
    
    ' Return result of attempt to change the display mode
    SetFullscreenMode = bSuccess
End Function

