VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Type BITMAP           '14 bytes
  bmType As Long
  bmWidth As Long
  bmHeight As Long
  bmWidthBytes As Long
  bmPlanes As Long
  bmBitsPixel As Long
  bmBits As Long
End Type

Private Type BITMAPINFOHEADER  '40 bytes
  biSize As Long
  biWidth As Long
  biHeight As Long
  biPlanes As Integer
  biBitCount As Integer
  biCompression As Long
  biSizeImage As Long
  biXPelsPerMeter As Long
  biYPelsPerMeter As Long
  biClrUsed As Long
  biClrImportant As Long
End Type

Private Type RGBQUAD           '4 bytes
  rgbBlue As Byte
  rgbGreen As Byte
  rgbRed As Byte
  rgbReserved As Byte
End Type

Public Enum ENUM_TEXTURE_FILETYPE
    FILETYPE_BMP = 0
    FILETYPE_TGA = 1
    FILETYPE_RAW = 2
End Enum

Public Enum ENUM_TEXTURE_FILTER
    FILTER_NEAREST = 0
    FILTER_LINEAR = 1
    FILTER_MIPMAPPED = 2
End Enum

Private m_Textures(3) As Long
Private m_Filter As Integer
Private m_Width As Long, m_Height As Long


Public Sub useTexture()
    glBindTexture glTexture2D, m_Textures(m_Filter)
End Sub

Public Sub loadTexture(Filename As String, FileType As ENUM_TEXTURE_FILETYPE, Optional Width As Long, Optional Height As Long)
    Select Case FileType
        ' BMP
        Case 0:
            loadBitmap Filename
        ' TGA
        Case 1:
            loadTGA Filename
        ' RAW
        Case 2:
            If IsMissing(Width) Or IsMissing(Height) Then Exit Sub
            loadRAW Filename, Width, Height
    End Select
    
    setFilter FILTER_NEAREST
End Sub

Private Sub loadRAW(Filename As String, Width As Long, Height As Long)
' Assumes RGB top to bottom 24BPP format. No header info
    Dim ff As Integer
    Dim b() As Byte
    Dim c As Long
    Dim t As Byte
    Dim i As Long, a As Long
    
    ff = FreeFile()
    ReDim b((Width * Height) * 4)
    
    Open Filename For Binary As ff
        Get ff, , b
    Close ff
    
    makeTexturesFromByteArray b, Width, Height, True
End Sub

Private Sub loadBitmap(Filename As String)
    Dim bmFile As BITMAPFILEHEADER
    Dim bmInfo As BITMAPINFOHEADER
    Dim bmRGB() As RGBQUAD
    Dim iFile As Integer
    Dim lImageSize As Long
    Dim iPixelSize As Integer
    Dim baImageData() As Byte
  
    iFile = FreeFile
  
    Open Filename For Binary As iFile
        Get #iFile, , bmFile
        Get #iFile, , bmInfo
        
        If (bmInfo.biBitCount < 24) Then
            ReDim bmRGB(bmInfo.biClrUsed)
          
            Get #iFile, , bmRGB
        End If
        
        iPixelSize = bmInfo.biBitCount / 8
        
        lImageSize = bmInfo.biWidth * bmInfo.biHeight * iPixelSize
        
        ReDim baImageData(lImageSize)
        
        Get #iFile, , baImageData
    Close #iFile
  
    makeTexturesFromByteArray baImageData, bmInfo.biWidth, bmInfo.biHeight, False
End Sub

Private Sub loadTGA(Filename As String)
    Dim ty(2) As Byte
    Dim inf(5) As Byte
    Dim imageData() As Byte
    Dim imageWidth As Long, imageHeight As Long
    Dim imageBits As Integer, s As Long
    Dim ff As Integer
    
    ff = FreeFile
    
    Open Filename For Binary As ff
        Get ff, , ty
        Get ff, 13, inf
        
        If Not ty(1) = 0 And Not ty(2) = 2 Then MsgBox "Bad TGA Image Type": Exit Sub
        
        imageWidth = CLng(inf(0) + inf(1) * 256)
        imageHeight = CLng(inf(2) + inf(3) * 256)
        imageBits = CInt(inf(4))
        
        s = imageWidth * imageHeight
        
        If Not imageBits = 32 And Not imageBits = 24 Then MsgBox "Unsupported bitdepth": Exit Sub
        
        ReDim imageData(s * (imageBits \ 8))
        imageData = getData(ff, s, imageBits)
    Close ff
    
    makeTexturesFromByteArray imageData, imageWidth, imageHeight, IIf(imageBits = 32, True, False)
End Sub

Private Function getRGBA(FileNumber As Integer, s As Long) As Byte()
    Dim b() As Byte
    Dim t As Byte
    Dim i As Long
    
    ReDim b((s - 1) * 4)
    
    Get FileNumber, , b
    
    getRGBA = b
End Function

Private Function getRGB(FileNumber As Integer, s As Long) As Byte()
    Dim b() As Byte
    Dim t As Byte
    Dim i As Long
    
    ReDim b((s - 1) * 3)
    
    Get FileNumber, , b
    
    getRGB = b
End Function

Private Function getData(FileNumber As Integer, s As Long, iBits As Integer) As Byte()
    If iBits = 32 Then
        getData = getRGBA(FileNumber, s)
    Else
        getData = getRGB(FileNumber, s)
    End If
End Function

Private Sub makeTexturesFromByteArray(ByRef b() As Byte, ByVal w As Long, ByVal h As Long, hasAlpha As Boolean)
    glGenTextures 3, m_Textures(0)

    glBindTexture glTexture2D, m_Textures(0)
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_NEAREST
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_NEAREST
    glTexImage2D glTexture2D, 0, IIf(hasAlpha = True, 4, 3), w, h, _
        0, tiBGRExt, GL_UNSIGNED_BYTE, b(0)
  
    glBindTexture glTexture2D, m_Textures(1)
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR
    glTexImage2D glTexture2D, 0, IIf(hasAlpha = True, 4, 3), w, h, _
        0, tiBGRExt, GL_UNSIGNED_BYTE, b(0)
  
    glBindTexture glTexture2D, m_Textures(2)
    glTexParameteri glTexture2D, tpnTextureMagFilter, GL_LINEAR
    glTexParameteri glTexture2D, tpnTextureMinFilter, GL_LINEAR_MIPMAP_NEAREST
    gluBuild2DMipmaps glTexture2D, IIf(hasAlpha = True, 4, 3), w, h, tiBGRExt, _
        GL_UNSIGNED_BYTE, ByVal VarPtr(b(0))
        
    m_Width = w
    m_Height = h
End Sub

Public Sub setFilter(Filter As ENUM_TEXTURE_FILTER)
    m_Filter = Filter
End Sub

Public Function getFilter() As ENUM_TEXTURE_FILTER
    getFilter = m_Filter
End Function

Public Sub unloadTexture()
    glDeleteTextures 3, m_Textures(0)
End Sub

Private Sub Class_Terminate()
    unloadTexture
End Sub
