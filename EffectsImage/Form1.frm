VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8175
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   ScaleHeight     =   8175
   ScaleWidth      =   9945
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go!"
      Default         =   -1  'True
      Height          =   450
      Left            =   5100
      TabIndex        =   2
      Top             =   7575
      Width           =   1800
   End
   Begin VB.PictureBox picDest 
      AutoSize        =   -1  'True
      Height          =   7290
      Left            =   5025
      ScaleHeight     =   482
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   318
      TabIndex        =   1
      Top             =   75
      Width           =   4830
   End
   Begin VB.PictureBox picSource 
      AutoSize        =   -1  'True
      Height          =   7260
      Left            =   75
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   75
      Width           =   4860
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////////////////////////////////////
'Looking through semi-transparent glass.
'Based on GetDIBits and SetDIBits tutorial on some website. (Thanks)
'I don't remember the address of the website.
'Comments? Suggestions?
'Feel free to e-mail me at < minsin999@hotmail.com >
'////////////////////////////////////////////////////////////////

Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&

Private Const red As Integer = 3
Private Const green As Integer = 2
Private Const blue As Integer = 1

Private Type BITMAPINFOHEADER '40 bytes
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

Private Type BITMAPINFO
      bmiHeader As BITMAPINFOHEADER
     'bmiColors As RGBQUAD
End Type

Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long

Private Sub cmdGo_Click()
      Dim x As Integer, y As Integer
      Dim SourceWidth As Integer, SourceHeight As Integer
      Dim SourcePixels() As Byte
      Dim ModifiedPixels() As Byte

      Dim BmpInfo As BITMAPINFO
      
      SourceWidth = picSource.ScaleWidth
      SourceHeight = picSource.ScaleHeight
      
      ' Prepare the bitmap description.
      With BmpInfo.bmiHeader
            .biSize = 40
            .biWidth = picSource.ScaleWidth
            ' Use negative height to scan top-down.
            .biHeight = -picSource.ScaleHeight
            .biPlanes = 1
            .biBitCount = 32
            .biCompression = BI_RGB
            .biSizeImage = 0
      End With
      
      ' Load the bitmap's data.
      ReDim SourcePixels(1 To 4, 1 To picSource.ScaleWidth, 1 To picSource.ScaleHeight)
      ReDim ModifiedPixels(1 To 4, 1 To picSource.ScaleWidth, 1 To picSource.ScaleHeight)
      
      GetDIBits picSource.hdc, picSource.Image, 0, picSource.ScaleHeight, _
      SourcePixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS
      
      Dim RedVal As Integer, GreenVal As Integer, BlueVal As Integer
      Dim rx As Integer, ry As Integer
      
      For y = 1 To picSource.ScaleHeight
            For x = 1 To picSource.ScaleWidth
            
                  '//////////////////////////////////////////////
                  '*** Looking through semi-transparent glass ***
                  'Change '20' to the appropriate value
                  rx = Sin(x) * (Rnd * 20) + x
                  ry = Cos(y) * (Rnd * 20) + y
                  '//////////////////////////////////////////////
                  
                  '//////////////////////////////////////////////
                  '*** Diffuse ***
                  'Change '10' to the appropriate value
                  'rx = x + (Rnd * 10)
                  'ry = y + (Rnd * 10)
                  '//////////////////////////////////////////////
                  
                  If rx < 1 Then rx = 1
                  If rx >= SourceWidth Then rx = SourceWidth - 1
                  If ry < 1 Then ry = 1
                  If ry >= SourceHeight Then ry = SourceHeight - 1
                  
                  RedVal = SourcePixels(red, rx, ry)
                  GreenVal = SourcePixels(green, rx, ry)
                  BlueVal = SourcePixels(blue, rx, ry)
                  
                  ModifiedPixels(red, x, y) = RedVal
                  ModifiedPixels(blue, x, y) = BlueVal
                  ModifiedPixels(green, x, y) = GreenVal
            Next x
      Next y
      
      ' Display the result.
      SetDIBits picDest.hdc, picDest.Image, 0, picSource.ScaleHeight, _
      ModifiedPixels(1, 1, 1), BmpInfo, DIB_RGB_COLORS
      
      picDest.Picture = picDest.Image
      MsgBox "Done"
End Sub
