Attribute VB_Name = "modScreenShots"
#Const FinalMode = 1
Option Explicit
'''''''
'Types
'''''''
Private Type Guid
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Sub keybd_event Lib "user32" _
        (ByVal bVk As Byte, ByVal bScan As Byte, _
        ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
        
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function OleCreatePictureIndirect Lib "olepro32" (PicDesc As PicBmp, RefIID As Guid, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
        
        
Public Function GetWindowScreenshot(WndHandle As Long, SavePath As String, Optional BringFront As Integer = 1) As String
'
' Function to create screeenshot of specified window and store at specified path
'
    On Error GoTo ErrorHandler

    Dim hdcSrc As Long
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim WidthSrc As Long
    Dim HeightSrc As Long
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As Guid
    Dim rc As RECT
    Dim pictr As PictureBox
    
    'Bring window on top of all windows if specified
    If BringFront = 1 Then BringWindowToTop WndHandle
    
    'Get Window Size
    GetWindowRect WndHandle, rc
    WidthSrc = rc.Right - rc.Left
    HeightSrc = rc.Bottom - rc.Top
    
    'Get Window  device context
    hdcSrc = GetWindowDC(WndHandle)
    
    'create a memory device context
    hDCMemory = CreateCompatibleDC(hdcSrc)
    
    'create a bitmap compatible with window hdc
    hBmp = CreateCompatibleBitmap(hdcSrc, WidthSrc, HeightSrc)
    
    'copy newly created bitmap into memory device context
    hBmpPrev = SelectObject(hDCMemory, hBmp)
    
    'copy window window hdc to memory hdc
    Call BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, _
                hdcSrc, 0, 0, vbSrcCopy)
      
    'Get Bmp from memory Dc
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    
    'release the created objects and free memory
    Call DeleteDC(hDCMemory)
    Call ReleaseDC(WndHandle, hdcSrc)
    
    'fill in OLE IDispatch Interface ID
    With IID_IDispatch
       .Data1 = &H20400
       .Data4(0) = &HC0
       .Data4(7) = &H46
     End With
    
    'fill Pic with necessary parts
    With Pic
       .Size = Len(Pic)         'Length of structure
       .Type = vbPicTypeBitmap  'Type of Picture (bitmap)
       .hBmp = hBmp             'Handle to bitmap
       .hPal = 0&               'Handle to palette (may be null)
     End With
    
    'create OLE Picture object
    Call OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    
    'return the new Picture object
    SavePicture IPic, SavePath
    GetWindowScreenshot = ""
    Exit Function
    
ErrorHandler:
    GetWindowScreenshot = "GetWindowScreenshot got error code " & CStr(Err.Number) & ": " & CStr(Err.Description)

        
End Function

        
Public Sub GetScreenshot(ByRef frmS As Form, pictureName As String)
  #If FinalMode = 1 Then
  On Error GoTo gotErr
  #End If
  Dim strCompleteName As String
  Dim sRes As String
  Dim hwnd As Long
  strCompleteName = App.Path & "\" & pictureName
  hwnd = GetDesktopWindow()
  sRes = GetWindowScreenshot(hwnd, strCompleteName, 0)
  If sRes <> "" Then
    LogOnFile "errors.txt", sRes
  End If
  Exit Sub
  
  frmMenu.Hide
  frmAdvanced.Hide
  frmBackpacks.Hide
  frmBigText.Hide
  frmCavebot.Hide
  frmCheats.Hide
  frmCondEvents.Hide
  frmEvents.Hide
  frmHardcoreCheats.Hide
  frmHotkeys.Hide
  frmMagebomb.Hide
  frmHPmana.Hide
  frmMain.Hide
  frmMapReader.Hide
  frmTrainer.Hide
  frmTrueMap.Hide
  frmWarbot.Hide
  frmRunemaker.Hide
  frmConfirm.Hide
  frmS.Show
  DoEvents
  Clipboard.Clear
  keybd_event vbKeySnapshot, 0, 0, 0
  DoEvents
  keybd_event vbKeySnapshot, 0, &H2, 0
  DoEvents
  If IsNull(Clipboard.GetData(vbCFBitmap)) Then
    LogOnFile "errors.txt", "Clipboard content was null after screenshot!"
    Exit Sub
  End If
  frmS.picScreen.Picture = Clipboard.GetData(vbCFBitmap)
  DoEvents
  strCompleteName = App.Path & "\" & pictureName
  SavePicture frmS.picScreen.Picture, strCompleteName
  'SavePicture frmS.picScreen.Image, strCompleteName
  frmS.Hide
  Exit Sub
gotErr:
  LogOnFile "errors.txt", "Got error " & CStr(Err.Number) & " : " & Err.Description & " . This error happened while trying to save log at " & strCompleteName
End Sub

