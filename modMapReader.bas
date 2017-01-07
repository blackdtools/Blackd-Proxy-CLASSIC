Attribute VB_Name = "modMapReader"
#Const FinalMode = 1
#Const ShowMapLoading = 0
Option Explicit
Public Type TypeBMSquare
  Color As Long
  walkable As Long
End Type
Public Type TypeBMnameInfo
  x As Long
  y As Long
  z As Long
  Color As Long
End Type
Const ColorCaveWall = &H72
Const ColorCaveWalkable = &H79
Const ColorFloorChange = &HD2
Const ColorRedWall = &HBA
Const ColorGrayWalkable = &H81
Const ColorGrayWall = &H56
Const ColorDesertWalkable = &HCF
Const ColorGreenWall = &HC
Const ColorWater = &H28
Const ColorGreenWalkable = &H18
Const ColorSnowWalkable = &HB3
Const ColorSwampWall = &H1E
' Tibia 11 water color
Const ColorWater2 = &H33

' Before Tibia 11.00, we read 131072 bytes.
' first 65536 bytes = map color bytes
' next 65536 bytes = path info bytes
Const cteBytesPerMap = 131072
Const cteBytesPerMap1 = 131071
Public Const MINX = -63
Public Const MAXX = 64
Public Const MINY = -63
Public Const MAXY = 64
Public LoadingAmap As Boolean
Public currMapX As Long
Public currMapY As Long
Public currMapZ As Long

Public BigMapNamesX As Scripting.Dictionary
Public BigMapNamesY As Scripting.Dictionary
Public BigMapNamesZ As Scripting.Dictionary
Public BigMapNamesC As Scripting.Dictionary

Public MapIDTranslator As Scripting.Dictionary
Public TheVeryBigMap() As Byte

Public TibiaPath As String


' TRANSLATOR strID -> lngID
Public Sub AddMapTranslation(ByRef strID As String, ByRef lngID As Long)
  MapIDTranslator.item(strID) = lngID
End Sub

Public Sub RemoveAllMapTranslation()
  MapIDTranslator.RemoveAll
End Sub

Public Function GetMapTranslation(ByRef strID As String) As Long
  ' get the name from an ID
  Dim aRes As Long
  Dim res As Boolean
  If MapIDTranslator.Exists(strID) = True Then
    GetMapTranslation = MapIDTranslator.item(strID)
  Else
    GetMapTranslation = -1
  End If
End Function









Public Function ExistBigMapName(str As String) As Boolean
  ' already added?
  ExistBigMapName = BigMapNamesX.Exists(str)
End Function

Public Sub AddBigMapName(str As String, x As Long, y As Long, z As Long, c As Long)
  ' add item to dictionary
  BigMapNamesX.item(str) = x
  BigMapNamesY.item(str) = y
  BigMapNamesZ.item(str) = z
  BigMapNamesC.item(str) = c
End Sub
Public Sub RemoveBigMapName(str As String)
  ' remove item from dictionary
  If BigMapNamesX.Exists(str) = True Then
    BigMapNamesX.Remove (str)
    BigMapNamesY.Remove (str)
    BigMapNamesZ.Remove (str)
    BigMapNamesC.Remove (str)
  End If
End Sub
Public Function GetBigMapNameInfo(str As String) As TypeBMnameInfo
  ' get x,y,z,c from an name
  Dim res As TypeBMnameInfo
  If BigMapNamesX.Exists(str) = True Then
    res.x = BigMapNamesX.item(str)
    res.y = BigMapNamesY.item(str)
    res.z = BigMapNamesZ.item(str)
    res.Color = BigMapNamesC.item(str)
  Else
    res.x = 0
    res.y = 0
    res.z = 0
    res.Color = vbBlack
  End If
  GetBigMapNameInfo = res
End Function

Public Function GetOldMapName(x As Long, y As Long, z As Long) As String
  Dim xp As String
  Dim yp As String
  Dim zp As String
  Dim res As String
  xp = CStr(x \ 256)
  yp = CStr(y \ 256)
  If z < 10 Then
    zp = "0" & CStr(z)
  Else
   zp = CStr(z)
  End If
  res = xp & yp & zp
  GetOldMapName = res
End Function
Public Function CalcMapID(x As Long, y As Long, z As Long) As Long
  Dim xp As String
  Dim yp As String
  Dim zp As String
  Dim res As String
  xp = CStr(x \ 256)
  yp = CStr(y \ 256)
  If z < 10 Then
    zp = "0" & CStr(z)
  Else
   zp = CStr(z)
  End If
  res = xp & yp & zp
  CalcMapID = GetMapTranslation(res)
End Function

Public Sub GetBigMapSquareB(ByRef res As Boolean, x As Long, y As Long, z As Long)
  Dim bx As Long
  Dim by As Long
  Dim xp As String
  Dim yp As String
  Dim zp As String
  Dim xdif As Long
  Dim ydif As Long
  Dim squares As Long
  Dim mapbyte As Byte
  Dim theMapID As Long
  bx = x \ 256
  by = y \ 256
  theMapID = CalcMapID(x, y, z)
  xdif = x - (bx * 256)
  ydif = y - (by * 256)
  squares = (xdif * 256) + ydif '+0 instead +1 now
  If theMapID >= 0 Then
    mapbyte = TheVeryBigMap(squares, theMapID)
      Select Case mapbyte
          Case &H0
            res = False
          Case ColorCaveWall
            res = False
          Case ColorCaveWalkable
            res = True
          Case ColorRedWall
            res = False
          Case ColorGrayWalkable
            res = True
          Case ColorFloorChange
            res = False
          Case ColorGrayWall
            res = False
          Case ColorDesertWalkable
            res = True
          Case ColorGreenWall
            res = False
          Case ColorWater
            res = False
          Case ColorWater2
            res = False
          Case ColorGreenWalkable
            res = True
          Case ColorSnowWalkable
            res = True
          Case ColorSwampWall
            res = False
          Case Else
            res = False
        End Select
  Else ' map info not loaded for this point
    res = False
  End If
End Sub

Public Sub GetBigMapSquare(ByRef res As TypeBMSquare, x As Long, y As Long, z As Long)
  Dim bx As Long
  Dim by As Long
  Dim xp As String
  Dim yp As String
  Dim zp As String
  Dim xdif As Long
  Dim ydif As Long
  Dim squares As Long
  Dim mapbyte As Byte
  Dim theMapID As Long
  bx = x \ 256
  by = y \ 256
  theMapID = CalcMapID(x, y, z)
  xdif = x - (bx * 256)
  ydif = y - (by * 256)
  squares = (xdif * 256) + ydif '+0 instead +1 now
  If theMapID >= 0 Then
    mapbyte = TheVeryBigMap(squares, theMapID)
      Select Case mapbyte
          Case &H0
            res.Color = &H0&
            res.walkable = False
          Case ColorCaveWall
            res.Color = &H4080&
            res.walkable = False
          Case ColorCaveWalkable
            res.Color = &H80FF&
            res.walkable = True
          Case ColorRedWall
            res.Color = &HFF&
            res.walkable = False
          Case ColorGrayWalkable
            res.Color = &HC0C0C0
            res.walkable = True
          Case ColorFloorChange
            res.Color = &HFFFF&
            res.walkable = False
          Case ColorGrayWall
            res.Color = &H808080
            res.walkable = False
          Case ColorDesertWalkable
            res.Color = &HC0FFFF
            res.walkable = True
          Case ColorGreenWall
            res.Color = &H8000&
            res.walkable = False
          Case ColorWater
            res.Color = &HFF0000
            res.walkable = False
          Case ColorWater2
            res.Color = &HFF0000
            res.walkable = False
          Case ColorGreenWalkable
            res.Color = &HFF00&
            res.walkable = True
          Case ColorSnowWalkable
            res.Color = &HFFFFC0
            res.walkable = True
          Case ColorSwampWall
            res.Color = &H80FF80
            res.walkable = False
          Case Else
            res.Color = &HFFFFFF
            res.walkable = False
        End Select
  Else
    res.Color = &H0&
    res.walkable = False
  End If
End Sub

Public Sub LoadBigMap(ByRef map() As TypeBMSquare)
  Dim x As Long
  Dim y As Long
  Dim z As Long
  Dim firstX As Long
  Dim firstY As Long
  Dim endX As Long
  Dim endY As Long
  Dim Px As Long
  Dim Py As Long
  Dim res As TypeBMSquare
  Dim bx As Long
  Dim by As Long
  Dim xp As String
  Dim yp As String
  Dim zp As String
  Dim xdif As Long
  Dim ydif As Long
  Dim squares As Long
  Dim mapbyte As Byte
  Dim theMapID As Long
  #If ShowMapLoading = 1 Then
    If LoadingAmap = False Then
      LoadingAmap = True
    Else
      Exit Sub
    End If
  #End If
  firstX = currMapX + MINX
  endX = currMapX + MAXX
  firstY = currMapY + MINY
  endY = currMapY + MAXY
  Px = MINX
  For x = firstX To endX
    Py = MINY
    For y = firstY To endY
        z = currMapZ
        bx = x \ 256
        by = y \ 256
        theMapID = CalcMapID(x, y, z)
        xdif = x - (bx * 256)
        ydif = y - (by * 256)
        squares = (xdif * 256) + ydif
        If theMapID >= 0 Then
            mapbyte = TheVeryBigMap(squares, theMapID)
            Select Case mapbyte
            Case &H0
                res.Color = &H0&
                res.walkable = False
            Case ColorCaveWall
                res.Color = &H4080&
                res.walkable = False
            Case ColorCaveWalkable
                res.Color = &H80FF&
                res.walkable = True
            Case ColorRedWall
                res.Color = &HFF&
                res.walkable = False
            Case ColorGrayWalkable
                res.Color = &HC0C0C0
                res.walkable = True
            Case ColorFloorChange
                res.Color = &HFFFF&
                res.walkable = False
            Case ColorGrayWall
                res.Color = &H808080
                res.walkable = False
            Case ColorDesertWalkable
                res.Color = &HC0FFFF
                res.walkable = True
            Case ColorGreenWall
                res.Color = &H8000&
                res.walkable = False
            Case ColorWater
                res.Color = &HFF0000
                res.walkable = False
            Case ColorWater2
                res.Color = &HFF0000
                res.walkable = False
            Case ColorGreenWalkable
                res.Color = &HFF00&
                res.walkable = True
            Case ColorSnowWalkable
                res.Color = &HFFFFC0
                res.walkable = True
            Case ColorSwampWall
                res.Color = &H80FF80
                res.walkable = False
            Case Else
                res.Color = &HFFFFFF
                res.walkable = False
            End Select
        Else
            res.Color = &H0&
            res.walkable = False
        End If
        map(Px, Py).Color = res.Color
        map(Px, Py).walkable = res.walkable
        Py = Py + 1
    Next y
    Px = Px + 1
    #If ShowMapLoading = 1 Then
        DoEvents
    #End If
  Next x
  #If ShowMapLoading = 1 Then
    LoadingAmap = False
  #End If
End Sub

Public Sub DrawMap()
  Dim x As Long
  Dim y As Long
  Dim ix As Long
  Dim iy As Long
  Dim map(-127 To 128, -127 To 128) As TypeBMSquare
  Dim mfilename As String
  #If ShowMapLoading = 1 Then
  frmMapReader.cmdUpdate.enabled = False
  frmMapReader.cmdNorth.enabled = False
  frmMapReader.cmdSouth.enabled = False
  frmMapReader.cmdWest.enabled = False
  frmMapReader.cmdEast.enabled = False
  frmMapReader.cmdFloorUp.enabled = False
  frmMapReader.cmdFloorDown.enabled = False
  frmMapReader.cmdCenter2.enabled = False
  frmMapReader.lblName.Caption = "Loading, please wait ..."
  DoEvents
  #End If
  LoadBigMap map
  'frmMapReader.picMap.AutoRedraw = True
  For iy = MINX To MAXX
    For ix = MINY To MAXY
      x = ix - MINX
      y = iy - MINY
      frmMapReader.picMap.Line (x * 2, y * 2)-((x * 2) + 1, (y * 2) + 1), map(ix, iy).Color, BF
    Next ix
  Next iy
  #If ShowMapLoading = 1 Then
  frmMapReader.lblName.Caption = "Showing map around " & vbCrLf & CStr(currMapX) & " , " & CStr(currMapY) & " , " & CStr(currMapZ)
  frmMapReader.cmdUpdate.enabled = True
  frmMapReader.cmdNorth.enabled = True
  frmMapReader.cmdSouth.enabled = True
  frmMapReader.cmdWest.enabled = True
  frmMapReader.cmdEast.enabled = True
  frmMapReader.cmdFloorUp.enabled = True
  frmMapReader.cmdFloorDown.enabled = True
  frmMapReader.cmdCenter2.enabled = True
  #End If
  frmMapReader.RedrawAllMarks
End Sub
Public Sub DrawLine(x1 As Long, y1 As Long, z1 As Long, x2 As Long, y2 As Long, z2 As Long, Color As Long)
  Dim px1 As Long
  Dim py1 As Long
  Dim xBase1 As Long
  Dim yBase1 As Long
  Dim px2 As Long
  Dim py2 As Long
  Dim xBase2 As Long
  Dim yBase2 As Long
  xBase1 = x1 - currMapX
  yBase1 = y1 - currMapY
  xBase2 = x2 - currMapX
  yBase2 = y2 - currMapY
  If (xBase1 >= MINX) And (xBase1 <= MAXX) And (yBase1 >= MINY) And (yBase1 <= MAXY) And (z1 = currMapZ) And _
     (xBase2 >= MINX) And (xBase2 <= MAXX) And (yBase2 >= MINY) And (yBase2 <= MAXY) And (z2 = currMapZ) Then
    px1 = 1 + ((xBase1 - MINX) * 2)
    py1 = 1 + ((yBase1 - MINY) * 2)
    px2 = 1 + ((xBase2 - MINX) * 2)
    py2 = 1 + ((yBase2 - MINY) * 2)
    frmMapReader.picMap.AutoRedraw = False
    frmMapReader.picMap.DrawWidth = 2
    frmMapReader.picMap.Line (px1, py1)-(px2, py2), Color
    frmMapReader.picMap.DrawWidth = 1
    frmMapReader.picMap.AutoRedraw = True
  End If
End Sub
Public Sub DrawXYZPixel(x As Long, y As Long, z As Long, Color As Long)
  Dim Px As Long
  Dim Py As Long
  Dim xBase As Long
  Dim yBase As Long
  xBase = x - currMapX
  yBase = y - currMapY
  If (xBase >= MINX) And (xBase <= MAXX) And (yBase >= MINY) And (yBase <= MAXY) And (z = currMapZ) Then
    Px = 1 + ((xBase - MINX) * 2)
    Py = 1 + ((yBase - MINY) * 2)
    frmMapReader.picMap.AutoRedraw = False
    frmMapReader.picMap.FillStyle = 0
    frmMapReader.picMap.FillColor = Color
    frmMapReader.picMap.Circle (Px, Py), 1, Color
    frmMapReader.picMap.AutoRedraw = True
  End If
End Sub
Public Sub DrawXYZPixel2(x As Long, y As Long, z As Long, Color As Long)
  Dim Px As Long
  Dim Py As Long
  Dim xBase As Long
  Dim yBase As Long
  xBase = x - currMapX
  yBase = y - currMapY
  If (xBase >= MINX) And (xBase <= MAXX) And (yBase >= MINY) And (yBase <= MAXY) And (z = currMapZ) Then
    Px = 1 + ((xBase - MINX) * 2)
    Py = 1 + ((yBase - MINY) * 2)
    frmMapReader.picMap.AutoRedraw = False
    frmMapReader.picMap.FillStyle = 0
    frmMapReader.picMap.FillColor = Color
    frmMapReader.picMap.Circle (Px, Py), 1, Color
    frmMapReader.picMap.FillStyle = 1
    frmMapReader.picMap.Circle (Px, Py), 5, Color
    frmMapReader.picMap.AutoRedraw = True
  End If
End Sub
Public Sub DrawXYZnMap(x As Long, y As Long, z As Long, Color As Long)
  Dim Px As Long
  Dim Py As Long
  Dim xBase As Long
  Dim yBase As Long
  xBase = x - currMapX
  yBase = y - currMapY
  If (xBase >= MINX) And (xBase <= MAXX) And (yBase >= MINY) And (yBase <= MAXY) And (z = currMapZ) Then
    Px = 1 + ((xBase - MINX) * 2)
    Py = 1 + ((yBase - MINY) * 2)
    frmMapReader.picMap.AutoRedraw = False
    frmMapReader.picMap.FillStyle = 0
    frmMapReader.picMap.FillColor = Color
    frmMapReader.picMap.Circle (Px, Py), 4, Color
    frmMapReader.picMap.FillStyle = 1
    frmMapReader.picMap.Circle (Px, Py), 6, Color
    frmMapReader.picMap.Circle (Px, Py), 12, Color
    frmMapReader.picMap.Circle (Px, Py), 30, Color
    frmMapReader.picMap.AutoRedraw = True
  End If
End Sub

Public Sub waitThisMs(thisms As Long)
Dim tl As Long
Dim ct As Long
tl = GetTickCount() + thisms
ct = GetTickCount
While ct < tl
  ct = GetTickCount()
Wend
End Sub
Public Sub waitThisMs2(thisms As Long)
Dim tl As Long
Dim ct As Long
tl = GetTickCount() + thisms
ct = GetTickCount
While ct < tl
  ct = GetTickCount()
  DoEvents
Wend
End Sub

Public Function ReadHardiskMapsQ() As Long
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  Dim res As Long
  Dim Ammount As Long
  Dim readed As Long
  Dim progr As Double
  Dim fs As Scripting.FileSystemObject
  Dim f As Scripting.Folder
  Dim f1 As Scripting.File
  Dim fn As Integer
  Dim mapFile As String
  Dim strID As String
  Dim lngID As Long
  Dim thersize As Long
  Dim bigReadB(0 To cteBytesPerMap1) As Byte
  Dim w As Long
  Dim h As Long
  Dim bpp As Long
  Dim resPNG1 As Boolean
  Dim resPNG2 As Boolean
  Dim pathColor As String
  Dim pathWaypoints As String
  Dim posX As Long
  Dim posY As Long
  Dim posz As Long
  Dim part4 As String
  Dim nameParts() As String
  Dim loadedOK As Long
  RemoveAllMapTranslation
  Ammount = 0
  readed = 0
  Set fs = New Scripting.FileSystemObject
  If fs.FolderExists(TibiaPath) = False Then
    ReadHardiskMapsQ = -1
    Exit Function
  End If
  Set f = fs.GetFolder(TibiaPath)
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "png" Then
      If Len(f1.name) > 14 Then
         If (LCase(Left(f1.name, 14))) = "minimap_color_" Then
            Ammount = Ammount + 1
         End If
      End If
    End If
  Next
  If Ammount = 0 Then
    ReadHardiskMapsQ = -1
    Exit Function
  End If
  thersize = Ammount - 1
  loadedOK = 0
  ReDim TheVeryBigMap(0 To cteBytesPerMap1, 0 To thersize)
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "png" Then
      If Len(f1.name) > 13 Then
        If (LCase(Left(f1.name, 13))) = "minimap_color" Then
            lngID = readed
            readed = readed + 1
            If LoadWasCompleted = False Then
              progr = (readed / Ammount) * 100
              frmLoading.NotifyLoadProgress progr, "Loading map file " & CStr(readed) & " / " & CStr(Ammount)
            End If
            pathColor = TibiaPath & "\" & f1.name
            pathWaypoints = TibiaPath & "\Minimap_WaypointCost_" & Right$(f1.name, Len(f1.name) - 14)
            nameParts = Split(f1.name, "_", , vbTextCompare)
            posX = CLng(nameParts(2))
            posY = CLng(nameParts(3))
            part4 = nameParts(4)
            nameParts = Split(part4, ".", , vbTextCompare)
            posz = CLng(nameParts(0))
            resPNG1 = LoadPNGIntoByteArray(pathColor, bigReadB, w, h, bpp, 0, 90)
            resPNG2 = LoadPNGIntoByteArray(pathColor, bigReadB, w, h, bpp, cteBytesPerMap / 2, 90)
            If (resPNG1 And resPNG2) Then
                strID = GetOldMapName(posX, posY, posz)
                AddMapTranslation strID, lngID
                RtlMoveMemory TheVeryBigMap(0, lngID), bigReadB(0), cteBytesPerMap
                loadedOK = loadedOK + 1
            End If
        End If
      End If
    End If
  Next
  'Debug.Print "Loaded Tibia 11 maps! Total found: " & CStr(loadedOK)
  ReadHardiskMapsQ = 0
  Exit Function
gotErr:
  ReadHardiskMapsQ = -1
End Function

Public Function ReadHardiskMaps() As Long
  #If FinalMode Then
  On Error GoTo gotErr
  #End If
  If (TibiaVersionLong >= 1100) Then
    ReadHardiskMaps = ReadHardiskMapsQ()
    Exit Function
  End If
  Dim res As Long
  Dim Ammount As Long
  Dim readed As Long
  Dim progr As Double
  Dim fs As Scripting.FileSystemObject
  Dim f As Scripting.Folder
  Dim f1 As Scripting.File
  Dim fn As Integer
  Dim mapFile As String
  Dim strID As String
  Dim lngID As Long
  Dim thersize As Long
  Dim bigReadB(0 To cteBytesPerMap1) As Byte
  
  RemoveAllMapTranslation
  Ammount = 0
  readed = 0
  Set fs = New Scripting.FileSystemObject
  If fs.FolderExists(TibiaPath) = False Then
    ReadHardiskMaps = -1
    Exit Function
  End If
  Set f = fs.GetFolder(TibiaPath)
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "map" Then
      Ammount = Ammount + 1
    End If
  Next
  If Ammount = 0 Then
    ReadHardiskMaps = -1
    Exit Function
  End If
  thersize = Ammount - 1
  ReDim TheVeryBigMap(0 To cteBytesPerMap1, 0 To thersize)
  For Each f1 In f.Files
    If LCase(Right(f1.name, 3)) = "map" Then
      lngID = readed
      readed = readed + 1
      If LoadWasCompleted = False Then
        progr = (readed / Ammount) * 100
        frmLoading.NotifyLoadProgress progr, "Loading map file " & CStr(readed) & " / " & CStr(Ammount)
      End If
      fn = FreeFile
      mapFile = TibiaPath & "\" & f1.name
      strID = Left$(f1.name, 8)
      AddMapTranslation strID, lngID
      Open mapFile For Binary As fn
        Get fn, 1, bigReadB
      Close fn
      RtlMoveMemory TheVeryBigMap(0, lngID), bigReadB(0), cteBytesPerMap
    End If
  Next
  ReadHardiskMaps = 0
  Exit Function
gotErr:
  ReadHardiskMaps = -1
End Function

' Powerfull VB6 Function By Blackd
' Requires FreeImage (MFreeImage.bas and FreeImage.dll)
'
' Purpose: This is a function for copying the deciphered pixels of a PNG image into a raw byte array.
'
' Paramaters:
' Input/Output: arr, the byte array. It can be empty or already filled. Function will resize it if required.
' Optional Input: atPosition. Function will start writing bytes at atPosition
' Optional Input: withRotation. Function will rotate the image by this angle (in degrees) in first place.
' Output: imageWidth,imageHeight,imageBytesPerPixel (ByRef)
'
' Returns: False if it detects any error. True in other case.
Private Function LoadPNGIntoByteArray(ByRef filePath As String, ByRef arr() As Byte, _
 ByRef imageWidth As Long, _
 ByRef imageHeight As Long, _
 ByRef imageBytesPerPixel As Long, _
 Optional ByVal atPosition As Long = 0, _
 Optional ByVal withRotation As Double = 0) As Boolean
    On Error GoTo gotErr
    Dim imageRef As Long
    Dim pRow As Long
    Dim currentLine As Long
    Dim imageBytesPerLine As Long
    Dim totalBytes As Long
    Dim currentPos As Long
    Dim lastLine As Long
    Dim imageBitsPerPixel As Long
    Dim imageRef1 As Long
    ' Ensure our array is not an empty array, else we init it here.
    ensureMinimumArray arr
    ' Load the image
    imageRef = FreeImage_Load(FIF_PNG, filePath, FILO_PNG_DEFAULT)
    If imageRef = 0 Then
        ' ERROR: filePath is not PNG, or it does not exist
        LoadPNGIntoByteArray = False
        Exit Function
    End If
    If Not (withRotation = 0) Then
        imageRef1 = imageRef
        imageRef = FreeImage_RotateClassic(imageRef, withRotation)
        FreeImage_Unload imageRef1
        If imageRef = 0 Then
            ' ERROR: unable to rotate this PNG
            LoadPNGIntoByteArray = False
            Exit Function
        End If
    End If
    ' Determine some information from the loaded image
    imageWidth = FreeImage_GetWidth(imageRef)
    imageHeight = FreeImage_GetHeight(imageRef)
    imageBitsPerPixel = FreeImage_GetBPP(imageRef)
    imageBytesPerPixel = imageBitsPerPixel / 8
    imageBytesPerLine = FreeImage_GetLine(imageRef)
    totalBytes = imageBytesPerLine * imageHeight
    ' Is our array ready to hold so many bytes?
    If ((UBound(arr) - atPosition) < (totalBytes - 1)) Then
        ' If array is not big enough then we should increase it's size
        ReDim Preserve arr(atPosition + totalBytes - 1)
    End If
    currentPos = atPosition
    lastLine = imageHeight - 1
    For currentLine = 0 To lastLine
        pRow = FreeImage_GetScanline(imageRef, currentLine)
        If (Not (pRow = 0)) Then
          RtlMoveMemory arr(currentPos), ByVal pRow, imageBytesPerLine
        End If
        currentPos = currentPos + imageBytesPerLine
    Next currentLine
    ' Release the image
    FreeImage_Unload imageRef
    ' All looks OK. We return True
    LoadPNGIntoByteArray = True
    Exit Function
gotErr:
    '... else, we return False
    LoadPNGIntoByteArray = False
End Function

' If array is empty then init it to an array of 1 byte
Private Sub ensureMinimumArray(ByRef arr() As Byte)
    On Error GoTo gotErr
    Dim test As Long
    test = UBound(arr)
    Exit Sub
gotErr:
    ReDim arr(0)
End Sub
