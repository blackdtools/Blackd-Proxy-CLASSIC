Attribute VB_Name = "modTiles"
#Const FinalMode = 1
#Const TileDebug = 0
#Const ParseProperties = 1
Option Explicit
Public DBGtileError As String
Public Type TypeDatTile
  iscontainer As Boolean
  RWInfo As Long
  fluidcontainer As Boolean
  stackable As Boolean
  multitype As Boolean
  useable As Boolean
  notMoveable As Boolean
  alwaysOnTop As Boolean
  groundtile As Boolean
  blocking As Boolean
  blockpickupable As Boolean
  pickupable As Boolean
  blockingProjectile As Boolean
  canWalkThrough As Boolean
  noFloorChange As Boolean
  isDoor As Boolean
  isDoorWithLock As Boolean
  speed As Long
  canDecay As Boolean
  haveExtraByte As Boolean
  haveExtraByte2 As Boolean
  totalExtraBytes As Boolean
  isWater As Boolean
  stackPriority As Long
  haveFish As Boolean
  floorChangeUP As Boolean
  floorChangeDOWN As Boolean
  requireRightClick As Boolean
  requireRope As Boolean
  requireShovel As Boolean
  isFood As Boolean
  isField As Boolean
  isDepot As Boolean
  moreAlwaysOnTop As Boolean
  usable2 As Boolean
  multiCharge As Boolean
  haveName As Boolean
  itemName As String
End Type
Public highestDatTile As Long 'number of last Tile loaded
Public DatTiles() As TypeDatTile ' array of tiles
' Public DatTiles2() As TypeDatTile ' array of tiles - for debug compare
Public MAXDATTILES As Long
Public MAXTILEIDLISTSIZE As Long
Public AditionalStairsToDownFloor() As Long
Public AditionalStairsToUpFloor() As Long
Public AditionalRequireRope() As Long
Public AditionalRequireShovel() As Long

Private tileLog As String
Private atGraphicPart As Boolean
Private debugGraphicPart As Boolean
Private byteArray() As Byte

Public Function protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, lFactor) As Long
  On Error GoTo goterr
  Dim res As Long
  res = lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * lFactor
  protectedMult = res
  Exit Function
goterr:
  res = -1
End Function





' experimental loader of 7.40 dat:

Public Function LoadDatFile740(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim j As Long
  Dim lRare As Long
  #If FinalMode Then
    On Error GoTo badErr
  #End If

  tileOnDebug = 99999
  
  res = 0
  If (TibiaVersionLong >= 750) Then
    LoadDatFile740 = -2
    Exit Function
  End If
  
  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
    DatTiles(i).stackPriority = 1 ' custom flag, higher number, higher priority
    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0
  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2
  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  #If TileDebug = 1 Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If

  Open tibiadathere For Binary As fn
  #If TileDebug = 1 Then
  tileLog = "HEADER: "
  #End If
  For j = 1 To 12
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
  Next j
  #If TileDebug = 1 Then
      LogOnFile "tibiadatdebug.txt", tileLog
  #End If
  Do
    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    
    Get fn, , optByte
    
    While (optByte <> &HFF) And Not EOF(fn)
      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      
      Select Case optByte
      Case &H0 ' UNMODIFIED
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H2
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H3
        ' is container
        DatTiles(i).iscontainer = True
      Case &H4
        ' is stackable
        DatTiles(i).stackable = True
      Case &H5
        ' is useable
        DatTiles(i).useable = True
      Case &H6
        ' is ladder (only in id 1386)
      Case &H7
        ' unknown, 0 bytes
      Case &H8
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H9
        ' is fluid container
        DatTiles(i).fluidcontainer = True
      Case &HA
        ' multitype
        DatTiles(i).multitype = True
      Case &HB
        ' is blocking
        DatTiles(i).blocking = True
      Case &HC ' UNMODIFIED
        ' not moveable
        DatTiles(i).notMoveable = True
      Case &HD ' UNMODIFIED
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HE ' UNMODIFIED
        'blocks monster movement (flowers, parcels etc)
      Case &HF
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H10
        ' unknown, 0 bytes
      Case &H11
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H12
        ' ???
      Case &H13
        ' ???
      Case &H14
        ' player color templates
      Case &H15
        ' makes light -- skip bytes
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H16
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H17
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
      Case &H18
        ' unknown 4 bytes
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H19
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1A
        ' action posible
      Case &H1B
        'walls 2 types of them same material (total 4 pairs)
      Case &H1C
        ' unknown, 2 bytes
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1D
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        Select Case optbyte2
        Case &H4C
          'ladders
        Case &H4D
          'crate
        Case &H4E
          'rope spot?
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
          'hole
        Case &H57
          'items with special description?
        Case &H58
          'writtable?
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          debugByte = optByte
          ' ignore
        End Select 'optbyte2
        
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case &H1E ' UNMODIFIED
        ' unknown thing

      Case Else
        debugByte = optByte
        tileLog = tileLog & "!"
        ' ignore
      End Select 'optbyte
      Get fn, , optByte 'next optByte
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      LogOnFile "tibiadatdebug.txt", tileLog
    #End If

    ' add custom flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
     DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    Select Case i
    Case tileID_openHole, tileID_openHole2, tileID_sewerGate, tileID_stairsToDown, tileID_stairsToDown2, _
     tileID_woodenStairstoDown, tileID_trapdoor, tileID_trapdoor2, tileID_rampToDown, _
     tileID_closedHole, tileID_grassCouldBeHole, tileID_pitfall, tileID_desertLooseStonePile, _
     tileID_OpenDesertLooseStonePile, tileID_trapdoorKazordoon, tileID_stairsToDownKazordoon, _
     tileID_stairsToDownThais, tileID_down1, tileID_down2, tileID_down3
      DatTiles(i).floorChangeDOWN = True
    Case tileID_stairsToUp, tileID_woodenStairstoUp, tileID_ladderToUp, tileID_holeInCelling, _
     tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, tileID_rampToLeftCycMountain
      DatTiles(i).floorChangeUP = True
    End Select
    Select Case i
    Case tileID_sewerGate, tileID_ladderToUp
      DatTiles(i).requireRightClick = True
    End Select
    If i = tileID_holeInCelling Then
      DatTiles(i).requireRope = True
    End If
    
    Select Case i
    Case tileID_closedHole, tileID_desertLooseStonePile
      DatTiles(i).requireShovel = True
      DatTiles(i).floorChangeDOWN = True
      DatTiles(i).requireShovel = True
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).multitype = False
    End Select
  
  
  
    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' max priority
    End If
    
    'water
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If

    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    ' fields
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select







    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    
    
    'tileOnDebug = i
    
    Get fn, , b1
    
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = GoodHex(b1)
    End If
    #End If
    
    lWidth = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lHeight = CLng(b1)
    If lWidth > 1 Or lHeight > 1 Then
      'skip 1 byte
      Get fn, , b1
     #If TileDebug = 1 Then
      If i = tileOnDebug Then
        tileLog2 = tileLog2 & " " & GoodHex(b1)
      End If
      #End If
    End If
    Get fn, , b1
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lBlendframes = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lXdiv = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lYdiv = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lAnimcount = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lRare = CLng(b1) ' a strange new dimension for graphic info
    ' calculates the number of bytes of the graph and skip them
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 2)  'size = old formulae x lRare
    
    
    If DatTiles(i).haveExtraByte = True Then ' BYTECOUNTdat2
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
    If i = tileOnDebug Then
      tileLog2 = " Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
      For j = 1 To skipcount
        Get fn, , b1
        tileLog2 = tileLog2 & " " & GoodHex(b1)
      Next j
      LogOnFile "tibiadatdebug.txt", tileLog2
    Else
      a$ = Space$(skipcount)
      Get fn, , a$
    End If
    #Else
      a$ = Space$(skipcount)
      Get fn, , a$
    #End If










'
'    ' options zone done for this tile
'    ' now we get info about the graph of the tile...
'    Get fn, , b1
'    lWidth = CLng(b1)
'    Get fn, , b1
'    lHeight = CLng(b1)
'    If lWidth > 1 Or lHeight > 1 Then
'      'skip 1 byte
'       Get fn, , b1
'    End If
'    Get fn, , b1
'    lBlendframes = CLng(b1)
'    Get fn, , b1
'    lXdiv = CLng(b1)
'    Get fn, , b1
'    lYdiv = CLng(b1)
'    Get fn, , b1
'    lAnimcount = CLng(b1)
'    ' calculates the number of bytes of the graph and skip them
'    skipcount = lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * 2
'
'    If DatTiles(i).haveExtraByte = True Then ' BYTECOUNTdat1
'      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
'    End If
'    If DatTiles(i).haveExtraByte2 = True Then
'      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
'    End If
'
'    a$ = Space$(skipcount)
'    Get fn, , a$
    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile740 = -1
    Exit Function
  End If
   'DatTiles(&H9D3).haveExtraByte = True
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
endF:
  LoadDatFile740 = res
  Exit Function
badErr:
  LoadDatFile740 = -1
End Function











Public Function LoadDatFile(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim skipcount As Long
  Dim debugByte As Byte

  #If FinalMode Then
    On Error GoTo badErr
  #End If

  
  res = 0
  If (TibiaVersionLong >= 760) Then
    LoadDatFile = -2
    Exit Function
  End If
  
  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
    DatTiles(i).stackPriority = 1 ' custom flag, higher number, higher priority
    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0
  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2
  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  a$ = Space$(12)
  Get fn, , a$
  Do
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    
   ' If i = CLng(&HC0D) Or i = CLng(&HC40) Or i = CLng(&HC1A) Then
   '   LogOnFile "runes.text", "tile ID " & CStr(i) & " :"
   '   While (optByte <> &HFF)
   '     LogOnFile "runes.text", GoodHex(optByte)
   '     Get fn, , optByte
   '   Wend
   '   GoTo endAnalyze
   ' End If
    
    While (optByte <> &HFF) And Not EOF(fn)
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
        End If
        Get fn, , b2 'ignore next opt byte
      Case &H1
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H2
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H3
        ' is container
        DatTiles(i).iscontainer = True
      Case &H4
        ' is stackable
        DatTiles(i).stackable = True
      Case &H5
        ' is useable
        DatTiles(i).useable = True
      Case &H6
        ' is ladder (only in id 1386)
      Case &H7
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
      Case &H8
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        Get fn, , b2 ' always 4 max number of  newlines ?
      Case &H9
        ' is fluid container
        DatTiles(i).fluidcontainer = True
      Case &HA
        ' multitype
        DatTiles(i).multitype = True
      Case &HB
        ' is blocking
        DatTiles(i).blocking = True
      Case &HC
        ' not moveable
        DatTiles(i).notMoveable = True
      Case &HD
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HE
        'blocks monster movement (flowers, parcels etc)
      Case &HF
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H10
        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around
        Get fn, , b2 ' 0
        Get fn, , b1 ' = 215 for items , =208 for non items
        Get fn, , b2 ' 0
      Case &H11
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H12
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
      Case &H13
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        Get fn, , b2 ' always 0
      Case &H14
        ' player color templates
      Case &H16
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        Get fn, , b2
      Case &H17
        ' seems like decorables with 4 states of turning (exception first 4 are unique statues)
      Case &H18
        ' corpses that don't decay
        DatTiles(i).canDecay = False
      Case &H19
        'wall items
      Case &H1A
        ' action posible
      Case &H1B
        'walls 2 types of them same material (total 4 pairs)
      Case &H1C
        'monster has animation even when iddle (rot, wasp, slime, fe)
      Case &H1D
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        Select Case optbyte2
        Case &H4C
          'ladders
        Case &H4D
          'crate
        Case &H4E
          'rope spot?
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
          'hole
        Case &H57
          'items with special description?
        Case &H58
          'writtable?
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          debugByte = optByte
          ' ignore
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
      Case Else
        debugByte = optByte
        ' ignore
      End Select 'optbyte
      Get fn, , optByte 'next optByte
    Wend
endAnalyze:
    ' add custom flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
     DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    Select Case i
    Case tileID_openHole, tileID_openHole2, tileID_sewerGate, tileID_stairsToDown, tileID_stairsToDown2, _
     tileID_woodenStairstoDown, tileID_trapdoor, tileID_trapdoor2, tileID_rampToDown, _
     tileID_closedHole, tileID_grassCouldBeHole, tileID_pitfall, tileID_desertLooseStonePile, _
     tileID_OpenDesertLooseStonePile, tileID_trapdoorKazordoon, tileID_stairsToDownKazordoon, _
     tileID_stairsToDownThais, tileID_down1, tileID_down2, tileID_down3
      DatTiles(i).floorChangeDOWN = True
    Case tileID_stairsToUp, tileID_woodenStairstoUp, tileID_ladderToUp, tileID_holeInCelling, _
     tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, tileID_rampToLeftCycMountain
      DatTiles(i).floorChangeUP = True
    End Select
    Select Case i
    Case tileID_sewerGate, tileID_ladderToUp
      DatTiles(i).requireRightClick = True
    End Select
    If i = tileID_holeInCelling Then
      DatTiles(i).requireRope = True
    End If
    
    Select Case i
    Case tileID_closedHole, tileID_desertLooseStonePile
      DatTiles(i).requireShovel = True
      DatTiles(i).floorChangeDOWN = True
      DatTiles(i).requireShovel = True
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).multitype = False
    End Select
  
  
  
    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' max priority
    End If
    
    'water
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If

    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    ' fields
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select

    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    Get fn, , b1
    lWidth = CLng(b1)
    Get fn, , b1
    lHeight = CLng(b1)
    If lWidth > 1 Or lHeight > 1 Then
      'skip 1 byte
       Get fn, , b1
    End If
    Get fn, , b1
    lBlendframes = CLng(b1)
    Get fn, , b1
    lXdiv = CLng(b1)
    Get fn, , b1
    lYdiv = CLng(b1)
    Get fn, , b1
    lAnimcount = CLng(b1)
    ' calculates the number of bytes of the graph and skip them
    skipcount = lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * 2
    
    If DatTiles(i).haveExtraByte = True Then ' BYTECOUNTdat1
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    
    a$ = Space$(skipcount)
    Get fn, , a$
    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile = -1
    Exit Function
  End If
   'DatTiles(&H9D3).haveExtraByte = True
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
endF:
  LoadDatFile = res
  Exit Function
badErr:
  LoadDatFile = -1
End Function


Public Function LoadDatFile2(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999
  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
    DatTiles(i).stackPriority = 1 ' custom flag, higher number, higher priority
    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0
  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2
  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

   #If TileDebug = 1 Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If

  Open tibiadathere For Binary As fn
  #If TileDebug = 1 Then
  tileLog = "HEADER: "
  #End If
  For j = 1 To 12
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
  Next j
  #If TileDebug = 1 Then
      LogOnFile "tibiadatdebug.txt", tileLog
  #End If
  Do
    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)
      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' used to be &H1
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 'used to be &H2
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' used to be &H3
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' used to be &H4
        ' is stackable
        DatTiles(i).stackable = True
      Case &H7 ' used to be &H5
        ' is useable
        DatTiles(i).useable = True
      Case &H6
        ' new flag?
      Case &H8 'used to be &H7
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H9 'used to be &H8
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA 'used to be &H9
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' used to be &HA
        ' multitype
        DatTiles(i).multitype = True
      Case &HC ' used to be &HB
        ' is blocking
        DatTiles(i).blocking = True
      Case &HD ' used to be &HC
        ' not moveable
        DatTiles(i).notMoveable = True
      Case &HE 'used to be &HD
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF 'used to be &HE
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 'used to be &HF
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H15 'used to be &H10
        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H11
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H1E 'used to be &H12
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
      Case &H19 ' used to be &H13
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H14
         ' new property: unknown
      Case &H18 '
        ' new property : unknown
        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1C 'used to be &H16
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H17
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
      Case &H1A 'used to be &H18
        ' corpses that don't decay
        DatTiles(i).canDecay = False
      Case &H1B ' used to be &H19
        'wall items
      Case &H12 ' used to be &H1A
        ' action posible
      Case &H13 ' used to be &H1B
        'walls 2 types of them same material (total 4 pairs)
      Case &H1D ' not changed
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True

      DatTiles(i).alwaysOnTop = True
      DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          debugByte = optByte
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1) & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        LogOnFile "tibiadatdebug.txt", "ERROR AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " < "
      LogOnFile "tibiadatdebug.txt", tileLog
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    #If TileDebug = 1 Then
    If i = tileOnDebug Then
      tileLog2 = GoodHex(b1)
    End If
    #End If
    
    lWidth = CLng(b1)
    Get fn, , b1
    #If TileDebug Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lHeight = CLng(b1)
    If lWidth > 1 Or lHeight > 1 Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug Then
      If i = tileOnDebug Then
        tileLog2 = tileLog2 & " " & GoodHex(b1)
      End If
      #End If
    End If
    Get fn, , b1
    #If TileDebug Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lBlendframes = CLng(b1)
    Get fn, , b1
    #If TileDebug Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lXdiv = CLng(b1)
    Get fn, , b1
    #If TileDebug Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lYdiv = CLng(b1)
    Get fn, , b1
    #If TileDebug Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lAnimcount = CLng(b1)
    Get fn, , b1
    #If TileDebug Then
    If i = tileOnDebug Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    End If
    #End If
    lRare = CLng(b1) ' a strange new dimension for graphic info
    ' calculates the number of bytes of the graph and skip them
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 2)  'size = old formulae x lRare
    
    
    If DatTiles(i).haveExtraByte = True Then ' BYTECOUNTdat2
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    
    #If TileDebug Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
    If i = tileOnDebug Then
      tileLog2 = " Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
      For j = 1 To skipcount
        Get fn, , b1
        tileLog2 = tileLog2 & " " & GoodHex(b1)
      Next j
      LogOnFile "tibiadatdebug.txt", tileLog2
    Else
      a$ = Space$(skipcount)
    End If
    #Else
      a$ = Space$(skipcount)
    #End If
    Get fn, , a$
    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile2 = -1
    Exit Function
  End If
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
endF:
  LoadDatFile2 = res
  Exit Function
badErr:
  LoadDatFile2 = -1 ' bad format or wrong version of given tibia.dat
End Function

' Tibia function : tibia.dat reader for Tibia 7.8
' COPYRIGHT of Blackd ( www.blackdtools.com ) , please do not repost in other place without permission.
Public Function LoadDatFile3(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
    DatTiles(i).stackPriority = 1 ' custom flag, higher number, higher priority
    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0
  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2
  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1

  If (TibiaVersionLong >= 860) Then ' check version byte
    If (b1 <> &H46) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 854) Then ' check version byte
    If (b1 <> &H45) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 853) Then ' check version byte
    If (b1 <> &H44) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 852) Then ' check version byte
    If (b1 <> &H44) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 850) Then ' check version byte
    If (b1 <> &H44) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 840) Then ' check version byte
    If (b1 <> &H43) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 820) Then ' check version byte
    If (b1 <> &H39) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 810) Then ' check version byte
    If (b1 <> &H37) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 800) Then ' check version byte
    If (b1 <> &H23) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 792) Then ' check version byte
    If (b1 <> &H1F) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 773) Then
    If (b1 <> &H1B) Then
      LoadDatFile3 = -2
      Exit Function
    End If
  End If
  'a$ = Space$(3) ' descartado, podria dar problemas
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Do
    #If TileDebug Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)
      #If TileDebug Then
        tileLog = tileLog & " " & GoodHex(optByte)
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1 ' OK - used to be &H1
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' OK - used to be &H2
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' OK - used to be &H3
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' OK - used to be &H4
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' OK - used to be &H5
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' OK - used to be &H6
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' OK - NEW
        DatTiles(i).usable2 = True
      Case &H8 ' OK
        DatTiles(i).multiCharge = True
      Case &H9 ' OK - used to be &H8
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA 'used to be &H9
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case &HB 'used to be &HA
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HC ' used to be &HB
        ' multitype
        DatTiles(i).multitype = True
      Case &HD ' OK - used to be &HC
        ' is blocking
        DatTiles(i).blocking = True
      Case &HE ' OK - used to be &HD
        ' not moveable
        DatTiles(i).notMoveable = True
      Case &HF ' OK - used to be &HE
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &H10 'used to be &HF
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H11 'used to be &H10
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H16 'used to be &H15
        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H12 'used to be &H11
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H1F ' OK - used to be &H1E
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
      Case &H1A ' used to be &H19
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H15 ' used to be &H14
         ' unknown
      Case &H19 ' used to be &H18
        ' unknown
        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1D 'used to be &H1C
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H18 ' used to be &H17
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
      Case &H1B 'used to be &H1A
        ' corpses that don't decay
        DatTiles(i).canDecay = False
      Case &H1C ' used to be &H1B
        'wall items
      Case &H13 ' used to be &H12
        ' action posible
      Case &H14 ' used to be &H13
        'walls 2 types of them same material (total 4 pairs)
      Case &H1E ' used to be &H1D
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          debugByte = optByte
          #If TileDebug Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case &H20
        'new flag since tibia 8.57
  
      Case &H17
        'new flag since tibia 8.57
    
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug Then
      If nextB <= optByte Then
        LogOnFile "tibiadatdebug.txt", "ERROR AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug Then
      tileLog = tileLog & " " & GoodHex(optByte) & " < "
      LogOnFile "tibiadatdebug.txt", tileLog
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    #If TileDebug = 1 Then
   
      tileLog2 = GoodHex(b1)

    #End If
    
    lWidth = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    #End If
    lHeight = CLng(b1)
    If lWidth > 1 Or lHeight > 1 Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " " & GoodHex(b1)
      #End If
    End If
    Get fn, , b1
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    #End If
    lBlendframes = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    #End If
    lXdiv = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    #End If
    lYdiv = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    #End If
    lAnimcount = CLng(b1)
    Get fn, , b1
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " " & GoodHex(b1)
    #End If
    lRare = CLng(b1) ' a strange new dimension for graphic info
    ' calculates the number of bytes of the graph and skip them
    'LogOnFile "tibiadatdebug.txt", "tile #" & CStr(i) & ": " & tileLog2
    
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 2)
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile3 = -5 ' unexpected overflow
      Exit Function
    End If
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 2)  'size = old formulae x lRare
    If DatTiles(i).haveExtraByte = True Then ' BYTECOUNTdat3
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = " Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile3 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  LoadDatFile3 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile3 = -4 ' bad format or wrong version of given tibia.dat
End Function

' for tibia 8.6 and higher
Public Function LoadDatFile4(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 2110

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1

  If (TibiaVersionLong >= 860) Then ' check version byte
    If (b1 <> &H46) Then
      LoadDatFile4 = -2
      Exit Function
    End If
  Else
      LoadDatFile4 = -2
      Exit Function
  End If
  'a$ = Space$(3) ' descartado, podria dar problemas
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1 ' UNMODIFIED
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

'      Case &H8 ' DELETED !!
'        DatTiles(i).multiCharge = True
        
      Case &H8 ' used to be &H9 ' NEW - OK
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' used to be &HA ' NEW - OK
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' used to be &HB ' NEW - OK
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' used to be &HC ' NEW - OK
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' OK - used to be &HD ' NEW - OK
        ' is blocking
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' OK - used to be &HE ' NEW - OK
        ' not moveable
        DatTiles(i).notMoveable = True
      Case &HE ' OK - used to be &HF ' NEW - OK
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF 'used to be &H10 ' NEW - OK
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' used to be &H11 - ' NEW - OK
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H15 ' used to be &H17 - ' NEW - OK
        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H11 ' used to be &H12 - ' NEW - OK
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H1E ' used to be &H1F - ' NEW - OK
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True

      Case &H19 ' used to be &H1A ' NEW - OK
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H14 ' used to be &H15 ' NEW - OK
         ' unknown
      Case &H18 ' used to be &H19 ' NEW - OK
        ' unknown
        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1C 'used to be &H1D ' NEW - OK
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
        ' xxxxxxxx
         Case &H17 ' used to be &H18 ' NEW - OK
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
      Case &H1A ' used to be &H1B ' NEW - OK
        ' corpses that don't decay
        DatTiles(i).canDecay = False
      Case &H1B ' used to be &H1C ' NEW - OK
        'wall items
      Case &H12 ' used to be &H13 ' NEW - OK
        ' action posible
      Case &H13 ' used to be &H14 ' NEW - OK
        'walls 2 types of them same material (total 4 pairs)
      Case &H1D ' used to be &H1E ' NEW - OK
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          debugByte = optByte
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case &H1F  ' used to be &H20 ' NEW - OK
        'new flag since tibia 8.57
  
      Case &H16 ' used to be &H17 ' NEW - OK
        'new flag since tibia 8.57

    
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        LogOnFile "tibiadatdebug.txt", "ERROR AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If

    
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 2)
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile4 = -5 ' unexpected overflow
      Exit Function
    End If
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 2)  'size = old formulae x lRare
    If DatTiles(i).haveExtraByte = True Then ' BYTECOUNTdat4
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile4 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile4 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile4 = -4 ' bad format or wrong version of given tibia.dat
End Function









' for tibia 8.72 and higher
Public Function LoadDatFile5(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 2110

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  If TibiaVersionLong >= 910 Then
    If (b1 <> &H48) Then
      LoadDatFile5 = -2
      Exit Function
    End If
  ElseIf (TibiaVersionLong >= 860) Then ' check version byte
    If (b1 <> &H46) Then
      LoadDatFile5 = -2
      Exit Function
    End If
  Else
      LoadDatFile5 = -2
      Exit Function
  End If
  'a$ = Space$(3) ' descartado, podria dar problemas
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
                    
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1 ' UNMODIFIED
        
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

'      Case &H8 ' DELETED !!
'        DatTiles(i).multiCharge = True
        
      Case &H8 ' used to be &H9 ' NEW - OK
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' used to be &HA ' NEW - OK
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' used to be &HB ' NEW - OK
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' used to be &HC ' NEW - OK
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' OK - used to be &HD ' NEW - OK
        ' is blocking
        
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' OK - used to be &HE ' NEW - OK
        ' not moveable
                 
        DatTiles(i).notMoveable = True
      Case &HE ' OK - used to be &HF ' NEW - OK
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF 'used to be &H10 ' NEW - OK
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' used to be &H11 - ' NEW - OK
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H15 ' used to be &H17 - ' NEW - OK

        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H11 ' used to be &H12 - ' NEW - OK
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H1E ' used to be &H1F - ' NEW - OK
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True

      Case &H19 ' used to be &H1A ' NEW - OK
            
      
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H14 ' used to be &H15 ' NEW - OK
         ' unknown
      Case &H18 ' used to be &H19 ' NEW - OK
        ' unknown

        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1C 'used to be &H1D ' NEW - OK
        
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If

        
        
        
        
        
        ' xxxxxxxx
         Case &H17 ' used to be &H18 ' NEW - OK
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
      Case &H1A ' used to be &H1B ' NEW - OK
        ' corpses that don't decay
        DatTiles(i).canDecay = False
      Case &H1B ' used to be &H1C ' NEW - OK
        'wall items
      Case &H12 ' used to be &H13 ' NEW - OK
        ' action posible
      Case &H13 ' used to be &H14 ' NEW - OK
        'walls 2 types of them same material (total 4 pairs)
      Case &H1D ' used to be &H1E ' NEW - OK
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          debugByte = optByte
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case &H1F  ' used to be &H20 ' NEW - OK
        'new flag since tibia 8.57
        
      Case &H20
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
 
  
      Case &H16 ' used to be &H17 ' NEW - OK
        'new flag since tibia 8.57

    
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        LogOnFile "tibiadatdebug.txt", "WARNING AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If
    If lRare > &H1 Then
          DatTiles(i).haveExtraByte2 = True ' UNKNOWN , TEST
    End If
    If DatTiles(i).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then

      LogOnFile "tibiadatdebug.txt", tileLog2 & vbCrLf

    #End If
    
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 2)
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile5 = -5 ' unexpected overflow
      Exit Function
    End If
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 2)  'size = old formulae x lRare
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile5 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile5 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile5 = -4 ' bad format or wrong version of given tibia.dat
End Function



' for tibia 9.4 and higher
Public Function LoadDatFile6(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  Dim tmpSize As Long
  Dim tmpI As Long
  Dim tmpName As String
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 2110

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  
  If TibiaVersionLong >= 940 Then
    If (b1 <> &H4C) Then
      LoadDatFile6 = -2
      Exit Function
    End If
  Else
      LoadDatFile6 = -2
      Exit Function
  End If
  'a$ = Space$(3) ' descartado, podria dar problemas
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
                    
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1 ' UNMODIFIED
        
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

'      Case &H8 ' DELETED !!
'        DatTiles(i).multiCharge = True
        
      Case &H8 ' used to be &H9 ' NEW - OK
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' used to be &HA ' NEW - OK
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' used to be &HB ' NEW - OK
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' used to be &HC ' NEW - OK
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' OK - used to be &HD ' NEW - OK
        ' is blocking
        
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' OK - used to be &HE ' NEW - OK
        ' not moveable
                 
        DatTiles(i).notMoveable = True
      Case &HE ' OK - used to be &HF ' NEW - OK
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF 'used to be &H10 ' NEW - OK
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' used to be &H11 - ' NEW - OK
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H15 ' used to be &H17 - ' NEW - OK

        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H11 ' used to be &H12 - ' NEW - OK
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H1E ' used to be &H1F - ' NEW - OK
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True

      Case &H19 ' used to be &H1A ' NEW - OK
            
      
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H14 ' used to be &H15 ' NEW - OK
         ' unknown
      Case &H18 ' used to be &H19 ' NEW - OK
        ' unknown

        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1C 'used to be &H1D ' NEW - OK
        
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If

        
        
        
        
        
        ' xxxxxxxx
         Case &H17 ' used to be &H18 ' NEW - OK
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
      Case &H1A ' used to be &H1B ' NEW - OK
        ' corpses that don't decay
        DatTiles(i).canDecay = False
      Case &H1B ' used to be &H1C ' NEW - OK
        'wall items
      Case &H12 ' used to be &H13 ' NEW - OK
        ' action posible
      Case &H13 ' used to be &H14 ' NEW - OK
        'walls 2 types of them same material (total 4 pairs)
      Case &H1D ' used to be &H1E ' NEW - OK
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          debugByte = optByte
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case &H1F  ' used to be &H20 ' NEW - OK
        'new flag since tibia 8.57
        
      Case &H20
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
 
  
      Case &H16 ' used to be &H17 ' NEW - OK
        'new flag since tibia 8.57

      Case &H21 ' item group, something, and name (tibia 9.4)
        Get fn, , b1 ' item group (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' item group (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' size of text (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' size of text (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
        tmpSize = GetTheLong(b1, b2)
        tmpName = ""
        For tmpI = 1 To tmpSize
            Get fn, , b1 ' size of text
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            tmpName = tmpName & Chr(b1)
        Next tmpI
        DatTiles(i).haveName = True
        DatTiles(i).itemName = tmpName
        #If TileDebug = 1 Then
          tileLog = tileLog & " (" & tmpName & ")"
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        LogOnFile "tibiadatdebug.txt", "WARNING AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If
    If lRare > &H1 Then
          DatTiles(i).haveExtraByte2 = True ' UNKNOWN , TEST
    End If
    If DatTiles(i).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then

      LogOnFile "tibiadatdebug.txt", tileLog2 & vbCrLf

    #End If
    
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 2)
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile6 = -5 ' unexpected overflow
      Exit Function
    End If
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 2)  'size = old formulae x lRare
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile6 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile6 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile6 = -4 ' bad format or wrong version of given tibia.dat
End Function










' for tibia 9.6 and higher
Public Function LoadDatFile7(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  Dim tmpSize As Long
  Dim tmpI As Long
  Dim tmpName As String
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 2110

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  
  If TibiaVersionLong >= 980 Then
    If (b1 <> &H9E) Then
      LoadDatFile7 = -2
      Exit Function
    End If
  ElseIf TibiaVersionLong >= 960 Then
    If (b1 <> &H4C) Then
      LoadDatFile7 = -2
      Exit Function
    End If
  Else
      LoadDatFile7 = -2
      Exit Function
  End If
  'a$ = Space$(3) ' descartado, podria dar problemas
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
                    
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1 ' UNMODIFIED
        
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

      Case &H8 ' used to be &H9 ' NEW - OK
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' used to be &HA ' NEW - OK
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' used to be &HB ' NEW - OK
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' used to be &HC ' NEW - OK
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' OK - used to be &HD ' NEW - OK
        ' is blocking
        
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' OK - used to be &HE ' NEW - OK
        ' not moveable
                 
        DatTiles(i).notMoveable = True
      Case &HE ' OK - used to be &HF ' NEW - OK
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF 'used to be &H10 ' NEW - OK
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' used to be &H11 - ' NEW - OK
        ' pickupable / equipable
        DatTiles(i).pickupable = True
      Case &H15 ' used to be &H17 - ' NEW - OK

        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H11 ' used to be &H12 - ' NEW - OK
        ' can see what is under (ladder holes, stairs holes etc)
      Case &H1E ' used to be &H1F - ' NEW - OK
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True

      Case &H19 ' used to be &H1A ' NEW - OK
            
      
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H14 ' used to be &H15 ' NEW - OK
         ' unknown
      Case &H18 ' used to be &H19 ' NEW - OK
        ' unknown

        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1C 'used to be &H1D ' NEW - OK
        
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If

        
        
        
        
        
        ' xxxxxxxx
         Case &H17 ' used to be &H18 ' NEW - OK
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
      Case &H1A ' used to be &H1B ' NEW - OK
        ' corpses that don't decay
        DatTiles(i).canDecay = False
      Case &H1B ' used to be &H1C ' NEW - OK
        'wall items
      Case &H12 ' used to be &H13 ' NEW - OK
        ' action posible
      Case &H13 ' used to be &H14 ' NEW - OK
        'walls 2 types of them same material (total 4 pairs)
      Case &H1D ' used to be &H1E ' NEW - OK
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          
          debugByte = optByte
          Debug.Print "Tile loader found unexpected properties (" & GoodHex(optByte) & ")"
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
      Case &H1F  ' used to be &H20 ' NEW - OK
        'new flag since tibia 8.57
        
      Case &H20
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
 
  
      Case &H16 ' used to be &H17 ' NEW - OK
        'new flag since tibia 8.57

      Case &H21 ' item group, something, and name (tibia 9.4)
        Get fn, , b1 ' item group (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' item group (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' size of text (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' size of text (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
        tmpSize = GetTheLong(b1, b2)
        tmpName = ""
        For tmpI = 1 To tmpSize
            Get fn, , b1 ' size of text
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            tmpName = tmpName & Chr(b1)
        Next tmpI
        DatTiles(i).haveName = True
        DatTiles(i).itemName = tmpName
        #If TileDebug = 1 Then
          tileLog = tileLog & " (" & tmpName & ")"
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        LogOnFile "tibiadatdebug.txt", "WARNING AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If
    If lRare > &H1 Then
          DatTiles(i).haveExtraByte2 = True ' UNKNOWN , TEST
    End If
    If DatTiles(i).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then

      LogOnFile "tibiadatdebug.txt", tileLog2 & vbCrLf

    #End If
    ' NEW since Tibia 9.6: double size for graphic item references
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 4)
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile7 = -5 ' unexpected overflow
      Exit Function
    End If
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 4)  'size = old formulae x lRare
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile7 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile7 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile7 = -4 ' bad format or wrong version of given tibia.dat
End Function



' for tibia 10.0 and higher
Public Function LoadDatFile8(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  Dim tmpSize As Long
  Dim tmpI As Long
  Dim tmpName As String
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 2110

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  
  If TibiaVersionLong >= 980 Then
    If (b1 <> &H9E) Then
      LoadDatFile8 = -2
      Exit Function
    End If
  ElseIf TibiaVersionLong >= 960 Then
    If (b1 <> &H4C) Then
      LoadDatFile8 = -2
      Exit Function
    End If
  Else
      LoadDatFile8 = -2
      Exit Function
  End If
  'a$ = Space$(3) ' descartado, podria dar problemas
  Get fn, , b1
  Get fn, , b1
  Get fn, , b1
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
                    
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1 ' UNMODIFIED
        
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

      Case &H8 ' UNMODIFIED
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' UNMODIFIED
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' UNMODIFIED
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' UNMODIFIED
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' UNMODIFIED
        ' is blocking
        
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' UNMODIFIED
        ' not moveable
                 
        DatTiles(i).notMoveable = True
      Case &HE ' UNMODIFIED
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF ' UNMODIFIED
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' UNMODIFIED
        ' pickupable / equipable
        DatTiles(i).pickupable = True
    
      Case &H11 ' UNMODIFIED
        ' can see what is under (ladder holes, stairs holes etc)


      Case &H12 ' UNMODIFIED
        ' action posible
      Case &H13 ' UNMODIFIED
        'walls 2 types of them same material (total 4 pairs)
      Case &H14 ' UNMODIFIED
         ' unknown
      Case &H15 ' NEW?? - UNTESTED
         ' unknown
       
      Case &H16 ' used to be &H15 - ' NEW - UNTESTED

        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H18 '  used to be &H17 - ' NEW - UNTESTED
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
        
      Case &H19 ' used to be &H18 - ' NEW - UNTESTED
        ' unknown

        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1A ' used to be &H19 - ' NEW - UNTESTED
            
      
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1B ' used to be &H1A - ' NEW - UNTESTED
        ' corpses that don't decay
        DatTiles(i).canDecay = False
        
      Case &H1C ' used to be &H1B - ' NEW - UNTESTED
        'wall items
        
      Case &H1D ' used to be &H1C - ' NEW - UNTESTED
        
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If

        
        
        
        
        
        ' xxxxxxxx



      Case &H1E ' used to be &H1D - ' NEW - UNTESTED
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          
          debugByte = optByte
          Debug.Print "Tile loader found unexpected properties (" & GoodHex(optByte) & ")"
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
        
      Case &H1F ' used to be &H1E - ' NEW - UNTESTED
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
        
      Case &H20  ' used to be &H1F - ' NEW - UNTESTED
        'new flag since tibia 8.57
        
      Case &H21 ' used to be &H20 - ' NEW - UNTESTED
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
 
  
 

      Case &H22 ' used to be &H21 - ' NEW - UNTESTED
        Get fn, , b1 ' item group (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' item group (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' size of text (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' size of text (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
        tmpSize = GetTheLong(b1, b2)
        tmpName = ""
        For tmpI = 1 To tmpSize
            Get fn, , b1 ' size of text
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            tmpName = tmpName & Chr(b1)
        Next tmpI
        DatTiles(i).haveName = True
        DatTiles(i).itemName = tmpName
        #If TileDebug = 1 Then
          tileLog = tileLog & " (" & tmpName & ")"
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        LogOnFile "tibiadatdebug.txt", "WARNING AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If
    If lRare > &H1 Then
          DatTiles(i).haveExtraByte2 = True ' UNKNOWN , TEST
    End If
    If DatTiles(i).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then

      LogOnFile "tibiadatdebug.txt", tileLog2 & vbCrLf

    #End If
    ' NEW since Tibia 9.6: double size for graphic item references
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 4)
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile8 = -5 ' unexpected overflow
      Exit Function
    End If
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 4)  'size = old formulae x lRare
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile8 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile8 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile8 = -4 ' bad format or wrong version of given tibia.dat
End Function



' for tibia 10.21 and higher
Public Function LoadDatFile9(ByVal tibiadathere As String) As Integer
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  Dim tmpSize As Long
  Dim tmpI As Long
  Dim tmpName As String
  
    Dim limit_ITEM_COUNT As Long
  Dim limit_OUTFIT_COUNT As Long
  Dim limit_EFFECT_COUNT As Long
  Dim limit_DISTANCE_COUNT As Long
  Dim dat_version As Long
  
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 2110

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

   Open tibiadathere For Binary As fn
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2

  'Debug.Print GoodHex(b1)
  Get fn, , b3
  'Debug.Print GoodHex(b1)
  Get fn, , b4
 ' Debug.Print GoodHex(b1)

 dat_version = FourBytesLong(b1, b2, b3, b4)
 ' tibia 10.58 = 1412240103
  Get fn, , b1
  'Debug.Print GoodHex(b1)
  Get fn, , b2
  'Debug.Print GoodHex(b2)
  limit_ITEM_COUNT = GetTheLong(b1, b2)
  
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2
  'Debug.Print GoodHex(b2)
  limit_OUTFIT_COUNT = GetTheLong(b1, b2)
  Get fn, , b1
  'Debug.Print GoodHex(b1)

  Get fn, , b2
 ' Debug.Print GoodHex(b2)
   limit_EFFECT_COUNT = GetTheLong(b1, b2)
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  
  Get fn, , b2
  'Debug.Print GoodHex(b2)
   limit_DISTANCE_COUNT = GetTheLong(b1, b2)
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        lonNumber = CLng(b1)
        DatTiles(i).speed = lonNumber
        If lonNumber = 0 Then
          DatTiles(i).blocking = True
                    
        End If
        Get fn, , b2 'ignore next opt byte
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H1 ' UNMODIFIED
        
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

      Case &H8 ' UNMODIFIED
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' UNMODIFIED
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' UNMODIFIED
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' UNMODIFIED
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' UNMODIFIED
        ' is blocking
        
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' UNMODIFIED
        ' not moveable
                 
        DatTiles(i).notMoveable = True
      Case &HE ' UNMODIFIED
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF ' UNMODIFIED
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' UNMODIFIED
        ' pickupable / equipable
        DatTiles(i).pickupable = True
    
      Case &H11 ' UNMODIFIED
        ' can see what is under (ladder holes, stairs holes etc)


      Case &H12 ' UNMODIFIED
        ' action posible
      Case &H13 ' UNMODIFIED
        'walls 2 types of them same material (total 4 pairs)
      Case &H14 ' UNMODIFIED
         ' unknown
      Case &H15 ' NEW?? - UNTESTED
         ' unknown
       
      Case &H16 ' used to be &H15 - ' NEW - UNTESTED

        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H17 ' NEW?? - UNTESTED
         ' unknown
      Case &H18 '  used to be &H17 - ' NEW - UNTESTED
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
        
      Case &H19 ' used to be &H18 - ' NEW - UNTESTED
        ' unknown

        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1A ' used to be &H19 - ' NEW - UNTESTED
            
      
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1B ' used to be &H1A - ' NEW - UNTESTED
        ' corpses that don't decay
        DatTiles(i).canDecay = False
        
      Case &H1C ' used to be &H1B - ' NEW - UNTESTED
        'wall items
        
      Case &H1D ' used to be &H1C - ' NEW - UNTESTED
        
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If

        
        
        
        
        
        ' xxxxxxxx



      Case &H1E ' used to be &H1D - ' NEW - UNTESTED
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2
        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          
          debugByte = optByte
          Debug.Print "Tile loader found unexpected properties (" & GoodHex(optByte) & ")"
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
        
      Case &H1F ' used to be &H1E - ' NEW - UNTESTED
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
        
      Case &H20  ' used to be &H1F - ' NEW - UNTESTED
        'new flag since tibia 8.57
        
      Case &H21 ' used to be &H20 - ' NEW - UNTESTED
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
 
  
 

      Case &H22 ' used to be &H21 - ' NEW - UNTESTED
        Get fn, , b1 ' item group (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' item group (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' size of text (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' size of text (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
        tmpSize = GetTheLong(b1, b2)
        tmpName = ""
        For tmpI = 1 To tmpSize
            Get fn, , b1 ' size of text
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            tmpName = tmpName & Chr(b1)
        Next tmpI
        DatTiles(i).haveName = True
        DatTiles(i).itemName = tmpName
        #If TileDebug = 1 Then
          tileLog = tileLog & " (" & tmpName & ")"
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
        
      Case &H23 ' NEW since 10.21 - UNTESTED
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
    
      Case &HFE
        ' unknown new flag since tibia 10.21
        
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        If (optByte = &H20) Or (optByte = &H21) Or (optByte = &H22) Or (optByte = &H23) Or (optByte = &HFE) Then
        
        Else
          LogOnFile "tibiadatdebug.txt", "WARNING AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
        End If
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If
    If lRare > &H1 Then
          DatTiles(i).haveExtraByte2 = True ' UNKNOWN , TEST
    End If
    If DatTiles(i).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then

      LogOnFile "tibiadatdebug.txt", tileLog2 & vbCrLf

    #End If
    ' NEW since Tibia 9.6: double size for graphic item references
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 4)
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile9 = -5 ' unexpected overflow
      Exit Function
    End If
    skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 4)  'size = old formulae x lRare
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile9 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile9 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile9 = -4 ' bad format or wrong version of given tibia.dat
End Function






' for tibia 10.5 and higher
Public Function LoadDatFile10(ByVal tibiadathere As String) As Integer
  Dim addToSkipCount As Long
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  Dim tmpSize As Long
  Dim tmpI As Long
  Dim tmpName As String
  Dim limit_ITEM_COUNT As Long
  Dim limit_OUTFIT_COUNT As Long
  Dim limit_EFFECT_COUNT As Long
  Dim limit_DISTANCE_COUNT As Long
  Dim dat_version As Long
  
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 2110

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2

  'Debug.Print GoodHex(b1)
  Get fn, , b3
  'Debug.Print GoodHex(b1)
  Get fn, , b4
 ' Debug.Print GoodHex(b1)

 dat_version = FourBytesLong(b1, b2, b3, b4)
 ' tibia 10.58 = 1412240103
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2
  'Debug.Print GoodHex(b2)
  limit_ITEM_COUNT = GetTheLong(b1, b2)
  
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2
 ' Debug.Print GoodHex(b2)
  limit_OUTFIT_COUNT = GetTheLong(b1, b2)
  Get fn, , b1
  'Debug.Print GoodHex(b1)

  Get fn, , b2
 ' Debug.Print GoodHex(b2)
   limit_EFFECT_COUNT = GetTheLong(b1, b2)
  Get fn, , b1
  'Debug.Print GoodHex(b1)
  
  Get fn, , b2
 ' Debug.Print GoodHex(b2)
   limit_DISTANCE_COUNT = GetTheLong(b1, b2)
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        If ((TibiaVersionLong >= 1058) And (i = 21505)) Then
          ' rare case: only skip 1
            Get fn, , b1
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            lonNumber = CLng(b1)
            DatTiles(i).speed = lonNumber
            If lonNumber = 0 Then
              DatTiles(i).blocking = True
            End If

        Else
            Get fn, , b1
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            lonNumber = CLng(b1)
            DatTiles(i).speed = lonNumber
            If lonNumber = 0 Then
              DatTiles(i).blocking = True
                        
            End If
            Get fn, , b2 'ignore next opt byte
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b2)
            #End If
            
        End If
      Case &H1 ' UNMODIFIED
        
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

      Case &H8 ' UNMODIFIED
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' UNMODIFIED
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' UNMODIFIED
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' UNMODIFIED
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' UNMODIFIED
        ' is blocking
        
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' UNMODIFIED
        ' not moveable
                 
        DatTiles(i).notMoveable = True
      Case &HE ' UNMODIFIED
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF ' UNMODIFIED
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' UNMODIFIED
        ' pickupable / equipable
        DatTiles(i).pickupable = True
    
      Case &H11 ' UNMODIFIED
        ' can see what is under (ladder holes, stairs holes etc)


      Case &H12 ' UNMODIFIED
        ' action posible
      Case &H13 ' UNMODIFIED
        'walls 2 types of them same material (total 4 pairs)
      Case &H14 ' UNMODIFIED
         ' unknown
      Case &H15 ' NEW?? - UNTESTED
         ' unknown
       
      Case &H16 ' used to be &H15 - ' NEW - UNTESTED

        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H17 ' NEW?? - UNTESTED
         ' unknown
      Case &H18 '  used to be &H17 - ' NEW - UNTESTED
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
        
      Case &H19 ' used to be &H18 - ' NEW - UNTESTED
        ' unknown

        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1A ' used to be &H19 - ' NEW - UNTESTED
            
      
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1B ' used to be &H1A - ' NEW - UNTESTED
        ' corpses that don't decay
        DatTiles(i).canDecay = False
        
      Case &H1C ' used to be &H1B - ' NEW - UNTESTED
        'wall items
        
      Case &H1D ' used to be &H1C - ' NEW - UNTESTED
        
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If

        
        
        
        
        
        ' xxxxxxxx



      Case &H1E ' used to be &H1D - ' NEW - UNTESTED
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2

        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          
          debugByte = optByte
          Debug.Print "Tile loader found unexpected properties at " & GoodHex(optByte) & ": " & GoodHex(optbyte2)
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
        
      Case &H1F ' used to be &H1E - ' NEW - UNTESTED
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
        
      Case &H20  ' used to be &H1F - ' NEW - UNTESTED
        'new flag since tibia 8.57
        
      Case &H21 ' used to be &H20 - ' NEW - UNTESTED
        '  body restriction
        ' 0 two handed
        ' 1 helmet
        ' 2 amulet
        ' 3 backpack<
        ' 4 armor
        ' 5 shield
        ' 6 weapon
        ' 7 legs
        ' 8 boots
        ' 9 ring
        ' 10 belt
        ' 11 purse
      
      
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
 
  
 

      Case &H22 ' used to be &H21 - ' NEW - UNTESTED
        Get fn, , b1 ' item group (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' item group (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' size of text (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' size of text (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
        tmpSize = GetTheLong(b1, b2)
        tmpName = ""
        For tmpI = 1 To tmpSize
            Get fn, , b1 ' size of text
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            tmpName = tmpName & Chr(b1)
        Next tmpI
        DatTiles(i).haveName = True
        DatTiles(i).itemName = tmpName
        #If TileDebug = 1 Then
          tileLog = tileLog & " (" & tmpName & ")"
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
        
      Case &H23 ' NEW since 10.21
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
   
      Case &HFE
        ' unknown new flag since tibia 10.21
        
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        If (optByte = &H20) Or (optByte = &H21) Or (optByte = &H22) Or (optByte = &H23) Or (optByte = &HFE) Then
        
        Else
          LogOnFile "tibiadatdebug.txt", "WARNING AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
        End If
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
    ' layers
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
    'PatternWidth
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
    'PatternHeight
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
    'PatternDepth
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
    'Phases
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If
    addToSkipCount = 0
    If lRare > &H1 Then
          DatTiles(i).haveExtraByte2 = True ' UNKNOWN , TEST
          ' new since Tibia 10.5
          addToSkipCount = 6 + (8 * lRare)
    End If
    If DatTiles(i).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then

      LogOnFile "tibiadatdebug.txt", tileLog2 & vbCrLf

    #End If
    ' NEW since Tibia 9.6: double size for graphic item references
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 4) + addToSkipCount
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile10 = -5 ' unexpected overflow
      Exit Function
    End If
   ' skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 4) + addToSkipCount 'size = old formulae x lRare
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If

  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile10 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile10 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile10 = -4 ' bad format or wrong version of given tibia.dat
End Function




' for tibia 10.58 and higher
Public Function LoadDatFile11(ByVal tibiadathere As String) As Integer
  Dim addToSkipCount As Long
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileLog As String
  Dim tileLog2 As String
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  Dim tmpSize As Long
  Dim tmpI As Long
  Dim tmpName As String
  Dim limit_ITEM_COUNT As Long
  Dim limit_OUTFIT_COUNT As Long
  Dim limit_EFFECT_COUNT As Long
  Dim limit_DISTANCE_COUNT As Long
  Dim dat_version As Long
  
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  res = 0
  tileOnDebug = 99999 ' last debug done at tile 238

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  #If TileDebug Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your tibia.dat file :"
  #End If
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Open tibiadathere For Binary As fn
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2

  'Debug.Print GoodHex(b1)
  Get fn, , b3
  'Debug.Print GoodHex(b1)
  Get fn, , b4
 ' Debug.Print GoodHex(b1)

 dat_version = FourBytesLong(b1, b2, b3, b4)
 If (dat_version < 1412240103) Then
    LoadDatFile11 = -2
 End If
 ' tibia 10.58 = 1412240103
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2
  'Debug.Print GoodHex(b2)
  limit_ITEM_COUNT = GetTheLong(b1, b2)
  
  Get fn, , b1
 ' Debug.Print GoodHex(b1)
  Get fn, , b2
 ' Debug.Print GoodHex(b2)
  limit_OUTFIT_COUNT = GetTheLong(b1, b2)
  Get fn, , b1
  'Debug.Print GoodHex(b1)

  Get fn, , b2
 ' Debug.Print GoodHex(b2)
   limit_EFFECT_COUNT = GetTheLong(b1, b2)
  Get fn, , b1
  'Debug.Print GoodHex(b1)
  
  Get fn, , b2
 ' Debug.Print GoodHex(b2)
   limit_DISTANCE_COUNT = GetTheLong(b1, b2)
  Do

    #If TileDebug = 1 Then
      tileLog = "tile #" & CStr(i) & ":"
    #End If
    Get fn, , optByte
    ' analyze all option Bytes until we read the byte &HFF
    ' note that some options are ignored
    ' and the meaning of some bytes are still unknown
    ' however this will get enough info for most purposes
    While (optByte <> &HFF) And Not EOF(fn)

      #If TileDebug = 1 Then
        tileLog = tileLog & " <" & GoodHex(optByte) & ">"
      #End If
      Select Case optByte
      Case &H0
        'is groundtile
        DatTiles(i).groundtile = True
        If ((TibiaVersionLong >= 1058) And (i = 21505)) Then
          ' rare case: only skip 1
            Get fn, , b1
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            lonNumber = CLng(b1)
            DatTiles(i).speed = lonNumber
            If lonNumber = 0 Then
              DatTiles(i).blocking = True
            End If

        Else
            Get fn, , b1
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            Get fn, , b2
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b2)
            #End If
            lonNumber = GetTheLong(b1, b2)
            DatTiles(i).speed = lonNumber
            If lonNumber = 0 Then
              DatTiles(i).blocking = True
                        
            End If
      
            
        End If
      Case &H1 ' UNMODIFIED
        
        ' new property : alwaysOnTop of higher priority
        DatTiles(i).moreAlwaysOnTop = True
      Case &H2 ' UNMODIFIED
        'always on top
        DatTiles(i).alwaysOnTop = True
      Case &H3 ' UNMODIFIED
        ' can walk through (open doors, arces ...)
        DatTiles(i).canWalkThrough = True
        DatTiles(i).alwaysOnTop = True
      Case &H4 ' UNMODIFIED
        ' is container
        DatTiles(i).iscontainer = True
      Case &H5 ' UNMODIFIED
        ' is stackable
        DatTiles(i).stackable = True
      Case &H6 ' UNMODIFIED
        ' is useable
        DatTiles(i).useable = True
      Case &H7 ' UNMODIFIED
        DatTiles(i).usable2 = True ' deleted since tibia 8.6 ?
        'DatTiles(i).multiCharge = True ' deleted since tibia 8.6 ?

      Case &H8 ' UNMODIFIED
        ' writtable objects
        DatTiles(i).RWInfo = 3 ' can writen + can be read
        Get fn, , b1 ' max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' max number of  newlines ? 0, 2, 4, 7
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
     Case &H9 ' UNMODIFIED
        ' writtable objects that can't be edited
        DatTiles(i).RWInfo = 1 ' can be read only
        Get fn, , b1 'always 0 max characters that can be written in it (0 unlimited)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 4 max number of  newlines ?
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &HA ' UNMODIFIED
        ' is fluid container
       DatTiles(i).fluidcontainer = True
      Case &HB ' UNMODIFIED
        ' multitype
        DatTiles(i).multitype = True ' DELETED ON TIBIA 8.6
      Case &HC ' UNMODIFIED
        ' is blocking
        
        DatTiles(i).blocking = True
        

        
        
      Case &HD ' UNMODIFIED
        ' not moveable
                 
        DatTiles(i).notMoveable = True
      Case &HE ' UNMODIFIED
        ' block missiles
        DatTiles(i).blockingProjectile = True
      Case &HF ' UNMODIFIED
        ' Slight obstacle (include fields and certain boxes)
        ' I prefer to don't consider a generic obstable and
        ' do special cases for fields and ignore the boxes
      Case &H10 ' UNMODIFIED
        ' pickupable / equipable
        DatTiles(i).pickupable = True
    
      Case &H11 ' UNMODIFIED
        ' can see what is under (ladder holes, stairs holes etc)


      Case &H12 ' UNMODIFIED
        ' action posible
      Case &H13 ' UNMODIFIED
        'walls 2 types of them same material (total 4 pairs)
      Case &H14 ' UNMODIFIED
         ' unknown
      Case &H15 ' NEW?? - UNTESTED
         ' unknown
       
      Case &H16 ' used to be &H15 - ' NEW - UNTESTED

        ' makes light -- skip bytes
        Get fn, , b1 ' number of tiles around

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1 ' = 215 for items , =208 for non items

        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
      Case &H17 ' NEW?? - UNTESTED
         ' unknown
      Case &H18 '  used to be &H17 - ' NEW - UNTESTED
        ' stairs to down
        DatTiles(i).floorChangeDOWN = True
        
      Case &H19 ' used to be &H18 - ' NEW - UNTESTED
        ' unknown

        Get fn, , b1 ' 4 bytes of extra info
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        Get fn, , b1
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1A ' used to be &H19 - ' NEW - UNTESTED
            
      
        ' mostly blocking items, but also items that can pile up in level (boxes, chairs etc)
        DatTiles(i).blockpickupable = False
        Get fn, , b1 ' always 8
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' always 0
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
      Case &H1B ' used to be &H1A - ' NEW - UNTESTED
        ' corpses that don't decay
        DatTiles(i).canDecay = False
        
      Case &H1C ' used to be &H1B - ' NEW - UNTESTED
        'wall items
        
      Case &H1D ' used to be &H1C - ' NEW - UNTESTED
        
        ' for minimap drawing
        Get fn, , b1 ' 2 bytes for colour
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If

        
        
        
        
        
        ' xxxxxxxx



      Case &H1E ' used to be &H1D - ' NEW - UNTESTED
        ' line spot ...
        Get fn, , optbyte2 '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(optbyte2)
        #End If
        Select Case optbyte2

        Case &H4C
          'ladders
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRightClick = True
        Case &H4D
          'crate - trapdor?
          DatTiles(i).requireRightClick = True
        Case &H4E
          'rope spot?
          DatTiles(i).floorChangeUP = True
          DatTiles(i).requireRope = True
        Case &H4F
          'switch
        Case &H50
          'doors
          DatTiles(i).isDoor = True
        Case &H51
          'doors with locks
          DatTiles(i).isDoorWithLock = True
        Case &H52
          'stairs to up floor
          DatTiles(i).floorChangeUP = True
        Case &H53
          'mailbox
        Case &H54
          'depot
          DatTiles(i).isDepot = True
        Case &H55
          'trash
        Case &H56
         'hole
          DatTiles(i).floorChangeDOWN = True
          DatTiles(i).requireShovel = True
          DatTiles(i).alwaysOnTop = True
          DatTiles(i).multitype = False
        Case &H57
          'items with special description?
        Case &H58
          'writtable
          DatTiles(i).RWInfo = 1 ' read only
        Case Else
          ' should not happen
          
          debugByte = optByte
          Debug.Print "Tile loader found unexpected properties at " & GoodHex(optByte) & ": " & GoodHex(optbyte2)
          #If TileDebug = 1 Then
            tileLog = tileLog & " " & GoodHex(b1) & "!"
          #End If
        End Select 'optbyte2
        Get fn, , b1 ' always value 4
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
        
      Case &H1F ' used to be &H1E - ' NEW - UNTESTED
        ' ground tiles that don't cause level change
        DatTiles(i).noFloorChange = True
        
      Case &H20  ' used to be &H1F - ' NEW - UNTESTED
        'new flag since tibia 8.57
        
      Case &H21 ' used to be &H20 - ' NEW - UNTESTED
        '  body restriction
        ' 0 two handed
        ' 1 helmet
        ' 2 amulet
        ' 3 backpack<
        ' 4 armor
        ' 5 shield
        ' 6 weapon
        ' 7 legs
        ' 8 boots
        ' 9 ring
        ' 10 belt
        ' 11 purse
      
      
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
 
  
 

      Case &H22 ' used to be &H21 - ' NEW - UNTESTED
        Get fn, , b1 ' item group (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' item group (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' size of text (byte 1)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b2 ' size of text (byte 2)
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b2)
        #End If
        
        tmpSize = GetTheLong(b1, b2)
        tmpName = ""
        For tmpI = 1 To tmpSize
            Get fn, , b1 ' size of text
            #If TileDebug = 1 Then
              tileLog = tileLog & " " & GoodHex(b1)
            #End If
            tmpName = tmpName & Chr(b1)
        Next tmpI
        DatTiles(i).haveName = True
        DatTiles(i).itemName = tmpName
        #If TileDebug = 1 Then
          tileLog = tileLog & " (" & tmpName & ")"
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        
        
      Case &H23 ' NEW since 10.21
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
        Get fn, , b1 ' unknown meaning
        #If TileDebug = 1 Then
          tileLog = tileLog & " " & GoodHex(b1)
        #End If
   
      Case &HFE
        ' unknown new flag since tibia 10.21
        
      Case Else
        ' should not happen
        debugByte = optByte
        #If TileDebug = 1 Then
          tileLog = tileLog & "?"
        #End If
      End Select 'optbyte
      Get fn, , nextB 'next optByte
      #If TileDebug = 1 Then
      If nextB <= optByte Then
        If (optByte = &H20) Or (optByte = &H21) Or (optByte = &H22) Or (optByte = &H23) Or (optByte = &HFE) Then
        
        Else
          LogOnFile "tibiadatdebug.txt", "WARNING AT tileID #" & CStr(i) & " : " & GoodHex(nextB) & " <= " & GoodHex(optByte)
        End If
      End If
      #End If
      optByte = nextB
    Wend
endAnalyze:
    #If TileDebug = 1 Then
      tileLog = tileLog & " " & GoodHex(optByte) & " OK"
      LogOnFile "tibiadatdebug.txt", tileLog
      If tileOnDebug = i Then
        Debug.Print tileLog
      End If
    #End If

    ' some flags can be made by a combination of existing flags
    If DatTiles(i).stackable = True Or DatTiles(i).multitype = True Or _
      DatTiles(i).fluidcontainer = True Then
      DatTiles(i).haveExtraByte = True
    End If
    
    If DatTiles(i).multiCharge = True Then
      DatTiles(i).haveExtraByte = True
    End If

    If DatTiles(i).alwaysOnTop = True Then
      DatTiles(i).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(i).moreAlwaysOnTop = True Then
      DatTiles(i).alwaysOnTop = True
      DatTiles(i).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case i
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(i).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(i).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If i = tileID_waterWithFish Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If i = tileID_waterEmpty Then
      DatTiles(i).isWater = True
    End If
    If TibiaVersionLong >= 781 Then
        If i = tileID_blockingBox Then
            DatTiles(i).blocking = True
        End If
    End If
    
    If TibiaVersionLong >= 760 Then

    If (i >= tileID_waterWithFish) And (i <= tileID_waterWithFishEnd) Then
      DatTiles(i).isWater = True
      DatTiles(i).haveFish = True
    End If
    If (i >= tileID_waterEmpty) And (i <= tileID_waterEmptyEnd) Then
      DatTiles(i).isWater = True
    End If

    End If
    ' food, for autoeater
    If i >= tileID_firstFoodTileID And i <= tileID_lastFoodTileID Then
      DatTiles(i).isFood = True
    End If
    If (i >= tileID_firstMushroomTileID) And (i <= tileID_lastMushroomTileID) Then
      DatTiles(i).isFood = True
    End If
    
    Select Case i ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(i).isFood = True
    End Select
    
    If (i >= 8010) And (i <= 8020) Then ' special food
      DatTiles(i).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If i >= tileID_firstFieldRangeStart And i <= tileID_firstFieldRangeEnd Then
      DatTiles(i).isField = True
    End If
    If (i >= tileID_secondFieldRangeStart) And (i <= tileID_secondFieldRangeEnd) Then
      DatTiles(i).isField = True
    End If
    Select Case i
    Case tileID_campFire1, tileID_campFire2
      DatTiles(i).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(i).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If i = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(i).floorChangeUP = True
    End If
    If i = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(i).pickupable = False
    End If
    '[/CUSTOM FLAGS FOR BLACKDPROXY]
    
    ' options zone done for this tile
    ' now we get info about the graph of the tile...
    ' but as we are not interested on it, just skip enough bytes
    Get fn, , b1
    
    lWidth = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = "[lWidth=" & GoodHex(b1) & "]"
    #End If
    
    
    Get fn, , b1
    lHeight = CLng(b1)
    #If TileDebug = 1 Then
      tileLog2 = tileLog2 & " [lHeight=" & GoodHex(b1) & "]"
    #End If
    If (lWidth > 1) Or (lHeight > 1) Then
      'skip 1 byte
      Get fn, , b1
      #If TileDebug = 1 Then
        tileLog2 = tileLog2 & " [SkipByte=" & GoodHex(b1) & "]"
      #End If
    End If
    

    Get fn, , b1
    lBlendframes = CLng(b1)
    #If TileDebug = 1 Then
    ' layers
      tileLog2 = tileLog2 & " [lBlendframes=" & GoodHex(b1) & "]"
    #End If
    
    Get fn, , b1
    lXdiv = CLng(b1)
    #If TileDebug = 1 Then
    'PatternWidth
      tileLog2 = tileLog2 & " [lXdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lYdiv = CLng(b1)
    #If TileDebug = 1 Then
    'PatternHeight
      tileLog2 = tileLog2 & " [lYdiv=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lAnimcount = CLng(b1)
    #If TileDebug = 1 Then
    'PatternDepth
      tileLog2 = tileLog2 & " [lAnimcount=" & GoodHex(b1) & "]"
    #End If

    Get fn, , b1
    lRare = CLng(b1)
    #If TileDebug = 1 Then
    'Phases
      tileLog2 = tileLog2 & " [lRare=" & GoodHex(b1) & "]"
    #End If
    addToSkipCount = 0
    If lRare > &H1 Then
          DatTiles(i).haveExtraByte2 = True ' UNKNOWN , TEST
          ' new since Tibia 10.5
          addToSkipCount = 6 + (8 * lRare)
    End If
    If DatTiles(i).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    If DatTiles(i).haveExtraByte2 = True Then
      DatTiles(i).totalExtraBytes = DatTiles(i).totalExtraBytes + 1
    End If
    #If TileDebug = 1 Then

      LogOnFile "tibiadatdebug.txt", tileLog2 & vbCrLf

    #End If
    ' NEW since Tibia 9.6: double size for graphic item references
    skipcount = protectedMult(lWidth, lHeight, lBlendframes, lXdiv, lYdiv, lAnimcount, lRare, 4) + addToSkipCount
    If skipcount = -1 Then
      DBGtileError = "The function failed exactly because this overflow: " & vbCrLf & _
       CStr(lWidth) & " * " & CStr(lHeight) & " * " & CStr(lBlendframes) & " * " & CStr(lXdiv) & " * " & CStr(lYdiv) & " * " & CStr(lAnimcount) & " * " & CStr(lRare) & " * 2" & _
       vbCrLf & "tibia.dat path = tibiadatHere"
      LoadDatFile11 = -5 ' unexpected overflow
      Exit Function
    End If
   ' skipcount = (lWidth * lHeight * lBlendframes * lXdiv * lYdiv * lAnimcount * lRare * 4) + addToSkipCount 'size = old formulae x lRare
    
    
    #If TileDebug = 1 Then
    ' if you are curious about graphic data of certain tile, then just set tileOnDebug=your desired tileID
        If i = tileOnDebug Then
          tileLog2 = "Debug graphic part for tile # " & CStr(i) & " : " & tileLog2 & " : "
          For j = 1 To skipcount
            Get fn, , b1
            tileLog2 = tileLog2 & " " & GoodHex(b1)
          Next j
          LogOnFile "tibiadatdebug.txt", tileLog2
          Debug.Print tileLog2
        Else
            For expI = 1 To skipcount
                Get fn, , bTmp
            Next expI
        End If
    #Else
        For expI = 1 To skipcount
            Get fn, , bTmp
        Next expI
    #End If

    i = i + 1
    If i > MAXDATTILES Then
      res = -3  ' need to increase const MAXDATTILES
      GoTo endF
    End If
    If i > limit_ITEM_COUNT Then
      Exit Do
    End If
  Loop Until EOF(fn)
  ' Close the file
  Close fn
  ' last one is not a valid tile id! -> i - 1
  highestDatTile = i - 1
  If highestDatTile < 1 Then
    LoadDatFile11 = -1
    Exit Function
  End If
endF:
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
  ' Debug.Print tileLog
  'Debug.Print highestDatTile
  
  LoadDatFile11 = res
  Exit Function
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFile11 = -4 ' bad format or wrong version of given tibia.dat
End Function

Public Function GetOnlyPath(ByVal pathname As String) As String
    Dim pathstr As String
  Dim posn As Long
    posn = InStrRev(pathname, "\")
    If posn > 0 Then
        pathstr = Left$(pathname, posn)
    Else
        pathstr = ""
    End If
    GetOnlyPath = pathstr
End Function

Private Function movePos(ByRef pos As Long, ByRef positions As Long, _
Optional ByVal addVBCRLF As Boolean = False, _
Optional ByVal markPosition0 As Long = -1, _
Optional ByVal markPosition1 As Long = -1, _
Optional ByVal markPosition2 As Long = -1)
    #If TileDebug = 1 Then
      If (atGraphicPart = False) Or (debugGraphicPart And atGraphicPart) Then
        Dim i As Integer
        For i = 0 To positions - 1
          If (i = markPosition0) Or (i = markPosition1) Or (i = markPosition2) Then
            tileLog = tileLog & " <" & GoodHex(byteArray(pos + i)) & ">"
          Else
            tileLog = tileLog & " " & GoodHex(byteArray(pos + i))
          End If
        Next i
        If (addVBCRLF) Then
            tileLog = tileLog & vbCrLf
        End If
      End If
    #End If
    pos = pos + positions
End Function

Public Sub getProtobufBytesFrom4BYTESNUMBER(ByRef arr() As Byte, ByVal Number As Long)
    Const cte1 As Long = 128
    Const cte2 As Long = 16384 ' 128 * 128
    Const cte3 As Long = 2097152 ' 128 * 128 * 128
    Const cte4 As Long = 268435456 ' 128 * 128 * 128 * 128
    If Number < cte1 Then
        ReDim arr(0)
        arr(0) = CByte(Number)
    ElseIf Number < cte2 Then
        ReDim arr(1)
        arr(1) = CByte(Number \ cte1)
        Number = Number - (arr(1) * cte1)
        arr(0) = CByte(cte1 + Number)
    ElseIf Number < cte3 Then
        ReDim arr(2)
        arr(2) = CByte(Number \ cte2)
        Number = Number - (arr(2) * cte2)
        arr(1) = CByte(cte1 + (Number \ cte1))
        Number = Number - ((arr(1) - cte1) * cte1)
        arr(0) = CByte(cte1 + Number)
    ElseIf Number < cte4 Then
        ReDim arr(3)
        arr(3) = CByte(Number \ cte3)
        Number = Number - (arr(3) * cte3)
        arr(2) = CByte(cte1 + (Number \ cte2))
        Number = Number - ((arr(2) - cte1) * cte2)
        arr(1) = CByte(cte1 + (Number \ cte1))
        Number = Number - ((arr(1) - cte1) * cte1)
        arr(0) = CByte(cte1 + Number)
    Else
        ReDim arr(4)
        arr(4) = CByte(Number \ cte4)
        Number = Number - (arr(4) * cte4)
        arr(3) = CByte(cte1 + (Number \ cte3))
        Number = Number - ((arr(3) - cte1) * cte3)
        arr(2) = CByte(cte1 + (Number \ cte2))
        Number = Number - ((arr(2) - cte1) * cte2)
        arr(1) = CByte(cte1 + (Number \ cte1))
        Number = Number - ((arr(1) - cte1) * cte1)
        arr(0) = CByte(cte1 + Number)
    End If
    Debug.Print arrayToString(arr)
End Sub

Public Sub testGetProto(ByVal Number As Long)
    Dim arr() As Byte
    getProtobufBytesFrom4BYTESNUMBER arr, Number
End Sub

Public Function readProtobufNumber(ByRef arr() As Byte, ByVal pos As Long, ByRef positionsRead As Long) As Long
     Dim num As Long
     Dim b1 As Byte
     Dim b2 As Byte
     Dim b3 As Byte
     Dim b4 As Byte
     Dim b5 As Byte
     On Error GoTo goterr
     b1 = arr(pos)
     num = CLng(b1)
     If (b1 < &H80) Then
        positionsRead = 1
        num = CLng(b1)
     Else
        b2 = arr(pos + 1)
        num = num + (128 * (CLng(b2) - 1))
        If (b2 < &H80) Then
            positionsRead = 2
        Else
            b3 = arr(pos + 2)
            num = num + (16384 * (CLng(b3) - 1)) ' 16384 = 128*128
            If (b3 < &H80) Then
                positionsRead = 3
            Else
                b4 = arr(pos + 3)
                num = num + (2097152 * (CLng(b4) - 1)) ' 2097152 = 16384*128
                If (b4 < &H80) Then
                    positionsRead = 4
                Else
                    b5 = arr(pos + 4)
              
                    If (arr(pos + 5) = &HFF) Then
                       positionsRead = 10
                       num = -1
                    Else
                       num = num + (268435456 * (CLng(b5) - 1)) ' 268435456 = 2097152*128
                       positionsRead = 5
                    End If
                End If
            End If
        End If
     End If
     readProtobufNumber = num
     Exit Function
goterr:
     readProtobufNumber = -2
End Function

' Example: testProtobufNumber ("80 80 01")
Public Sub testProtobufNumber(ByVal str As String)
    Dim arr() As String
    Dim arrb() As Byte
    Dim res As Long
    Dim i As Integer
    Dim positions As Long
    Dim tam As Long
    str = Trim$(str)
    arr = Split(str, " ")
    tam = UBound(arr)
    ReDim arrb(tam)
    For i = 0 To UBound(arr)
        arrb(i) = CByte(CLng("&H" & arr(i)))
    Next i
    res = readProtobufNumber(arrb, 0, positions)
    Debug.Print str & " = " & CStr(res) & " (used " & CStr(positions) & " positions)"
End Sub

Public Sub AddBlackdProxyFlags()
Dim i As Integer
 For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToUpFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToUpFloor(i)).floorChangeUP = True
    End If
  Next i
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireRope(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireRope(i)).floorChangeUP = True
      DatTiles(AditionalRequireRope(i)).requireRope = True
    End If
  Next i
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalRequireShovel(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalRequireShovel(i)).floorChangeDOWN = True
      DatTiles(AditionalRequireShovel(i)).requireShovel = True
      DatTiles(AditionalRequireShovel(i)).alwaysOnTop = True
      DatTiles(AditionalRequireShovel(i)).multitype = False
    End If
  Next i
  
  
  For i = 0 To MAXTILEIDLISTSIZE
    If (AditionalStairsToDownFloor(i) = 0) Then
      Exit For
    Else
      DatTiles(AditionalStairsToDownFloor(i)).floorChangeDOWN = True
    End If
  Next i
End Sub

' Expects a path for tibia 11.0+ dat and catalog-content.json
' Also allows path+tibia.dat of tibia 10.98
Public Function LoadDatFileQ(ByVal tibiadathere As String) As Integer
Const expectedVersion As Long = 17223
  If (LCase(Right$(tibiadathere, 9)) = "tibia.dat") Then
    LoadDatFileQ = LoadDatFile11(tibiadathere)
    Exit Function
  End If
  #If FinalMode Then
    On Error GoTo badErr
  #End If
  GetInfoOfNewDatFile TibiaExePathWITHTIBIADAT, CurrentTibiaDatVERSION, CurrentTibiaDatFILE
  If (CurrentTibiaDatVERSION = -1) Then
    LoadDatFileQ = -1
    Exit Function
  End If
  ' Debug.Print "Current .dat version=" & CurrentTibiaDatVERSION

  tibiadathere = tibiadathere & CurrentTibiaDatFILE
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  If fso.FileExists(tibiadathere) = False Then
    LoadDatFileQ = -1
    Exit Function
  End If
  Dim addToSkipCount As Long
  Dim res As Integer
  Dim i As Long
  Dim j As Long
  Dim fn As Integer
  Dim optByte As Byte
  Dim optbyte2 As Byte
  Dim b1 As Byte
  Dim b2 As Byte
  Dim b3 As Byte
  Dim b4 As Byte
  Dim a As String
  Dim lonNumber As Long
  Dim lWidth  As Long
  Dim lHeight As Long
  Dim lBlendframes As Long
  Dim lXdiv As Long
  Dim lYdiv As Long
  Dim lAnimcount As Long
  Dim lRare As Long
  Dim skipcount As Long
  Dim debugByte As Byte
  Dim tileOnDebug As Long
  Dim nextB As Byte
  Dim expI As Long
  Dim bTmp As Byte
  Dim tmpSize As Long
  Dim tmpI As Long
  Dim tmpName As String
  Dim limit_ITEM_COUNT As Long
  Dim limit_OUTFIT_COUNT As Long
  Dim limit_EFFECT_COUNT As Long
  Dim limit_DISTANCE_COUNT As Long
  Dim dat_version As Long
  

  res = 0
  tileOnDebug = 99999 ' last debug done at tile 238

  ' init the array of tiles with default values
  For i = 0 To MAXDATTILES
    DatTiles(i).iscontainer = False
    DatTiles(i).RWInfo = 0
    DatTiles(i).fluidcontainer = False
    DatTiles(i).stackable = False
    DatTiles(i).multitype = False
    DatTiles(i).useable = False
    DatTiles(i).notMoveable = False
    DatTiles(i).alwaysOnTop = False
    DatTiles(i).groundtile = False
    DatTiles(i).blocking = False
    DatTiles(i).pickupable = False
    DatTiles(i).blockingProjectile = False
    DatTiles(i).canWalkThrough = False
    DatTiles(i).noFloorChange = False
    DatTiles(i).blockpickupable = True
    DatTiles(i).isDoor = False
    DatTiles(i).isDoorWithLock = False
    DatTiles(i).speed = 0
    DatTiles(i).canDecay = True
    DatTiles(i).haveExtraByte = False 'custom flag
    DatTiles(i).haveExtraByte2 = False 'custom flag
    DatTiles(i).totalExtraBytes = 0 'custom flag
    DatTiles(i).floorChangeUP = False 'custom flag
    DatTiles(i).floorChangeDOWN = False 'custom flag
    DatTiles(i).requireRightClick = False 'custom flag
    DatTiles(i).requireRope = False 'custom flag
    DatTiles(i).requireShovel = False 'custom flag
    DatTiles(i).isWater = False ' custom flag
 
    DatTiles(i).stackPriority = 1

    DatTiles(i).haveFish = False
    DatTiles(i).isFood = False
    DatTiles(i).isField = False
    DatTiles(i).isDepot = False
    DatTiles(i).moreAlwaysOnTop = False
    DatTiles(i).usable2 = False
    DatTiles(i).multiCharge = False
    DatTiles(i).haveName = False
    DatTiles(i).itemName = ""
  Next i
  DatTiles(0).stackPriority = 0

  DatTiles(97).stackPriority = 2
  DatTiles(98).stackPriority = 2
  DatTiles(99).stackPriority = 2

  DatTiles(97).blocking = True
  DatTiles(98).blocking = True
  DatTiles(99).blocking = True
  i = 100 ' i = tileID
  
  
  
  fn = FreeFile
  ' Open the file tibia.dat for binary access
  ' it look for it in the same path than this program (App.Path)

  Dim lastByte As Long

  Dim pos As Long

  Dim nextMarkerWillStartNewItem As Boolean
  nextMarkerWillStartNewItem = True
  pos = 0
  ' read full file at once
  Open tibiadathere For Binary Access Read As fn
    lastByte = LOF(fn) - 1
    If (lastByte > 0) Then
        ReDim byteArray(lastByte)
        Get fn, , byteArray
    End If
  Close fn

  tileLog = ""
  debugGraphicPart = False
 
  Dim properties_header_size As Long
  Dim graphicPart_start As Long
  Dim graphicPart_size As Long
  Dim graphicPart_end As Long
  Dim graphicPart2_size As Long
  Dim properties_start As Long
  Dim properties_end As Long
  Dim properties_size_without_header As Long
  Dim nextByte As Byte
  Dim totalSkip As Long
  Dim totalBytesHere As Long
  Dim positionsRead As Long
  Dim currentTile As Long
  Dim lNew As Long
  Dim lPrev1 As Long
  Dim lPrev2 As Long
  Dim expectedBytes As Long
  Dim rareBytes As Long
  Dim gDebug As String
  Dim gTileCount As Long
  Dim gproperties_end As Long
  Dim realGstart As Long
  Dim expectedGTiles As Long
  #If TileDebug = 1 Then
    OverwriteOnFile "tibiadatdebug.txt", "Here is what Blackd Proxy could read in your .dat file :"
  #End If
  
  Do While pos < lastByte
    If (Not (byteArray(pos) = &HA)) Then
        If (currentTile < 23725) Then
             #If TileDebug = 1 Then
                LoadDatFileQ = -1
                tileLog = tileLog & vbCrLf & "File continues:"
                movePos pos, 800
                LogOnFile "tibiadatdebug.txt", tileLog
             #End If
             LoadDatFileQ = -1
             Exit Function
        Else
            ' we already finished reading the usefull part for us
            highestDatTile = currentTile
            Exit Do
        End If
    End If
    atGraphicPart = True
    graphicPart_start = pos

    totalBytesHere = readProtobufNumber(byteArray, pos + 1, positionsRead)
    properties_end = pos + positionsRead + totalBytesHere
    pos = pos + 2 + positionsRead
    currentTile = readProtobufNumber(byteArray, pos, positionsRead)
    pos = pos + 1 + positionsRead
    graphicPart_size = readProtobufNumber(byteArray, pos, positionsRead)
    pos = pos + 1 + positionsRead
    lPrev1 = readProtobufNumber(byteArray, pos, positionsRead)
    pos = pos + 1 + positionsRead
    lPrev2 = readProtobufNumber(byteArray, pos, positionsRead)
    pos = pos + 1 + positionsRead
    graphicPart2_size = readProtobufNumber(byteArray, pos, positionsRead)
    realGstart = pos + positionsRead
    
    properties_start = pos + positionsRead + graphicPart2_size
    graphicPart_end = properties_start - 1
    pos = pos + 1 + positionsRead
    lXdiv = readProtobufNumber(byteArray, pos, positionsRead)
    pos = pos + 1 + positionsRead
    lYdiv = readProtobufNumber(byteArray, pos, positionsRead)
    pos = pos + 1 + positionsRead
    lAnimcount = readProtobufNumber(byteArray, pos, positionsRead)
    pos = pos + 1 + positionsRead
    lNew = readProtobufNumber(byteArray, pos, positionsRead)
    expectedGTiles = lXdiv * lYdiv * lAnimcount
    expectedBytes = 10 + (expectedGTiles * 4)
 
'    If (currentTile = 134) Then
'    Debug.Print expectedBytes & "vs" & graphicPart2_size
'    Debug.Print "ok"
'    End If
    If (graphicPart2_size > expectedBytes + 2) Then
        lRare = 2 ' This is not the real value, however we only need to know that it is bigger than 1
    Else
        lRare = 1
    End If
    #If TileDebug = 1 Then
          tileLog = "tile #" & CStr(currentTile) & ":"
      
          gDebug = "[lXdiv=" & GoodHex(CByte(lXdiv)) & "] [lYdiv=" & GoodHex(CByte(lYdiv)) & "] [lAnimcount=" & GoodHex(CByte(lAnimcount)) & "]"
'          If (graphicPart2_size > expectedBytes) Then
'              gDebug = gDebug & " [lRare>01]"
'          End If
         ' gDebug = gDebug & " [g2=" & graphicPart2_size & "] [ex=" & expectedBytes & "]"
          ' Following cases should not happen, else write log about it:
          If lPrev1 <> 2 Then
            tileLog = tileLog & "lPrev1!!!!!!!!!!!" & vbCrLf
          End If
          If lPrev2 <> 2 Then
            tileLog = tileLog & "lPrev2!!!!!!!!!!!" & vbCrLf
          End If
          If lNew > 1 Then
            tileLog = tileLog & "lNew!!!!!!!!!!!" & vbCrLf
          End If
    #End If

    #If TileDebug = 1 Then
        If (currentTile = tileOnDebug) Then
            debugGraphicPart = True
        Else
            debugGraphicPart = False
        End If
    #End If
   ' debugGraphicPart = True ' uncomment this line to log full content


   ' pos = graphicPart_start
  '  movePos pos, graphicPart_end - graphicPart_start + 1, True
    pos = realGstart
    gTileCount = 0
    Do While pos <= graphicPart_end
            optByte = byteArray(pos)
            Select Case optByte
            Case &H1A
               movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H8
              movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H10
              movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H18
              movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H20
              movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H28
               gTileCount = gTileCount + 1
                 movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H32
               movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H38
               movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H40
              movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case &H4A
               movePos pos, 0, True
               movePos pos, 1, False, 0
               lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
               movePos pos, positionsRead
            Case Else
                ' Our parser needs an update
                #If TileDebug = 1 Then
                     tileLog = tileLog & vbCrLf & "GPARSER ERROR - Unexpected gtile: " & GoodHex(optByte) & vbCrLf
                     tileLog = tileLog & vbCrLf & "File continues:"
                     movePos pos, 100
                     LogOnFile "tibiadatdebug.txt", tileLog
                #End If
                LoadDatFileQ = -4 ' Unknown property
                Exit Function
            End Select
    Loop
    lRare = gTileCount / expectedGTiles
    gDebug = gDebug & " [lRare=" & GoodHex(CByte(lRare)) & "]"
    atGraphicPart = False
    
    ' Skip properties block header (It only contains number of bytes of real properties)
    properties_size_without_header = readProtobufNumber(byteArray, pos, positionsRead)
    properties_header_size = 1 + positionsRead
    movePos pos, properties_header_size
    
    #If ParseProperties = 0 Then
        movePos pos, properties_end - properties_start + 1 - properties_header_size, True
    #Else
        Do While pos <= properties_end
            optByte = byteArray(pos)
            Select Case optByte
            Case &HA
            ' OLD <00> "is ground tile"
            ' 0A 02 08 00
            ' 0A 03 08 A0 01
                DatTiles(currentTile).groundtile = True
                movePos pos, 3, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
                DatTiles(currentTile).speed = lonNumber
                If lonNumber = 0 Then
                    DatTiles(currentTile).blocking = True
                End If
            Case &H10
            ' OLD <01> "alwaysOnTop of higher priority"
            ' 10 01
                DatTiles(currentTile).moreAlwaysOnTop = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H18
            ' OLD <02> "always on top"
            ' 18 01
                DatTiles(currentTile).alwaysOnTop = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H20
            ' OLD <03> "can walk through" (open doors, arces ...)
            ' 20 01
                DatTiles(currentTile).canWalkThrough = True
                DatTiles(currentTile).alwaysOnTop = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H28
            ' OLD <04> "is container"
            ' 28 01
              
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                If DatTiles(currentTile).haveName = False Then
                  DatTiles(currentTile).iscontainer = True
                End If
                movePos pos, positionsRead
            Case &H30
            ' OLD <05> "is stackable"
            ' 30 01
            ' 30 32
            ' 30 8C 01
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                If lonNumber = 1 Then
                    DatTiles(currentTile).stackable = True
                End If
                movePos pos, positionsRead
            Case &H38
            ' OLD <FE> --- unknown meaning
            ' 38 01
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H40
            ' OLD <06> "is usable"
            ' 40 01
                DatTiles(currentTile).useable = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H48
            ' OLD <07> "is drinkable"
            ' 48 01
                DatTiles(currentTile).usable2 = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H52
            ' OLD <08> "writtable"
            ' 52 03 08 C8 01
                DatTiles(currentTile).RWInfo = 3
                movePos pos, 3, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H5A
            ' OLD <09> "read only"
            ' 5A 03 08 80 08
                DatTiles(currentTile).RWInfo = 1
                movePos pos, 3, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H60
            ' OLD <0B> "multitype"
            ' 60 01
                DatTiles(currentTile).multitype = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H68
            ' OLD <0C> "is blocking"
            ' 68 01
                DatTiles(currentTile).blocking = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H70
            ' OLD <0D> "not moveable" CONFIRMED
            ' 70 01
                DatTiles(currentTile).notMoveable = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H78
            ' OLD <0E> "block missiles"
            ' 78 01
                DatTiles(currentTile).blockingProjectile = True
                movePos pos, 1, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H80
            ' OLD <0F> AND <1F>
            ' 80 01 01 = 0F
            ' 80 02 01 = 1F
                optbyte2 = byteArray(pos + 1)
                Select Case optbyte2
                Case &H1
                     ' OLD <0F> "slight obstacle"
                Case &H2
                     ' OLD <1F> "ground tiles that don't cause level change"
                     DatTiles(currentTile).noFloorChange = True
                Case Else
                    Debug.Print "!!!!!"
                End Select
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H88
            ' old <10> AND <20>
            ' 88 01 01 = OLD <10>
            ' 88 02 01 = OLD <20>
                optbyte2 = byteArray(pos + 1)
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                Select Case optbyte2
                Case &H1
                     ' OLD <10> "pickupable"
                     DatTiles(currentTile).pickupable = True
                Case &H2
                     ' OLD <20> --- unknown meaning
                Case Else
                    Debug.Print "!!!!!"
                End Select
                movePos pos, positionsRead
            Case &H90
            ' OLD <11> "can see what is under (ladder holes, stairs holes etc)"
            ' 90 01 01
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H92
            ' OLD <21> "body restriction"
            ' 92 02 02 08 01
                movePos pos, 4, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H98
            ' old <0A> "fluid container"
            ' 98 01 01
                DatTiles(currentTile).fluidcontainer = True
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &H9A
            ' OLD <23> --- unknown meaning
            ' 9A 02 02 08 04
                movePos pos, 4, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HA0
            ' OLD <12> "action is possible"
            ' A0 01 01
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HA2
            ' OLD <22> "have name"
            ' A2 02 14 08 09 10 66 18 66 22 0C 77 68 69 74 65 20 66 6C 6F 77 65 72
            ' A2 02 19 08 09 10 80 01 18 80 01 22 0F 73 69 67 6E 65 64 20 63 6F 6E 74 72 61 63 74
                movePos pos, 6, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
                movePos pos, 1
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
                movePos pos, 1
                tmpSize = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
                tmpName = ""
                For tmpI = pos To pos + tmpSize - 1
                    tmpName = tmpName & Chr(byteArray(tmpI))
                Next tmpI
                DatTiles(currentTile).haveName = True
                DatTiles(currentTile).itemName = tmpName
                movePos pos, tmpSize
                #If TileDebug = 1 Then
                   tileLog = tileLog & " (" & tmpName & ")"
                #End If
            Case &HA8
            ' OLD <25> --- unknown meaning
            ' A8 02 01
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HAA
            ' OLD <14> --- unknown meaning
            ' AA 01 02 08 02
                movePos pos, 4, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HB0
            ' OLD <15> --- unknown meaning
            ' B0 01 01
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HBA
            ' OLD <16> "makes light"
            ' BA 01 05 08 03 10 9C 01
            ' BA 01 05 08 02 10 CE 01
            ' BA 01 04 08 04 10 23
                movePos pos, 6, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HC0
            ' old <17> --- unknown meaning
            ' C0 01 01
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HC8
            ' OLD <18> "stairs to lower floor"
            ' C8 01 01
                DatTiles(currentTile).floorChangeDOWN = True
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HDA
            ' OLD <1A> "block pickupable"
            ' DA 01 02 08 08
                DatTiles(currentTile).blockpickupable = False
                movePos pos, 5, , 0
            Case &HD2
            ' OLD <19> --- unknown meaning
            ' D2 01 04 08 08 10 08
                movePos pos, 7, , 0
            Case &HE0
            ' OLD <1B> "corpses that don't decay"
            ' E0 01 01
                DatTiles(currentTile).canDecay = False
                movePos pos, 2, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HF2
            ' OLD <1D> "color for minimap drawing"
            ' F2 01 03 08 81 01
                movePos pos, 4, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                movePos pos, positionsRead
            Case &HFA
            ' OLD <1E> "special item"
            ' FA 01 03 08 CE 08
                movePos pos, 4, , 0
                lonNumber = readProtobufNumber(byteArray, pos, positionsRead)
                optbyte2 = LowByteOfLong(lonNumber)
                '86 -> openable holes, 77-> can be used to go down, 76 can be used to go up, 82 -> stairs up, 79 switch,...
                Select Case optbyte2
                Case &H4C
                  'ladders
                  DatTiles(currentTile).floorChangeUP = True
                  DatTiles(currentTile).requireRightClick = True
                Case &H4D
                  'crate - trapdoor
                  DatTiles(currentTile).requireRightClick = True
                Case &H4E
                  'rope spot?
                  DatTiles(currentTile).floorChangeUP = True
                  DatTiles(currentTile).requireRope = True
                Case &H4F
                  'switch
                Case &H50
                  'doors
                  DatTiles(currentTile).isDoor = True
                Case &H51
                  'doors with locks
                  DatTiles(currentTile).isDoorWithLock = True
                Case &H52
                  'stairs to up floor
                  DatTiles(currentTile).floorChangeUP = True
                Case &H53
                  'mailbox
                Case &H54
                  'depot
                  DatTiles(currentTile).isDepot = True
                Case &H55
                  'trash
                Case &H56
                 'hole
                  DatTiles(currentTile).floorChangeDOWN = True
                  DatTiles(currentTile).requireShovel = True
                  DatTiles(currentTile).alwaysOnTop = True
                  DatTiles(currentTile).multitype = False
                Case &H57
                  'items with special description?
                Case &H58
                  'writtable
                  DatTiles(currentTile).RWInfo = 1 ' read only
                Case Else
                  ' should not happen
                  debugByte = optByte
                  Debug.Print "Tile loader found unexpected properties at " & GoodHex(optByte) & ": " & GoodHex(optbyte2)
                  #If TileDebug = 1 Then
                    tileLog = tileLog & " (unexpected subtype " & GoodHex(optByte) & " !!!!)"
                  #End If
                End Select 'optbyte2
                movePos pos, positionsRead
            Case Else
                ' Our parser needs an update
                #If TileDebug = 1 Then
                     tileLog = tileLog & vbCrLf & "PARSER ERROR - Unexpected tile: " & GoodHex(optByte) & vbCrLf
                     tileLog = tileLog & vbCrLf & "File continues:"
                     movePos pos, 100
                     LogOnFile "tibiadatdebug.txt", tileLog
                #End If
                LoadDatFileQ = -4 ' Unknown property
                Exit Function
            End Select
            
        Loop
        #If TileDebug = 1 Then
           tileLog = tileLog & vbCrLf
        #End If
    #End If
    
    ' some flags can be made by a combination of existing flags
    If DatTiles(currentTile).stackable = True Or DatTiles(currentTile).multitype = True Or _
      DatTiles(currentTile).fluidcontainer = True Then
      DatTiles(currentTile).haveExtraByte = True
    End If
    
    If DatTiles(currentTile).multiCharge = True Then
      DatTiles(currentTile).haveExtraByte = True
    End If

    If DatTiles(currentTile).alwaysOnTop = True Then
      DatTiles(currentTile).stackPriority = 3 ' high priority
    End If
    
    If DatTiles(currentTile).moreAlwaysOnTop = True Then
      DatTiles(currentTile).alwaysOnTop = True
      DatTiles(currentTile).stackPriority = 4 ' max priority
    End If
    
    ' add special cases of floor changers, for cavebot
    Select Case currentTile
    ' ramps that change floor when you step in
    Case tileID_rampToNorth, tileID_rampToSouth, tileID_rampToRightCycMountain, _
     tileID_rampToLeftCycMountain, tileID_rampToNorth, tileID_desertRamptoUp, _
     tileID_jungleStairsToNorth, tileID_jungleStairsToLeft
      DatTiles(currentTile).floorChangeUP = True
    Case tileID_grassCouldBeHole ' grass that will turn into a hole when you step in
      DatTiles(currentTile).floorChangeDOWN = True
    End Select
    
    '[CUSTOM FLAGS FOR BLACKDPROXY]
    'water, for smart autofisher
    If currentTile = tileID_waterWithFish Then
      DatTiles(currentTile).isWater = True
      DatTiles(currentTile).haveFish = True
    End If
    If currentTile = tileID_waterEmpty Then
      DatTiles(currentTile).isWater = True
    End If
    
    If currentTile = tileID_blockingBox Then
        DatTiles(currentTile).blocking = True
    End If
   
    


    If (currentTile >= tileID_waterWithFish) And (currentTile <= tileID_waterWithFishEnd) Then
      DatTiles(currentTile).isWater = True
      DatTiles(currentTile).haveFish = True
    End If
    If (currentTile >= tileID_waterEmpty) And (currentTile <= tileID_waterEmptyEnd) Then
      DatTiles(currentTile).isWater = True
    End If


    ' food, for autoeater
    If currentTile >= tileID_firstFoodTileID And currentTile <= tileID_lastFoodTileID Then
      DatTiles(currentTile).isFood = True
    End If
    If (currentTile >= tileID_firstMushroomTileID) And (currentTile <= tileID_lastMushroomTileID) Then
      DatTiles(currentTile).isFood = True
    End If
    
    Select Case currentTile ' special food
    Case &HA9, &H344, &H349, &H385, &HCB2, &H13E8, &H162E, &H1885, &H1886, &H18F8, &H18F9, &H18F9, &H18F9, &H1964, &H198D, &H198E, &H198F, &H1990, &H1991, &H19A9, &H19AE, &H1BF6, &H1BF7, &H1CCC, &H1CCD
      DatTiles(currentTile).isFood = True
    End Select
    
    If (currentTile >= 8010) And (currentTile <= 8020) Then ' special food
      DatTiles(currentTile).isFood = True
    End If
    
    
    ' fields, for a* smart path
    If currentTile >= tileID_firstFieldRangeStart And currentTile <= tileID_firstFieldRangeEnd Then
      DatTiles(currentTile).isField = True
    End If
    If (currentTile >= tileID_secondFieldRangeStart) And (currentTile <= tileID_secondFieldRangeEnd) Then
      DatTiles(currentTile).isField = True
    End If
    Select Case currentTile
    Case tileID_campFire1, tileID_campFire2
      DatTiles(currentTile).isField = True
    Case tileID_walkableFire1, tileID_walkableFire2, tileID_walkableFire3
      DatTiles(currentTile).isField = False 'dont consider fields that doesnt do any harm
    End Select
    If currentTile = tileID_woodenStairstoUp Then 'special stairs
      DatTiles(currentTile).floorChangeUP = True
    End If
    If currentTile = tileID_WallBugItem Then 'bug on walls, cant pick it!
      DatTiles(currentTile).pickupable = False
    End If
    
    ' Build some more combined properties:
    If lRare > &H1 Then
          DatTiles(currentTile).haveExtraByte2 = True ' UNKNOWN , TEST
    End If
    If DatTiles(currentTile).haveExtraByte = True Then 'BYTECOUNTdat5
      DatTiles(currentTile).totalExtraBytes = DatTiles(currentTile).totalExtraBytes + 1
    End If
    If DatTiles(currentTile).haveExtraByte2 = True Then
      DatTiles(currentTile).totalExtraBytes = DatTiles(currentTile).totalExtraBytes + 1
    End If
    
    #If TileDebug = 1 Then
         tileLog = tileLog & gDebug & vbCrLf
         If (Not (tileLog = "")) Then
           LogOnFile "tibiadatdebug.txt", tileLog
         End If
    #End If
    pos = properties_end + 1
  Loop
  AddBlackdProxyFlags ' Add our custom properties
  
  
  
'  CopyToDatTiles2
'  Dim resIgnore As Long
'  resIgnore = LoadDatFile11("c:\tibiaold\Tibia.dat")
'  CompareDatTiles



  LoadDatFileQ = 0
  Exit Function
  
badErr:
  DBGtileError = "Error number = " & CStr(Err.Number) & vbCrLf & "Error description = " & Err.Description & vbCrLf & "Path = " & tibiadathere
  LoadDatFileQ = -4 ' bad format or wrong version of given tibia.dat
End Function
'
'
'Public Sub CopyToDatTiles2()
'    Dim lastT As Long
'    Dim i As Integer
'
'    lastT = UBound(DatTiles)
'    ReDim DatTiles2(lastT)
'    For i = 0 To lastT
'        DatTiles2(i).alwaysOnTop = DatTiles(i).alwaysOnTop
'        DatTiles2(i).blocking = DatTiles(i).blocking
'        DatTiles2(i).blockingProjectile = DatTiles(i).blockingProjectile
'        DatTiles2(i).blockpickupable = DatTiles(i).blockpickupable
'        DatTiles2(i).canDecay = DatTiles(i).canDecay
'        DatTiles2(i).canWalkThrough = DatTiles(i).canWalkThrough
'        DatTiles2(i).floorChangeDOWN = DatTiles(i).floorChangeDOWN
'        DatTiles2(i).floorChangeUP = DatTiles(i).floorChangeUP
'        DatTiles2(i).fluidcontainer = DatTiles(i).fluidcontainer
'        DatTiles2(i).groundtile = DatTiles(i).groundtile
'        DatTiles2(i).haveExtraByte = DatTiles(i).haveExtraByte
'        DatTiles2(i).haveExtraByte2 = DatTiles(i).haveExtraByte2
'        DatTiles2(i).haveFish = DatTiles(i).haveFish
'        DatTiles2(i).haveName = DatTiles(i).haveName
'        DatTiles2(i).iscontainer = DatTiles(i).iscontainer
'        DatTiles2(i).isDepot = DatTiles(i).isDepot
'        DatTiles2(i).isDoor = DatTiles(i).isDoor
'        DatTiles2(i).isDoorWithLock = DatTiles(i).isDoorWithLock
'        DatTiles2(i).isField = DatTiles(i).isField
'        DatTiles2(i).isFood = DatTiles(i).isFood
'        DatTiles2(i).isWater = DatTiles(i).isWater
'        DatTiles2(i).itemName = DatTiles(i).itemName
'        DatTiles2(i).moreAlwaysOnTop = DatTiles(i).moreAlwaysOnTop
'        DatTiles2(i).multiCharge = DatTiles(i).multiCharge
'        DatTiles2(i).multitype = DatTiles(i).multitype
'        DatTiles2(i).noFloorChange = DatTiles(i).noFloorChange
'        DatTiles2(i).notMoveable = DatTiles(i).notMoveable
'        DatTiles2(i).pickupable = DatTiles(i).pickupable
'        DatTiles2(i).requireRightClick = DatTiles(i).requireRightClick
'        DatTiles2(i).requireRope = DatTiles(i).requireRope
'        DatTiles2(i).requireShovel = DatTiles(i).requireShovel
'        DatTiles2(i).RWInfo = DatTiles(i).RWInfo
'        DatTiles2(i).speed = DatTiles(i).speed
'        DatTiles2(i).stackable = DatTiles(i).stackable
'        DatTiles2(i).stackPriority = DatTiles(i).stackPriority
'        DatTiles2(i).totalExtraBytes = DatTiles(i).totalExtraBytes
'        DatTiles2(i).usable2 = DatTiles(i).usable2
'        DatTiles2(i).useable = DatTiles(i).useable
'    Next i
'End Sub
'
'Public Sub CompareDatTiles()
'    Dim lastT As Long
'    Dim i As Integer
'    Dim str As String
'    OverwriteOnFile "oldnewdiff.txt", "Differences found:" & vbCrLf
'    lastT = UBound(DatTiles)
'    For i = 0 To lastT
'        If Not (DatTiles(i).alwaysOnTop = DatTiles2(i).alwaysOnTop) Then
'            str = CStr(i) & ": alwaysOnTop"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).blocking = DatTiles2(i).blocking) Then
'            str = CStr(i) & ": blocking"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).blockingProjectile = DatTiles2(i).blockingProjectile) Then
'            str = CStr(i) & ": blockingProjectile"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).blockpickupable = DatTiles2(i).blockpickupable) Then
'            str = CStr(i) & ": blockpickupable"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).canDecay = DatTiles2(i).canDecay) Then
'            str = CStr(i) & ": canDecay"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).canWalkThrough = DatTiles2(i).canWalkThrough) Then
'            str = CStr(i) & ": canWalkThrough"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).floorChangeDOWN = DatTiles2(i).floorChangeDOWN) Then
'            str = CStr(i) & ": floorChangeDOWN"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).floorChangeUP = DatTiles2(i).floorChangeUP) Then
'            str = CStr(i) & ": floorChangeUP"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).fluidcontainer = DatTiles2(i).fluidcontainer) Then
'            str = CStr(i) & ": fluidcontainer"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).groundtile = DatTiles2(i).groundtile) Then
'            str = CStr(i) & ": groundtile"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).haveExtraByte = DatTiles2(i).haveExtraByte) Then
'            str = CStr(i) & ": haveExtraByte"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).haveExtraByte2 = DatTiles2(i).haveExtraByte2) Then
'            str = CStr(i) & ": haveExtraByte2"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).haveFish = DatTiles2(i).haveFish) Then
'            str = CStr(i) & ": haveFish"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).haveName = DatTiles2(i).haveName) Then
'            str = CStr(i) & ": haveName"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).iscontainer = DatTiles2(i).iscontainer) Then
'            str = CStr(i) & ": iscontainer"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).isDepot = DatTiles2(i).isDepot) Then
'            str = CStr(i) & ": isDepot"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).isDoor = DatTiles2(i).isDoor) Then
'            str = CStr(i) & ": isDoor"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).isDoorWithLock = DatTiles2(i).isDoorWithLock) Then
'            str = CStr(i) & ": isDoorWithLock"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).isField = DatTiles2(i).isField) Then
'            str = CStr(i) & ": isField"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).isFood = DatTiles2(i).isFood) Then
'            str = CStr(i) & ": isFood"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).isWater = DatTiles2(i).isWater) Then
'            str = CStr(i) & ": isWater"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).itemName = DatTiles2(i).itemName) Then
'            str = CStr(i) & ": itemName"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).moreAlwaysOnTop = DatTiles2(i).moreAlwaysOnTop) Then
'            str = CStr(i) & ": moreAlwaysOnTop"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).multiCharge = DatTiles2(i).multiCharge) Then
'            str = CStr(i) & ": multiCharge"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).multitype = DatTiles2(i).multitype) Then
'            str = CStr(i) & ": multitype"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).noFloorChange = DatTiles2(i).noFloorChange) Then
'            str = CStr(i) & ": noFloorChange"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).notMoveable = DatTiles2(i).notMoveable) Then
'            str = CStr(i) & ": notMoveable"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).pickupable = DatTiles2(i).pickupable) Then
'            str = CStr(i) & ": pickupable"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).requireRightClick = DatTiles2(i).requireRightClick) Then
'            str = CStr(i) & ": requireRightClick"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).requireRope = DatTiles2(i).requireRope) Then
'            str = CStr(i) & ": requireRope"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).requireShovel = DatTiles2(i).requireShovel) Then
'            str = CStr(i) & ": requireShovel"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).RWInfo = DatTiles2(i).RWInfo) Then
'            str = CStr(i) & ": RWInfo"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).speed = DatTiles2(i).speed) Then
'            str = CStr(i) & ": speed"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).stackable = DatTiles2(i).stackable) Then
'            str = CStr(i) & ": stackable"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).stackPriority = DatTiles2(i).stackPriority) Then
'            str = CStr(i) & ": stackPriority"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).totalExtraBytes = DatTiles2(i).totalExtraBytes) Then
'            str = CStr(i) & ": totalExtraBytes"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).usable2 = DatTiles2(i).usable2) Then
'            str = CStr(i) & ": usable2"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'        If Not (DatTiles(i).useable = DatTiles2(i).useable) Then
'            str = CStr(i) & ": useable"
'            LogOnFile "oldnewdiff.txt", str
'        End If
'    Next i
'End Sub

Public Function GetTileInfoString(b1 As Byte, b2 As Byte) As String
  ' 2 bytes indentify a tile in a packet
  Dim pos As Long
  Dim strRes As String
  #If FinalMode Then
  On Error GoTo exitf
  #End If
  strRes = ""
  pos = GetTheLong(b1, b2)
  If pos > 99 And pos <= highestDatTile Then
    strRes = ""
    If DatTiles(pos).alwaysOnTop = True Then
      strRes = " + alwaysOnTop=TRUE"
    End If
    If DatTiles(pos).moreAlwaysOnTop = True Then
      strRes = strRes & " + moreAlwaysOnTop=TRUE"
    End If
    If DatTiles(pos).usable2 = True Then
      strRes = strRes & " + usable2=TRUE"
    End If
    If DatTiles(pos).multiCharge = True Then
      strRes = strRes & " + multiCharge=TRUE"
    End If
    
    If DatTiles(pos).blocking = True Then
      strRes = strRes & " + blocking=TRUE"
    End If
    If DatTiles(pos).blockingProjectile = True Then
      strRes = strRes & " + blockingProyectile=TRUE"
    End If
    If DatTiles(pos).blockpickupable = False Then
      strRes = strRes & " + blockpickupable=FALSE"
    End If
    If DatTiles(pos).canDecay = False Then
      strRes = strRes & " + canDecay=FALSE"
    End If
    If DatTiles(pos).canWalkThrough = True Then
      strRes = strRes & " + canWalkThrough=TRUE"
    End If
    If DatTiles(pos).fluidcontainer = True Then
      strRes = strRes & " + fluidcontainer=TRUE"
    End If
    If DatTiles(pos).groundtile = True Then
      strRes = strRes & " + groundtile=TRUE"
    End If
    If DatTiles(pos).iscontainer = True Then
      strRes = strRes & " + iscontainer=TRUE"
    End If
    If DatTiles(pos).isDoor = True Then
      strRes = strRes & " + isDoor=TRUE"
    End If
    If DatTiles(pos).multitype = True Then
      strRes = strRes & " + multitype=TRUE"
    End If
    If DatTiles(pos).noFloorChange = True Then
      strRes = strRes & " + noFloorChange=TRUE"
    End If
    If DatTiles(pos).notMoveable = True Then
      strRes = strRes & " + notMoveable=TRUE"
    End If
    If DatTiles(pos).pickupable = True Then
      strRes = strRes & " + pickupable=TRUE"
    End If
    If DatTiles(pos).RWInfo <> 0 Then
      strRes = strRes & " + RWInfo=" & DatTiles(pos).RWInfo
    End If
    If DatTiles(pos).speed <> 0 Then
      strRes = strRes & " + speed=" & DatTiles(pos).speed
    End If
    If DatTiles(pos).stackable = True Then
      strRes = strRes & " + stackeable=TRUE"
    End If
    If DatTiles(pos).useable = True Then
      strRes = strRes & " + useable=TRUE"
    End If
    If DatTiles(pos).haveExtraByte = True Then
      strRes = strRes & " + haveExtraByte=TRUE"
    End If
    If DatTiles(pos).haveExtraByte2 = True Then
      strRes = strRes & " + haveExtraByte2=TRUE"
    End If
    If DatTiles(pos).isWater = True Then
      strRes = strRes & " + isWater=TRUE"
    End If
    If DatTiles(pos).floorChangeUP = True Then
      strRes = strRes & " + floorChangeUP=TRUE"
    End If
    If DatTiles(pos).floorChangeDOWN = True Then
      strRes = strRes & " + floorChangeDOWN=TRUE"
    End If
    If DatTiles(pos).requireRightClick = True Then
      strRes = strRes & " + requireRightClick=TRUE"
    End If
    If DatTiles(pos).requireRope = True Then
      strRes = strRes & " + requireRope=TRUE"
    End If
    If DatTiles(pos).requireShovel = True Then
      strRes = strRes & " + requireShovel=TRUE"
    End If
    
    If DatTiles(pos).isFood = True Then
      strRes = strRes & " + isFood=TRUE"
    End If
    If DatTiles(pos).isField = True Then
      strRes = strRes & " + isField=TRUE"
    End If
    
    If DatTiles(pos).isDepot = True Then
      strRes = strRes & " + isDepot=TRUE"
    End If
    
    If DatTiles(pos).haveName = True Then
      strRes = strRes & " + itemName=" & DatTiles(pos).itemName
    End If
    
    If strRes = "" Then
      strRes = "This tile doesn't have any special tag!"
    End If
  ElseIf pos = 99 Then
    strRes = "about player-monster-npc"
  Else
    strRes = "This is not a valid tile ID!"
  End If
exitf:
  GetTileInfoString = strRes
End Function


Public Sub GetInfoOfNewDatFile(ByVal strFilePath As String, ByRef par_version As Long, ByRef par_file As String)
  On Error GoTo goterr
  par_version = 0
  par_file = ""
  If (Right$(strFilePath, 9) = "tibia.dat") Then
      par_version = 0
      par_file = "tibia.dat"
      Exit Sub
  End If
  Dim fso As Scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  Dim Filename As String
  Dim parts() As String
  Dim appearancesFound As String
  Dim mustContinueLoop As Boolean
  mustContinueLoop = True
  appearancesFound = False
  Set fso = New Scripting.FileSystemObject
    Filename = strFilePath & "catalog-content.json"
    If fso.FileExists(Filename) = True Then
      fn = FreeFile
      Open Filename For Input As #fn
      While ((Not EOF(fn)) And (mustContinueLoop = True))
        Line Input #fn, strLine
        strLine = Trim$(strLine)
        If strLine <> "" Then
            strLine = Replace(strLine, """", "")
            strLine = Replace(strLine, ",", "")
            parts = Split(strLine, ":")
            If (UBound(parts) > 0) Then
                If (parts(0) = "type") Then
                    If (parts(1) = "appearances") Then
                        appearancesFound = True
                    End If
                End If
                If (parts(0) = "file") Then
                    If (appearancesFound) Then
                        par_file = parts(1)
                    End If
                End If
                If (parts(0) = "version") Then
                    If (appearancesFound) Then
                        par_version = CLng(parts(1))
                        mustContinueLoop = False
                    End If
                End If
            End If
        End If
        If (Left$(strLine, 1) = "}") Then
            appearancesFound = False
        End If
      Wend
      Close #fn
    End If
    Exit Sub
goterr:
    dateErrDescription = "Error " & Err.Number & " at GetDATEOfFile. Here:" & vbCrLf & _
     strLine & vbCrLf & vbCrLf & "Error description: " & Err.Description
    par_version = -1
    par_file = ""
End Sub


Public Function UnifiedLoadDatFile(ByVal strPath As String) As Long
  Dim res As Long
  If TibiaVersionLong <= 740 Then
    firstValidOutfit = 2
    lastValidOutfit = 142
    'res = LoadDatFile740(strPath)
    res = LoadDatFile2(strPath)
  ElseIf TibiaVersionLong <= 750 Then
    firstValidOutfit = 2
    lastValidOutfit = 142
    res = LoadDatFile(strPath)
  ElseIf TibiaVersionLong < 773 Then
    firstValidOutfit = 2
    lastValidOutfit = 142
    res = LoadDatFile2(strPath)
  ElseIf TibiaVersionLong < 860 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile3(strPath)
  ElseIf TibiaVersionLong < 872 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile4(strPath)
  ElseIf TibiaVersionLong < 940 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile5(strPath)
  ElseIf TibiaVersionLong < 960 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile6(strPath)
  ElseIf TibiaVersionLong < 1000 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile7(strPath)
  ElseIf TibiaVersionLong < 1021 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile8(strPath)
  ElseIf TibiaVersionLong < 1050 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile9(strPath)
  ElseIf TibiaVersionLong < 1058 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile10(strPath)
  ElseIf TibiaVersionLong < 1100 Then
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFile11(strPath)
  Else
    firstValidOutfit = 2
    lastValidOutfit = 160
    res = LoadDatFileQ(strPath)
  End If
  UnifiedLoadDatFile = res
End Function


