Attribute VB_Name = "modAstar"
#Const FinalMode = 1
Option Explicit
' move orders
Public Const MoveRight = &H1
Public Const MoveNorthRight = &H2
Public Const MoveNorth = &H3
Public Const MoveNorthLeft = &H4
Public Const MoveLeft = &H5
Public Const MoveSouthLeft = &H6
Public Const MoveSouth = &H7
Public Const MoveSouthRight = &H8
Public Const MoveStartPoint = &H0

Public Const StrMoveRight = "1"
Public Const StrMoveNorthRight = "2"
Public Const StrMoveNorth = "3"
Public Const StrMoveNorthLeft = "4"
Public Const StrMoveLeft = "5"
Public Const StrMoveSouthLeft = "6"
Public Const StrMoveSouth = "7"
Public Const StrMoveSouthRight = "8"
Public Const StrMoveStartPoint = "0"

Public Const ExtraCostDiagonal = 5
Public Const CostWalkable = 1
Public Const CostNearHandicap = 2
Public Const CostHandicap = 100
Public Const CostBlock = 10000

Const MAX_ITERATIONS = 2000

Public ProcesingBigMapPath As Boolean

Public RequiredMoveBuffer() As String
Public ReadyBuffer() As Boolean

Public Type TypePathResult
  Id As Double
  tileID As Long
  melee As Boolean
  hmm As Boolean
  requireShovel As Boolean
  requireRightClick As Boolean
  requireRope As Boolean
  X As Long
  y As Long
End Type

Public Type TypeAstarMatrix
  cost(-9 To 10, -7 To 8) As Long
  ' no walkable > 9999
  ' fire-poison-energy fields : 100
  ' walkable : 1
  ' (diagonal : x3)
End Type

Public Type TypeListItem
  X As Long
  y As Long
  currF As Long ' current F value (F=G+H)
  currG As Long ' current G value
  movingFrom As Byte ' move order from current parent to reach this point
End Type
Public Function ManhattanDistance(x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Long
  ManhattanDistance = (Abs((x1 - x2)) + Abs((y1 - y2)))
End Function
Public Function PositionInAstarList(aList() As TypeListItem, X As Long, y As Long) As Long
Dim res As Long
Dim i As Long
res = 0
For i = 1 To UBound(aList)
  If (aList(i).X = X) And (aList(i).y = y) Then
    res = i
    Exit For
  End If
Next i
PositionInAstarList = res
End Function
Public Sub ExtendOpenList(X As Long, y As Long, currG As Long, ByRef openList() As TypeListItem, ByRef closedList() As TypeListItem, ByRef m As TypeAstarMatrix, goalX As Long, goalY As Long)
  Dim newX As Long
  Dim newY As Long
  Dim i As Long
  Dim f As Long
  Dim g As Long
  Dim h As Long
  If (X >= 10) Or (X <= -9) Or (y <= -7) Or (y >= 8) Then
    Exit Sub
  End If
  newX = X - 1
  newY = y - 1
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY) + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveNorthLeft
  End If
  newX = X
  newY = y - 1
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY)
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveNorth
  End If
  newX = X + 1
  newY = y - 1
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY) + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveNorthRight
  End If
  newX = X + 1
  newY = y
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY)
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveRight
  End If
  newX = X + 1
  newY = y + 1
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY) + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveSouthRight
  End If
  newX = X
  newY = y + 1
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY)
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveSouth
  End If
  newX = X - 1
  newY = y + 1
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY) + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveSouthLeft
  End If
  newX = X - 1
  newY = y
  If (m.cost(newX, newY) < CostBlock) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + m.cost(newX, newY)
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      AddAstarItem openList, newX, newY, f, g, MoveLeft
  End If
End Sub



Public Sub AddAstarItem(aList() As TypeListItem, X As Long, y As Long, f As Long, g As Long, mfrom As Byte)
  Dim nb As Long
  Dim found As Boolean
  Dim shouldChange As Boolean
  Dim shouldChangeI As Long
  Dim tmp As Long
  Dim i As Long
  found = False

  tmp = PositionInAstarList(aList, X, y)
  If tmp = 0 Then
    found = False
  Else
    found = True
    If aList(tmp).currG > g Then
      shouldChange = True
      shouldChangeI = tmp
    End If
  End If
  If shouldChange = True Then
    ' just update
    aList(shouldChangeI).currF = f
    aList(shouldChangeI).currG = g
    aList(shouldChangeI).movingFrom = mfrom
    aList(shouldChangeI).X = X
    aList(shouldChangeI).y = y
  End If
  If found = False Then
    ' resize list
    nb = UBound(aList) + 1
    ReDim Preserve aList(nb)
    ' write item in list
    aList(nb).currF = f
    aList(nb).currG = g
    aList(nb).movingFrom = mfrom
    aList(nb).X = X
    aList(nb).y = y
    ' update number of items on the list
    aList(0).X = nb
  End If
End Sub
Public Sub JustAddAstarItem(aList() As TypeListItem, X As Long, y As Long, f As Long, g As Long, mfrom As Byte)
  Dim nb As Long
  Dim found As Boolean
  Dim shouldChange As Boolean
  Dim shouldChangeI As Long
  Dim tmp As Long
  Dim i As Long
  found = False
  tmp = PositionInAstarList(aList, X, y)
  If tmp = 0 Then
    found = False
  Else
    found = True

  End If
  If found = False Then
    ' resize list
    nb = UBound(aList) + 1
    ReDim Preserve aList(nb)
    ' write item in list
    aList(nb).currF = f
    aList(nb).currG = g
    aList(nb).movingFrom = mfrom
    aList(nb).X = X
    aList(nb).y = y
    ' update number of items on the list
    aList(0).X = nb
  End If
End Sub





Public Sub DeleteAstarItem(aList() As TypeListItem, X As Long, y As Long)
  Dim found As Boolean
  Dim foundpos As Long
  Dim i As Long
  Dim nb As Long
  found = False
  nb = UBound(aList)
  For i = 1 To nb
    If aList(i).X = X And aList(i).y = y Then
      'found in list
      found = True
      foundpos = i
      Exit For
    End If
  Next i
  If found = True Then
    ' resize list
    nb = nb - 1
    ' move all elements 1 position back
    For i = foundpos To nb
      aList(i).currF = aList(i + 1).currF
      aList(i).currG = aList(i + 1).currG
      aList(i).movingFrom = aList(i + 1).movingFrom
      aList(i).X = aList(i + 1).X
      aList(i).y = aList(i + 1).y
    Next i
    ' shrink list (last item is lost)
    ReDim Preserve aList(nb)
    ' update number of items on the list
    aList(0).X = nb
  End If
End Sub
Public Sub InitAstarList(ByRef aList() As TypeListItem)
  ReDim aList(0)
  aList(0).X = 0
End Sub
Public Function Astar(x1 As Long, y1 As Long, goalX As Long, goalY As Long, ByRef m As TypeAstarMatrix) As String
  Dim openList() As TypeListItem
  Dim closedList() As TypeListItem
  Dim res As String
  Dim i As Long
  Dim currX As Long
  Dim currY As Long
  Dim found As Boolean
  Dim bestF As Long
  Dim bestI As Long
  Dim tmpF As Long
  Dim tmpX As Long
  Dim tmpY As Long
  Dim tmpG As Long
  Dim tmpFrom As Byte
  Dim exitf As Boolean
  Dim foundstart As Boolean
  Dim elem As Long
  Dim fromB As Byte
  Dim Px As Long
  Dim Py As Long
  Dim gFound As Long
  Dim res2 As String
  Dim tmpStr As String
  res = ""
  ' step 0. trivial check. start point = goal point?
  If x1 = goalX And y1 = goalY Then
    Astar = ""
    Exit Function ' -> trivial solution, end
  End If
  ' step 1. init lists - in the field .x of item 0 we will store the number of items
  InitAstarList openList
  InitAstarList closedList
  ' step 2. we add first item to openList
  tmpF = ManhattanDistance(x1, y1, goalX, goalY)
  JustAddAstarItem openList, x1, y1, tmpF, 0, MoveStartPoint
  ' step 3. main loop
  exitf = False
  Do
  bestF = CostBlock
  bestI = 0
  For i = 1 To UBound(openList)
  tmpF = openList(i).currF
    If tmpF < bestF Then
      bestI = i
      bestF = tmpF
    End If
  Next i
  If bestI = 0 Then ' if no items left (best=none) then we can't continue
    res = "X" '  -> no path found , end
    exitf = True
  Else
    tmpX = openList(bestI).X
    tmpY = openList(bestI).y
    tmpF = openList(bestI).currF
    tmpG = openList(bestI).currG
    tmpFrom = openList(bestI).movingFrom
    ' extend the node in openList with min F, add it to closedList and delete it from openList
    JustAddAstarItem closedList, tmpX, tmpY, tmpF, tmpG, tmpFrom
    ExtendOpenList tmpX, tmpY, tmpG, openList, closedList, m, goalX, goalY
    DeleteAstarItem openList, tmpX, tmpY
  End If
  gFound = PositionInAstarList(closedList, goalX, goalY) ' step 4. loop to 3 until goal found in closed list
  If gFound > 0 Then
    ' goal found
    ' -> now build shortest way from goal and return it
    res = ""
    foundstart = False
    Px = goalX
    Py = goalY
    Do
      elem = PositionInAstarList(closedList, Px, Py)
      fromB = closedList(elem).movingFrom
      tmpStr = Hex(fromB)
      Select Case fromB
      Case MoveNorth
        res = tmpStr & res
        Py = Py + 1
      Case MoveRight
        res = tmpStr & res
        Px = Px - 1
      Case MoveSouth
        res = tmpStr & res
        Py = Py - 1
      Case MoveLeft
        res = tmpStr & res
        Px = Px + 1
      Case MoveNorthRight
        res = tmpStr & res
        Px = Px - 1
        Py = Py + 1
      Case MoveSouthRight
        res = tmpStr & res
        Px = Px - 1
        Py = Py - 1
      Case MoveSouthLeft
        res = tmpStr & res
        Px = Px + 1
        Py = Py - 1
      Case MoveNorthLeft
        res = tmpStr & res
        Px = Px + 1
        Py = Py + 1
      Case MoveStartPoint
        foundstart = True
      End Select
    Loop Until foundstart = True
    exitf = True
  End If
  Loop Until exitf = True
  If Len(res) >= maxStepsPerMovement Then
    res = Left$(res, maxStepsPerMovement)
  End If
  Astar = res
End Function

'AstarGiveBestGoal will return x and y of best goal to a long distance point
Public Function AstarGiveBestGoal(x1 As Long, y1 As Long, goalX As Long, goalY As Long, ByRef m As TypeAstarMatrix) As TypePathResult
  Dim openList() As TypeListItem
  Dim closedList() As TypeListItem
  Dim res As TypePathResult
  Dim i As Long
  Dim currX As Long
  Dim currY As Long
  Dim found As Boolean
  Dim bestF As Long
  Dim bestI As Long
  Dim tmpF As Long
  Dim tmpX As Long
  Dim tmpY As Long
  Dim tmpG As Long
  Dim tmpFrom As Byte
  Dim exitf As Boolean
  Dim foundstart As Boolean
  Dim elem As Long
  Dim fromB As Byte
  Dim Px As Long
  Dim Py As Long
  Dim gFound As Long
  Dim res2 As String
  Dim tmpStr As String
  Dim bestH As Long
  Dim tmpH As Long
  res.X = 0
  res.y = 0
  ' step 0. trivial check. start point = goal point?
  If x1 = goalX And y1 = goalY Then
    AstarGiveBestGoal = res
    Exit Function ' -> trivial solution, end
  End If
  ' step 1. init lists - in the field .x of item 0 we will store the number of items
  InitAstarList openList
  InitAstarList closedList
  ' step 2. we add first item to openList
  tmpF = ManhattanDistance(x1, y1, goalX, goalY)
  JustAddAstarItem openList, x1, y1, tmpF, 0, MoveStartPoint
  ' step 3. main loop
  exitf = False
  Do
  bestF = CostBlock
  bestI = 0
  For i = 1 To UBound(openList)
  tmpF = openList(i).currF
    If tmpF < bestF Then
      bestI = i
      bestF = tmpF
    End If
  Next i
  If bestI = 0 Then ' if no items left return pathable goal with lowest H
    bestH = 1000000
    tmpH = 0
    bestI = 0
    res.X = 0
    res.y = 0
    For i = 1 To UBound(closedList)
    tmpH = ManhattanDistance(closedList(i).X, closedList(i).y, goalX, goalY)
    If tmpH < bestH Then
       bestH = tmpH
       bestI = i
    End If
    Next i
    If bestI <> 0 Then
      res.X = closedList(bestI).X
      res.y = closedList(bestI).y
    End If
    exitf = True
    AstarGiveBestGoal = res

    Exit Function
  Else
    tmpX = openList(bestI).X
    tmpY = openList(bestI).y
    tmpF = openList(bestI).currF
    tmpG = openList(bestI).currG
    tmpFrom = openList(bestI).movingFrom
    ' extend the node in openList with min F, add it to closedList and delete it from openList
    JustAddAstarItem closedList, tmpX, tmpY, tmpF, tmpG, tmpFrom
    ExtendOpenList tmpX, tmpY, tmpG, openList, closedList, m, goalX, goalY
    DeleteAstarItem openList, tmpX, tmpY
  End If
  gFound = PositionInAstarList(closedList, goalX, goalY) ' step 4. loop to 3 until goal found in closed list
  If gFound > 0 Then
    ' goal found!
    res.X = goalX
    res.y = goalY
    exitf = True
  End If
  Loop Until exitf = True

  AstarGiveBestGoal = res
End Function


Public Sub AstarBig(idConnection As Integer, x1 As Long, y1 As Long, goalX As Long, goalY As Long, z As Long, debMode As Boolean)
  Dim openList() As TypeListItem
  Dim closedList() As TypeListItem
  Dim res As String
  Dim i As Long
  Dim currX As Long
  Dim currY As Long
  Dim found As Boolean
  Dim bestF As Long
  Dim bestI As Long
  Dim tmpF As Long
  Dim tmpX As Long
  Dim tmpY As Long
  Dim tmpG As Long
  Dim tmpFrom As Byte
  Dim exitf As Boolean
  Dim foundstart As Boolean
  Dim elem As Long
  Dim fromB As Byte
  Dim Px As Long
  Dim Py As Long
  Dim gFound As Long
  Dim res2 As String
  Dim tmpStr As String
  Dim iter As Long
  Dim firstChecks As Boolean
  ProcesingBigMapPath = True
  'ReadyBuffer(idconnection) = False
  iter = 0
  res = ""
  ' step 0. trivial check. start point = goal point?
  If x1 = goalX And y1 = goalY Then
    ProcesingBigMapPath = False
    RequiredMoveBuffer(idConnection) = ""
    ReadyBuffer(idConnection) = True
    Exit Sub ' -> trivial solution, end
  End If
  GetBigMapSquareB firstChecks, goalX, goalY, z
  If firstChecks = False Then
    ' maybe it was a stair?
    GetBigMapSquareB firstChecks, goalX + 1, goalY, z
    If firstChecks = True Then
      goalX = goalX + 1
      GoTo continueIt
    End If
    GetBigMapSquareB firstChecks, goalX - 1, goalY, z
    If firstChecks = True Then
      goalX = goalX - 1
      GoTo continueIt
    End If
    GetBigMapSquareB firstChecks, goalX, goalY + 1, z
    If firstChecks = True Then
      goalY = goalY + 1
      GoTo continueIt
    End If
    GetBigMapSquareB firstChecks, goalX, goalY - 1, z
    If firstChecks = True Then
      goalY = goalY - 1
      GoTo continueIt
    End If
    ' it was blocked area, so it is impossible
    ProcesingBigMapPath = False
    RequiredMoveBuffer(idConnection) = "X"
    ReadyBuffer(idConnection) = True
    Exit Sub ' -> trivial no solution, end
  End If
continueIt:
  ' step 1. init lists - in the field .x of item 0 we will store the number of items
  InitAstarList openList
  InitAstarList closedList
  ' step 2. we add first item to openList
  tmpF = ManhattanDistance(x1, y1, goalX, goalY)
  JustAddAstarItem openList, x1, y1, tmpF, 0, MoveStartPoint
  ' step 3. main loop
  exitf = False
  Do
  iter = iter + 1
  'DoEvents
  bestF = CostBlock
  bestI = 0
  For i = 1 To UBound(openList)
  tmpF = openList(i).currF
    If tmpF < bestF Then
      bestI = i
      bestF = tmpF
    End If
  Next i
  If bestI = 0 Then ' if no items left (best=none) then we can't continue
      ProcesingBigMapPath = False
      RequiredMoveBuffer(idConnection) = "X"
      ReadyBuffer(idConnection) = True
      Exit Sub
  Else
    tmpX = openList(bestI).X
    tmpY = openList(bestI).y
    tmpF = openList(bestI).currF
    tmpG = openList(bestI).currG
    tmpFrom = openList(bestI).movingFrom
    ' extend the node in openList with min F, add it to closedList and delete it from openList
    JustAddAstarItem closedList, tmpX, tmpY, tmpF, tmpG, tmpFrom
    If iter > MAX_ITERATIONS Then
      ProcesingBigMapPath = False
      RequiredMoveBuffer(idConnection) = "X"
      ReadyBuffer(idConnection) = True
      Exit Sub
    End If
    ExtendOpenListBig tmpX, tmpY, tmpG, openList, closedList, goalX, goalY, z, debMode
    DeleteAstarItem openList, tmpX, tmpY
  End If
  gFound = PositionInAstarList(closedList, goalX, goalY) ' step 4. loop to 3 until goal found in closed list
  If gFound > 0 Then
    ' goal found
    ' -> now build shortest way from goal and return it
    res = ""
    foundstart = False
    Px = goalX
    Py = goalY
    Do
      elem = PositionInAstarList(closedList, Px, Py)
      fromB = closedList(elem).movingFrom
      tmpStr = Hex(fromB)
      Select Case fromB
      Case MoveNorth
        res = tmpStr & res
        Py = Py + 1
      Case MoveRight
        res = tmpStr & res
        Px = Px - 1
      Case MoveSouth
        res = tmpStr & res
        Py = Py - 1
      Case MoveLeft
        res = tmpStr & res
        Px = Px + 1
      Case MoveNorthRight
        res = tmpStr & res
        Px = Px - 1
        Py = Py + 1
      Case MoveSouthRight
        res = tmpStr & res
        Px = Px - 1
        Py = Py - 1
      Case MoveSouthLeft
        res = tmpStr & res
        Px = Px + 1
        Py = Py - 1
      Case MoveNorthLeft
        res = tmpStr & res
        Px = Px + 1
        Py = Py + 1
      Case MoveStartPoint
        foundstart = True
      End Select
    Loop Until foundstart = True
    exitf = True
  End If
  Loop Until exitf = True
  ProcesingBigMapPath = False
  If Len(res) >= maxStepsPerMovement Then
    res = Left$(res, maxStepsPerMovement)
  End If
  RequiredMoveBuffer(idConnection) = res
  ReadyBuffer(idConnection) = True
End Sub

Public Sub OptimizeBuffer(idConnection As Integer)
  Dim wBuffer As String
  Dim lBuffer As Long
  Dim currposX As Long
  Dim currposY As Long
  Dim firstO As String
  Dim j As Long
  Dim i As Long
  Dim orders As String
  Dim myMap As TypeAstarMatrix
  wBuffer = RequiredMoveBuffer(idConnection)
  If (wBuffer = "") Or (wBuffer = "X") Then
    RequiredMoveBuffer(idConnection) = ""
    Exit Sub
  End If
  lBuffer = Len(wBuffer)
  currposX = 0
  currposY = 0
  For i = 1 To lBuffer
    firstO = Left(wBuffer, 1)
    Select Case firstO
            Case StrMoveRight
              currposX = currposX + 1
            Case StrMoveNorthRight
              currposX = currposX + 1
              currposY = currposY - 1
            Case StrMoveNorth
              currposY = currposY - 1
            Case StrMoveNorthLeft
              currposX = currposX - 1
              currposY = currposY - 1
            Case StrMoveLeft
              currposX = currposX - 1
            Case StrMoveSouthLeft
              currposX = currposX - 1
              currposY = currposY + 1
            Case StrMoveSouth
              currposY = currposY + 1
            Case StrMoveSouthRight
              currposX = currposX + 1
              currposY = currposY + 1
    End Select
    If (Abs(currposX) > 7) Or (Abs(currposY) > 5) Then
      Exit For
    End If
    j = Len(wBuffer)
    If j > 1 Then
      wBuffer = Right(wBuffer, j - 1)
    End If
  Next i
   ReadTrueMap idConnection, myMap
  'force goal to be walkable
  myMap.cost(currposX, currposY) = CostWalkable
  orders = Astar(0, 0, currposX, currposY, myMap)
  If (orders <> "") And (orders <> "X") Then
    RequiredMoveBuffer(idConnection) = orders & wBuffer
  Else
    RequiredMoveBuffer(idConnection) = ""
  End If
End Sub


Public Sub ExtendOpenListBig(X As Long, y As Long, currG As Long, ByRef openList() As TypeListItem, _
 ByRef closedList() As TypeListItem, goalX As Long, goalY As Long, z As Long, debMode As Boolean)
  Dim newX As Long
  Dim newY As Long
  Dim i As Long
  Dim f As Long
  Dim g As Long
  Dim h As Long
  Dim res As Boolean
  newX = X - 1
  newY = y - 1
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1 + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveNorthLeft
  End If
  newX = X
  newY = y - 1
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveNorth
  End If
  newX = X + 1
  newY = y - 1
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1 + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveNorthRight
  End If
  newX = X + 1
  newY = y
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveRight
  End If
  newX = X + 1
  newY = y + 1
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1 + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveSouthRight
  End If
  newX = X
  newY = y + 1
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveSouth
  End If
  newX = X - 1
  newY = y + 1
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1 + ExtraCostDiagonal
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveSouthLeft
  End If
  newX = X - 1
  newY = y
  GetBigMapSquareB res, X, y, z
  If (res = True) And (PositionInAstarList(closedList, newX, newY) = 0) Then
      'add to open list
      g = currG + 1
      h = ManhattanDistance(newX, newY, goalX, goalY)
      f = g + h
      If debMode Then
        DrawXYZPixel newX, newY, z, vbYellow
      End If
      AddAstarItem openList, newX, newY, f, g, MoveLeft
  End If
End Sub



