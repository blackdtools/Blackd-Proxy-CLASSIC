Attribute VB_Name = "modTrainer"
#Const FinalMode = 0
Option Explicit
Public trainerIDselected As Long
Public Type TypeRefillSlot
  itemID_b1 As Byte
  itemID_b2 As Byte
  xvalue As Long
  cheked As Long
End Type
Public Type TypeTrainerOptions
  enabled As Long
  spearID_b1 As Byte
  spearID_b2 As Byte
  spearDest As Integer
  maxitems As Long
  PlayerSlots(1 To EQUIPMENT_SLOTS) As TypeRefillSlot
  misc_stoplowhp As Long
  stoplowhpHP As Long
  misc_dance_14min As Long
  misc_avoidID As Long
  idToAvoid As Double
  AllowedSides(0 To 8) As Boolean
End Type
Public Type TypePairOfBytes
  b1 As Byte
  b2 As Byte
End Type
Public TrainerOptions() As TypeTrainerOptions
Public subdebug651 As Long

Public Function safeConvertStringToPairOfBytes(str As String) As TypePairOfBytes
  Dim res As TypePairOfBytes
  Dim b11 As Byte
  Dim b12 As Byte
  Dim b21 As Byte
  Dim b22 As Byte
  Dim tileID As Long
  On Error GoTo safeEnd
  res.b1 = &H0
  res.b2 = &H0
  If Len(str) = 5 Then
    b11 = FromHexToDec(Mid$(str, 1, 1))
    b12 = FromHexToDec(Mid$(str, 2, 1))
    b21 = FromHexToDec(Mid$(str, 4, 1))
    b22 = FromHexToDec(Mid$(str, 5, 1))
    If ((b11 < 16) And (b12 < 16) And (b21 < 16) And (b22 < 16)) Then
      If ((tileID >= 0) And (tileID <= highestDatTile)) Then
      res.b1 = (b11 * 16) + b12
      res.b2 = (b21 * 16) + b22
      End If
    End If
  End If
  safeConvertStringToPairOfBytes = res
  Exit Function
safeEnd:
  res.b1 = &H0
  res.b2 = &H0
  safeConvertStringToPairOfBytes = res
End Function

Public Function safeConvertStringToLong(str As String) As Long
  Dim res As Long
  On Error GoTo safeEnd
  If IsNumeric(str) = True Then
    res = CLng(str)
  Else
    res = 0
  End If
  safeConvertStringToLong = res
  Exit Function
safeEnd:
  safeConvertStringToLong = 0
End Function

Public Function safeConvertStringToDouble(str As String) As Double
  Dim res As Double
  On Error GoTo safeEnd
  If IsNumeric(str) = True Then
    res = CDbl(str)
  Else
    res = 0
  End If
  safeConvertStringToDouble = res
  Exit Function
safeEnd:
  safeConvertStringToDouble = 0
End Function

Public Sub ResetInternalTrainerValues(i As Integer)
    Dim j As Integer
    subdebug651 = 0
    TrainerOptions(i).maxitems = 4
    subdebug651 = 1
    TrainerOptions(i).spearID_b1 = &HCD
    subdebug651 = 2
    TrainerOptions(i).spearID_b2 = &HC
    subdebug651 = 3
    TrainerOptions(i).spearDest = 0
    subdebug651 = 4
    For j = 1 To EQUIPMENT_SLOTS
      TrainerOptions(i).PlayerSlots(j).itemID_b1 = &HCD
      subdebug651 = 5
      TrainerOptions(i).PlayerSlots(j).itemID_b2 = &HC
      subdebug651 = 6
      TrainerOptions(i).PlayerSlots(j).xvalue = 1
      subdebug651 = 7
      TrainerOptions(i).PlayerSlots(j).cheked = 0
      subdebug651 = 8
    Next j
    TrainerOptions(i).misc_stoplowhp = 0
    subdebug651 = 9
    TrainerOptions(i).stoplowhpHP = 50
    subdebug651 = 10
    TrainerOptions(i).misc_dance_14min = 0
    subdebug651 = 11
    TrainerOptions(i).misc_avoidID = 0
    subdebug651 = 12
    TrainerOptions(i).idToAvoid = 0
    subdebug651 = 13
    For j = 0 To 8
      TrainerOptions(i).AllowedSides(j) = False
      subdebug651 = 14
    Next j
    subdebug651 = 15
End Sub
