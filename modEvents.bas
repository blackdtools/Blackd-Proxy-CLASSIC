Attribute VB_Name = "modEvents"
#Const FinalMode = 1
Option Explicit

Public Const MAXSCHEDULED = 20
Public Const DELAYBETWEENAUTOMSG_ms = 2000
Public Type TypeEvent
  id As Integer
  flags As String
  trigger As String
  action As String
End Type

Public Type TypeCondEvent
  thing1 As String
  operator As String
  thing2 As String
  delay As String
  lock As String
  keep As String
  action As String
  nextunlock As Long
End Type

Public Type TypeEventList
  Number As Long
  ev() As TypeEvent ' 1 To MAXEVENTS
End Type

Public Type TypeCondEventList
  Number As Long
  ev() As TypeCondEvent ' 1 To MAXCONDS
End Type

Public Type TypeScheduledAction
  pending As Boolean
  clientID As Integer
  action As String
  tickc As Long
End Type
Public scheduledActions(1 To MAXSCHEDULED) As TypeScheduledAction

Public eventsIDselected As Long
Public condEventsIDselected As Long
Public CustomEvents() As TypeEventList
Public CustomCondEvents() As TypeCondEventList
Public var_lastsender() As String
Public var_lastmsg() As String
Public nextAllowedmsg() As Long
Public TimerConditionTick As Long
Public TimerConditionTick2 As Long
Public MAXEVENTS As Long
Public MAXCONDS As Long

Public Sub ResetEventList(idConnection)
  CustomEvents(idConnection).Number = 0
End Sub

Public Sub ResetCondEventList(idConnection)
  CustomCondEvents(idConnection).Number = 0
End Sub

Public Sub TelephoneCall(strNumber As String)
  If strNumber = "" Then
    ShellExecute 0, "open", "callto://" & frmEvents.txtTelephoneNumber.Text, "", "", vbNormalFocus
  Else
    ShellExecute 0, "open", "callto://" & strNumber, "", "", vbNormalFocus
  End If
End Sub

Public Sub ChangePlayTheDangerSound(newValue As Boolean)
  Dim aRes As Long
  If ((newValue = True) And (PlayTheDangerSound = False)) Then
    If ((frmRunemaker.chkOnDangerSS.Value = 1) And (frmRunemaker.timerSS.enabled = False)) Then
        frmRunemaker.timerSS.enabled = True
    End If
    If frmEvents.chkTelephoneAlarm.Value = 1 Then
      TelephoneCall ""
    End If
  End If
  PlayTheDangerSound = newValue
End Sub

Public Sub AddSchedule(idConnection As Integer, action As String, tickc As Long)
  Dim i As Integer
  For i = 1 To MAXSCHEDULED
    If scheduledActions(i).pending = False Then
      scheduledActions(i).action = action
      scheduledActions(i).clientID = idConnection
      scheduledActions(i).tickc = tickc
      scheduledActions(i).pending = True
      Exit For
    End If
  Next i
End Sub
Public Sub ProcessEventMsg(idConnection As Integer, thetype As Byte)
  Dim aRes As Long
  Dim nEvents As Long
  Dim evType As Integer
  Dim partL As String
  Dim partR As String
  Dim intRes As Integer
  Dim executeThis As String
  Dim i As Integer
  Dim lotype As Long
  Dim mustdelay As Long
  Dim strTmp As String
  #If FinalMode Then
  On Error GoTo errIgnore
  #End If
  'aRes = SendLogSystemMessageToClient(idconnection, _
   var_lastsender(idconnection) & " sent you a message type " & GoodHex(thetype) & _
   ": " & var_lastmsg(idconnection))
 ' DoEvents
  If TibiaVersionLong >= 820 Then
    If thetype = 1 Then
      thetype = 1
    ElseIf thetype = 0 Then
      thetype = 0
    Else
      thetype = thetype - 2
    End If
  End If
  If thetype > 19 Then
    Exit Sub
  End If
  nEvents = CustomEvents(idConnection).Number
  partL = LCase(var_lastmsg(idConnection))
  lotype = CLng(thetype) + 1
  For i = 1 To nEvents
    evType = CustomEvents(idConnection).ev(i).id
    If evType = 0 Then 'contains ...
      If Mid(CustomEvents(idConnection).ev(i).flags, lotype, 1) = 1 Then
        If ((CheatsPaused(idConnection) = False) Or (Mid$(CustomEvents(idConnection).ev(i).flags, 19, 1) = "1")) Then
          partR = LCase(parseVars(idConnection, CustomEvents(idConnection).ev(i).trigger))
          If InStr(partL, partR) > 0 Then 'triggered
            executeThis = parseVars(idConnection, CustomEvents(idConnection).ev(i).action)
            strTmp = Right$(CustomEvents(idConnection).ev(i).flags, Len(CustomEvents(idConnection).ev(i).flags) - 20)
            mustdelay = CLng(strTmp)
            If mustdelay = 0 Then
              intRes = ExecuteInTibia(executeThis, idConnection, False)
            Else
              mustdelay = mustdelay + GetTickCount()
              AddSchedule idConnection, executeThis, mustdelay
            End If
          End If
        End If
      End If
    ElseIf evType = 1 Then 'exact
      If Mid(CustomEvents(idConnection).ev(i).flags, lotype, 1) = 1 Then
        If ((CheatsPaused(idConnection) = False) Or (Mid$(CustomEvents(idConnection).ev(i).flags, 19, 1) = "1")) Then
          partR = LCase(parseVars(idConnection, CustomEvents(idConnection).ev(i).trigger))
          If partL = partR Then 'triggered
            executeThis = parseVars(idConnection, CustomEvents(idConnection).ev(i).action)
            strTmp = Right$(CustomEvents(idConnection).ev(i).flags, Len(CustomEvents(idConnection).ev(i).flags) - 20)
            mustdelay = CLng(strTmp)
            If mustdelay = 0 Then
              intRes = ExecuteInTibia(executeThis, idConnection, False)
            Else
              mustdelay = mustdelay + GetTickCount()
              AddSchedule idConnection, executeThis, mustdelay
            End If
          End If
        End If
      End If
    End If

  Next i
  Exit Sub
errIgnore:
  frmEvents.Caption = "Error " & CStr(Err.Number) & " at ProcessEventMsg : " & Err.Description
End Sub
