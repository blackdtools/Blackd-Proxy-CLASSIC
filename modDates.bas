Attribute VB_Name = "modDates"
#Const FinalMode = 1
Option Explicit
Public Type TypeTrial
  mode As Integer
  bDays As Long
End Type
Public Function ReadTrial(ByRef backup As String) As TypeTrial
  Dim res As TypeTrial
  Dim str As String
  Dim impByte As String
  Dim impNum As Integer
  Dim bLen As Integer
  Dim bDays As Long
  Dim strRan As String
  Dim mode As Integer
  On Error GoTo gotErr
  res.bDays = -1
  res.mode = 2
  str = backup
  impByte = Left(str, 1)
  impNum = CInt(impByte)
  str = Right(str, Len(str) - 11 - impNum)
  bLen = CInt(Left(str, 1))
  str = Right(str, Len(str) - 1)
  bDays = CLng(Left(str, bLen))
  str = Right(str, Len(str) - bLen)
  impByte = Left(str, 1)
  impNum = CInt(impByte)
  str = Right(str, Len(str) - 1)
  strRan = Left(str, impNum)
  str = Right(str, Len(str) - impNum)
  impNum = CInt(strRan) + 100
  If Len(str) <> (impNum + 2) Then
    GoTo gotErr
  End If
  str = Right(str, Len(str) - impNum)
  str = Left(str, 1)
  mode = CInt(str)
  If (mode > 2) Then
    ' full version
    mode = 3
  End If
  'valid?
  If (mode <> 3) And ((bDays < 300) Or (bDays > 590)) Then 'max trial can be 1 Aug 2006
    GoTo gotErr
  End If
  'sucesfull end
  res.bDays = bDays
  res.mode = mode
  ReadTrial = res
  Exit Function
gotErr:
  res.bDays = -1
  res.mode = 2
  ReadTrial = res
End Function

Public Function ReadTrialFromFile() As TypeTrial
  Dim resT As TypeTrial
  Dim fn As Integer
  Dim i As Long
  Dim str As String
  On Error GoTo gotErr
  ' load memory adresses for login IPs
  fn = FreeFile
  Open App.Path & "\code.txt" For Input As #fn
    Line Input #fn, str
  Close #fn
  resT = ReadTrial(str)
  If resT.mode = 3 Then
    GoTo gotErr
  Else
    ReadTrialFromFile = resT
  End If
  Exit Function
gotErr:
  resT.bDays = -1
  resT.mode = 2
  ReadTrialFromFile = resT
End Function

Public Function CompareTibiaDate(tibiaDate As String) As Boolean
  'true : inside trial period
  'false : trial period expired
  #If FinalMode Then
    On Error GoTo justend
  #End If
  Dim t As String
  Dim tday As String
  Dim tmonth As String
  Dim tyear As String
  Dim cday As Integer
  Dim cmonth As Integer
  Dim cyear As Integer
  Dim blackdDays As Long
  t = tibiaDate
  tday = Left(t, 2)
  tmonth = Mid(t, 5, 3)
  tyear = Right(t, 4)
  cday = CInt(tday)
  Select Case tmonth
  Case "Jan"
    cmonth = 0
  Case "Feb"
    cmonth = 1
  Case "Mar"
    cmonth = 2
  Case "Apr"
    cmonth = 3
  Case "May"
    cmonth = 4
  Case "Jun"
    cmonth = 5
  Case "Jul"
    cmonth = 6
  Case "Aug"
    cmonth = 7
  Case "Sep"
    cmonth = 8
  Case "Oct"
    cmonth = 9
  Case "Nov"
    cmonth = 10
  Case "Dec"
    cmonth = 11
  End Select
  cyear = CInt(tyear) - 2005
  ' how many "blackd" days since 1 Jan 2005?
  blackdDays = (CLng(cyear) * 372) + (CLng(cmonth) * 31) + CLng(cday)
  ' they should be less than the trial limit
  If (blackdDays <= TrialLimit_Day) And (TrialLimit_Day <> -1) Then
    CompareTibiaDate = True
  Else
    CompareTibiaDate = False
  End If
  Exit Function
justend:
  End
End Function

Public Sub UpdateCompDate()
  Dim newDateTime As String
  Dim myCompDay As String
  Dim myCompMonth As String
  Dim myCompYear As String
  Dim cday As Integer
  Dim cmonth As Integer
  Dim cyear As Integer
  Dim blackdDays As Long
  #If FinalMode Then
    On Error GoTo justEndU
  #End If
  newDateTime = Format(Date, "dd/mm/yyyy")
  myCompDay = Left(newDateTime, 2)
  myCompMonth = Mid(newDateTime, 4, 2)
  myCompYear = Right(newDateTime, 4)
  cday = CInt(myCompDay)
  cmonth = CInt(myCompMonth) - 1
  cyear = CInt(myCompYear) - 2005
  blackdDays = (CLng(cyear) * 372) + (CLng(cmonth) * 31) + CLng(cday)
  If (blackdDays > TrialLimit_Day) Or (TrialLimit_Day = -1) Then
    'trial check failed
    GotTrialLock = True
    lastLockReason = "Failed Trial Check #3"
  End If
  Exit Sub
justEndU:
  End
End Sub

