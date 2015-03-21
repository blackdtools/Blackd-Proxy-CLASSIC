Attribute VB_Name = "modReadInmediately"
#Const FinalMode = 0
Option Explicit

Public Const cteOfficial As String = " - OFFICIAL TIBIA"
Public Const ctePreview As String = " - PREVIEW"

Public arrayConfigPaths() As String
Public arrayConfigVersions() As String
Public arrayConfigVersionsLong() As String

Public loadDebugStart As String
Public Sub ReadIniVeryFirst()
  ' This function will read some important vars first of all
  Dim i As Integer
  Dim strInfo As String
  Dim lonInfo As Long
  Dim here As String
  Dim tmpStr As String
  Dim res As Long
  Dim p1 As String
  Dim p2 As String
  Dim idLoginSP As Long
  Dim tmpNumber As Long
  Dim tmpVersion As String
  Dim debugPoint As Long
  Dim userHere As String
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  res = -1
  debugPoint = 1
  userHere = App.path
  debugPoint = 2
  If Right$(userHere, 1) = "\" Then
    userHere = userHere & "settings.ini"
  Else
    userHere = userHere & "\settings.ini"
  End If
  debugPoint = 3
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "configPath", "", strInfo, Len(strInfo), myMainConfigINIPath(), True)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    configPath = Trim$(strInfo)
  Else
    configPath = ""
  End If
  
  strInfo = String$(2500, 0)
  i = getBlackdINI("Proxy", "addConfigPaths", "", strInfo, Len(strInfo), myMainConfigINIPath(), True)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    addConfigPaths = Trim$(strInfo)
  Else
    addConfigPaths = ""
  End If
  
  strInfo = String$(2500, 0)
  i = getBlackdINI("Proxy", "addConfigVersions", "", strInfo, Len(strInfo), myMainConfigINIPath(), True)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    addConfigVersions = Trim$(strInfo)
  Else
    addConfigVersions = ""
  End If
  
  strInfo = String$(2500, 0)
  i = getBlackdINI("Proxy", "addConfigVersionsLongs", "", strInfo, Len(strInfo), myMainConfigINIPath(), True)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    addConfigVersionsLongs = Trim$(strInfo)
  Else
    addConfigVersionsLongs = ""
  End If
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "highestTibiaVersionLong", "", strInfo, Len(strInfo), myMainConfigINIPath(), True)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    lonInfo = CLng(Trim$(strInfo))
    highestTibiaVersionLong = lonInfo
  Else
    highestTibiaVersionLong = 0
  End If
  

  
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "TibiaVersionDefaultString", "", strInfo, Len(strInfo), myMainConfigINIPath(), True)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    TibiaVersionDefaultString = Trim$(strInfo)
  Else
    TibiaVersionDefaultString = ""
  End If
  
  strInfo = String$(250, 0)
  i = getBlackdINI("Proxy", "TibiaVersionForceString", "", strInfo, Len(strInfo), myMainConfigINIPath(), True)
  If i > 0 Then
    strInfo = Left(strInfo, i)
    TibiaVersionForceString = Trim$(strInfo)
  Else
    TibiaVersionForceString = ""
  End If
  
  ' FIX for people with computers that somehow don't read root config.ini correctly...
  If Len(addConfigPaths) < Len(FIX_addConfigPaths) Then
    addConfigPaths = FIX_addConfigPaths
    addConfigVersions = FIX_addConfigVersions
    addConfigVersionsLongs = FIX_addConfigVersionsLongs
    highestTibiaVersionLong = CLng(FIX_highestTibiaVersionLong)
    TibiaVersionDefaultString = FIX_TibiaVersionDefaultString
    TibiaVersionForceString = FIX_TibiaVersionForceString
  End If
  
  arrayConfigPaths = Split(addConfigPaths, ",")
  arrayConfigVersions = Split(addConfigVersions, ",")
  arrayConfigVersionsLong = Split(addConfigVersionsLongs, ",")
 
  Exit Sub
goterr:
   Debug.Print ("Error ReadIniVeryFirst " & Err.Description)
End Sub

Public Sub SafeAddNewVersions(ByRef cmbVersion As ComboBox)
  
  On Error GoTo goterr

 
    Dim strItem As String
    Dim addit As String
    Dim addit2 As String
    Dim lasti As Long
    Dim i As Long
    loadDebugStart = ""
    lasti = UBound(arrayConfigVersions)
    If lasti > -1 Then
        For i = lasti To 0 Step -1
          strItem = Trim$("" & arrayConfigVersions(i))
          addit = "Tibia " & strItem
          cmbVersion.AddItem (addit)
          If UBound(arrayConfigVersionsLong) >= i Then
            If arrayConfigVersionsLong(i) = highestTibiaVersionLong Then
                addit2 = addit & cteOfficial
                cmbVersion.AddItem (addit2)
            End If
          Else
            loadDebugStart = loadDebugStart & vbCrLf & "ERROR: ConfigVersionsLong and ConfigVersions have different number of elements!" & vbCrLf & "addConfigVersionsLongs=" & addConfigVersionsLongs & vbCrLf & _
            "addConfigVersions=" & addConfigVersions
            Exit Sub
          End If
        Next i
    End If
Exit Sub
goterr:
    loadDebugStart = loadDebugStart & vbCrLf & "Error at SafeAddNewVersions (code " & CStr(Err.Number) & " ) Description= " & Err.Description

End Sub

Public Function getConfigPathOf(ByVal caseSel As String) As String
    Dim i As Long
    Dim strTest As String
    Dim strItem As String
    Dim addit As String
    Dim addit2 As String
    Dim res As String
    Dim lasti As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
 
 
    res = "config" & highestTibiaVersionLong
    OVERWRITE_OT_MODE = True
    
    lasti = UBound(arrayConfigVersions)
    If lasti > -1 Then
        For i = lasti To 0 Step -1
          strItem = "" & arrayConfigVersions(i)
          addit = "Tibia " & strItem
          If caseSel = addit Then
                res = arrayConfigPaths(i)
                OVERWRITE_OT_MODE = True
                getConfigPathOf = res
                Exit Function
          End If
          If arrayConfigVersionsLong(i) = highestTibiaVersionLong Then
              addit2 = addit & cteOfficial
              If caseSel = addit2 Then
                res = arrayConfigPaths(i)
                OVERWRITE_OT_MODE = False
                getConfigPathOf = res
                Exit Function
              End If
          End If
        Next i
    End If
    getConfigPathOf = res
    Exit Function
goterr:
    OVERWRITE_OT_MODE = False
    getConfigPathOf = configPath = "config" & highestTibiaVersionLong
End Function

   Public Sub SetFutureTibiaVersion(ByVal parConfigPath As String)
   
     #If FinalMode Then
  On Error GoTo goterr
  #End If
     Dim i As Long
    Dim strTest As String
        Dim lasti As Long
        Dim strItem As String
    TibiaVersion = TibiaVersionDefaultString
    TibiaVersionLong = highestTibiaVersionLong
    
    
    lasti = UBound(arrayConfigPaths)
    If lasti > -1 Then
        For i = lasti To 0 Step -1
          strItem = "" & arrayConfigPaths(i)
          If arrayConfigPaths(i) = parConfigPath Then
            TibiaVersion = arrayConfigVersions(i)
            TibiaVersionLong = CLng(arrayConfigVersionsLong(i))
          End If
        Next i
    End If
    
    Exit Sub
goterr:
    TibiaVersion = TibiaVersionDefaultString
    TibiaVersionLong = highestTibiaVersionLong
   End Sub

