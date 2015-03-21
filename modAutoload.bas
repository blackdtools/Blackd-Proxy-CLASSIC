Attribute VB_Name = "modAutoload"
#Const FinalMode = 0
Option Explicit
Private Const CteAutoloadSubfolder As String = "autoload"
Public SettingsOfChar As scripting.Dictionary  ' A dictionary Char Name (string) -> Settings (string)
Private AutoloadUsable As Boolean
Private AutoloadPath As String

Public Aux_LastLoadedCond() As TypeCondEvent

Public Function BooleanToUnifiedString(blnValue As Boolean) As String
    If blnValue = True Then
        BooleanToUnifiedString = "1"
    Else
        BooleanToUnifiedString = "0"
    End If
End Function

Public Function UnifiedStringToBoolean(strValue As String) As Boolean
    If strValue = "1" Then
        UnifiedStringToBoolean = True
    Else
        UnifiedStringToBoolean = False
    End If
End Function

Private Sub LoadThisCharSetting(idConnection As Integer, strVar As String, strValue As String)
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    Dim i As Long
    Dim blnTemp As Boolean
    Dim aRes As Long
    Dim tmpStr As String
    Dim tempID As Long
    Dim subValue1 As String
    Dim subValue2 As String
    Dim pieces() As String
    
    'Debug.Print "Loaded:" & strVar & "=" & strValue & "<<<"
    Select Case strVar
    Case "BEGIN_CavebotScript"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmCavebot.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmCavebot.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            cavebotIDselected = frmCavebot.cmbCharacter.ListIndex
            cavebotScript(cavebotIDselected).RemoveAll
            cavebotLenght(cavebotIDselected) = 0
            frmCavebot.UpdateValues
        End If
    Case "ADD_CavebotLine"
        AddIDLine cavebotIDselected, cavebotLenght(cavebotIDselected), strValue
        cavebotLenght(cavebotIDselected) = cavebotLenght(cavebotIDselected) + 1
    Case "END_CavebotScript"
        frmCavebot.UpdateValues
    Case "LastCavebotFile"
        frmCavebot.txtFile.Text = strValue
    Case "CavebotEnabled"
        If strValue = "1" Then
          tmpStr = "exiva openbp"
          tempID = GetTickCount() + 1000
          AddSchedule idConnection, tmpStr, tempID
          frmCavebot.TurnCavebotState idConnection, True
        Else
            frmCavebot.TurnCavebotState idConnection, False
        End If
    Case "BEGIN_Runemaker"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmRunemaker.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmRunemaker.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            runemakerIDselected = frmRunemaker.cmbCharacter.ListIndex
            frmRunemaker.UpdateValues
        End If
    Case "Runemaker_autoEat"
        RuneMakerOptions(runemakerIDselected).autoEat = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoLogoutAnyFloor"
        RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoLogoutCurrentFloor"
        RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoLogoutOutOfRunes"
        RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes = UnifiedStringToBoolean(strValue)
    Case "Runemaker_autoWaste"
        RuneMakerOptions(runemakerIDselected).autoWaste = UnifiedStringToBoolean(strValue)
    Case "Runemaker_firstActionMana"
        RuneMakerOptions(runemakerIDselected).firstActionMana = CLng(strValue)
    Case "Runemaker_firstActionText"
        RuneMakerOptions(runemakerIDselected).firstActionText = strValue
    Case "Runemaker_LowMana"
        RuneMakerOptions(runemakerIDselected).LowMana = CLng(strValue)
    Case "Runemaker_ManaFluid"
        RuneMakerOptions(runemakerIDselected).ManaFluid = UnifiedStringToBoolean(strValue)
        If (RuneMakerOptions(runemakerIDselected).ManaFluid = False) Then
            RemoveSpamOrder CInt(runemakerIDselected), 4 'remove auto mana
        End If
    Case "Runemaker_msgSound"
        RuneMakerOptions(runemakerIDselected).msgSound = UnifiedStringToBoolean(strValue)
    Case "Runemaker_msgSound2"
        RuneMakerOptions(runemakerIDselected).msgSound2 = UnifiedStringToBoolean(strValue)
    Case "Runemaker_secondActionMana"
        RuneMakerOptions(runemakerIDselected).secondActionMana = CLng(strValue)
    Case "Runemaker_secondActionSoulpoints"
        RuneMakerOptions(runemakerIDselected).secondActionSoulpoints = CLng(strValue)
    Case "Runemaker_secondActionText"
        RuneMakerOptions(runemakerIDselected).secondActionText = strValue
    Case "Runemaker_activated"
        RuneMakerOptions(runemakerIDselected).activated = UnifiedStringToBoolean(strValue)
    Case "END_Runemaker"
        frmRunemaker.UpdateValues
        
        
    Case "BEGIN_CustomCondEvents"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmCondEvents.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmCondEvents.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            'frmCondEvents.UpdateValues
            condEventsIDselected = frmCondEvents.cmbCharacter.ListIndex
            frmCondEvents.DeleteAllCondEvents CLng(idConnection)
            frmCondEvents.UpdateValues
        End If
    Case "CustomCondEvents_thing1"
        Aux_LastLoadedCond(idConnection).thing1 = strValue
    Case "CustomCondEvents_operator"
        Aux_LastLoadedCond(idConnection).operator = strValue
    Case "CustomCondEvents_thing2"
        Aux_LastLoadedCond(idConnection).thing2 = strValue
    Case "CustomCondEvents_delay"
        Aux_LastLoadedCond(idConnection).delay = strValue
    Case "CustomCondEvents_lock"
        Aux_LastLoadedCond(idConnection).lock = strValue
    Case "CustomCondEvents_keep"
        Aux_LastLoadedCond(idConnection).keep = strValue
    Case "CustomCondEvents_action"
        Aux_LastLoadedCond(idConnection).action = strValue
    Case "CustomCondEvents_ADD"
        aRes = frmCondEvents.AddCondEvent(idConnection, _
         Aux_LastLoadedCond(idConnection).thing1, _
         Aux_LastLoadedCond(idConnection).operator, _
         Aux_LastLoadedCond(idConnection).thing2, _
         Aux_LastLoadedCond(idConnection).delay, _
         Aux_LastLoadedCond(idConnection).lock, _
         Aux_LastLoadedCond(idConnection).keep, _
         Aux_LastLoadedCond(idConnection).action)
    Case "END_CustomCondEvents"
         frmCondEvents.UpdateValues
    Case "BEGIN_Trainer"
        blnTemp = False
        For i = 1 To MAXCLIENTS
            If LCase(frmTrainer.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
                frmTrainer.cmbCharacter.ListIndex = i
                blnTemp = True
            End If
        Next i
        If blnTemp = True Then
            'frmCondEvents.UpdateValues
            trainerIDselected = frmTrainer.cmbCharacter.ListIndex
            frmTrainer.UpdateValues
        End If
    Case "Trainer_AllowedSides"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).AllowedSides(CLng(subValue1)) = UnifiedStringToBoolean(subValue2)
    Case "Trainer_idToAvoid"
        TrainerOptions(idConnection).idToAvoid = CLng(strValue)
    Case "Trainer_maxitems"
        TrainerOptions(idConnection).maxitems = CLng(strValue)
    Case "Trainer_misc_avoidID"
        TrainerOptions(idConnection).misc_avoidID = CLng(strValue)
    Case "Trainer_misc_stoplowhp"
        TrainerOptions(idConnection).misc_stoplowhp = CLng(strValue)
    Case "Trainer_spearDest"
        TrainerOptions(idConnection).spearDest = CLng(strValue)
    Case "Trainer_spearID_b1"
        TrainerOptions(idConnection).spearID_b1 = CByte("&H" & strValue)
    Case "Trainer_spearID_b2"
        TrainerOptions(idConnection).spearID_b2 = CByte("&H" & strValue)
    Case "Trainer_stoplowhpHP"
        TrainerOptions(idConnection).stoplowhpHP = CLng(strValue)
    Case "Trainer_PlayerSlots_cheked"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).cheked = CLng(subValue2)
    Case "Trainer_PlayerSlots_itemID_b1"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).itemID_b1 = CByte("&H" & subValue2)
     Case "Trainer_PlayerSlots_itemID_b2"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).itemID_b2 = CByte("&H" & subValue2)
     Case "Trainer_PlayerSlots_xvalue"
        pieces = Split(strValue, ",")
        subValue1 = pieces(0)
        If UBound(pieces) > 0 Then
          subValue2 = pieces(1)
        Else
          subValue2 = ""
        End If
        TrainerOptions(idConnection).PlayerSlots(CLng(subValue1)).xvalue = CLng(subValue2)
    Case "Trainer_enabled"
        TrainerOptions(idConnection).enabled = CLng(strValue)
    Case "END_Trainer"
      trainerIDselected = idConnection
      frmTrainer.UpdateValues
    End Select
    Exit Sub
goterr:
    Exit Sub
End Sub

Public Function OverwriteOnPathFileSimple(pathfile As String, strtext As String) As Long
  Dim fn As Integer
  Dim writeThis As String
  Dim a As Long
  On Error GoTo ignoreit
  a = 0
  fn = FreeFile
  writeThis = strtext
  Open pathfile For Output As #fn
    Print #fn, writeThis
  Close #fn
  OverwriteOnPathFileSimple = 0
  Exit Function
ignoreit:
  OverwriteOnPathFileSimple = CLng(Err.Number)
End Function


Public Function LoadCharSettings(idConnection As Integer, Optional charName As String = "") As String
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    Dim loadCharName As String
    Dim strSettings As String
    Dim pieces() As String
    Dim strLine As String
    Dim ai As Long
    Dim strVarName As String
    Dim strVarValue As String
    Dim posSpliter As Long
    Dim blnTemp As Boolean
    If AutoloadUsable = False Then
        LoadCharSettings = "Autoload is not usable in this environment"
        Exit Function
    End If
    If GameConnected(idConnection) = False Then
        LoadCharSettings = "Character is not connected"
        Exit Function
    End If
    If charName = "" Then
        loadCharName = CharacterName(idConnection)
    Else
        loadCharName = charName
    End If
    strSettings = GetSettingsOfChar(loadCharName)
    If strSettings = "" Then
        LoadCharSettings = "System could not find saved settings found for character " & loadCharName
        Exit Function
    End If
    pieces = Split(strSettings, vbCrLf)
    For ai = 0 To UBound(pieces)
      strLine = Trim$(pieces(ai))
      If strLine <> "" Then
       posSpliter = InStr(1, strLine, "=", vbTextCompare)
       If (posSpliter > 0) Then
        strVarName = Left$(strLine, posSpliter - 1)
        strVarValue = Right$(strLine, Len(strLine) - posSpliter)
        LoadThisCharSetting idConnection, strVarName, strVarValue
       End If
      End If
    Next ai
    LoadCharSettings = ""
    Exit Function
goterr:
    LoadCharSettings = "Unexpected error #" & CStr(Err.Number) & " at LoadCharSettings: " & Err.Description
End Function

Public Sub SaveCharSettings(idConnection As Integer)
    Dim aRes As Long
    Dim charName As String
    Dim myPath As String
    Dim strSettings As String
    Dim tmpRes As Long
    Dim blnTemp As Long
    Dim i As Long
    Dim j As Long
    #If FinalMode Then
    On Error GoTo goterr
    #End If
    If GameConnected(idConnection) = True Then
        charName = CharacterName(idConnection)
    End If
    If AutoloadUsable = False Then
        aRes = GiveGMmessage(idConnection, "Unable to load or save settings in your system (Because Folder/hard disk/security restrictions) " & CStr(Err.Number), "BlackdProxy")
        DoEvents
        Exit Sub
    End If
    myPath = App.path
    If (Right$(myPath, 1) <> "\") And (Right$(myPath, 1) <> "/") Then
      myPath = myPath & "\" & CteAutoloadSubfolder & "\" & CharacterName(idConnection) & ".txt"
    Else
      myPath = myPath & CteAutoloadSubfolder & "\" & CharacterName(idConnection) & ".txt"
    End If
    strSettings = ""
    ' save cavebot
    blnTemp = False
    For i = 1 To MAXCLIENTS
        If LCase(frmCavebot.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
            frmCavebot.cmbCharacter.ListIndex = i
            blnTemp = True
        End If
    Next i
    If blnTemp = True Then
        cavebotIDselected = frmCavebot.cmbCharacter.ListIndex
        frmCavebot.UpdateValues
        If cavebotLenght(cavebotIDselected) > 0 Then
            strSettings = "BEGIN_CavebotScript=1" & vbCrLf
            For j = 0 To cavebotLenght(cavebotIDselected) - 1
                strSettings = strSettings & "ADD_CavebotLine=" & _
                GetStringFromIDLine(idConnection, j) & vbCrLf
            Next j
            strSettings = strSettings & "END_CavebotScript=1" & vbCrLf
            strSettings = strSettings & "LastCavebotFile=" & frmCavebot.txtFile.Text & vbCrLf
        End If
    End If
    If frmCavebot.chkEnabled.Value = 1 Then
        strSettings = strSettings & "CavebotEnabled=1" & vbCrLf
    Else
        strSettings = strSettings & "CavebotEnabled=0" & vbCrLf
    End If
    
    ' save runemaker
    blnTemp = False
    For i = 1 To MAXCLIENTS
        If LCase(frmRunemaker.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
            frmRunemaker.cmbCharacter.ListIndex = i
            blnTemp = True
        End If
    Next i
    If blnTemp = True Then
        runemakerIDselected = frmRunemaker.cmbCharacter.ListIndex
        strSettings = strSettings & "BEGIN_Runemaker=1" & vbCrLf
        strSettings = strSettings & "Runemaker_autoEat=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).autoEat) & vbCrLf
        strSettings = strSettings & "Runemaker_autoLogoutAnyFloor=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).autoLogoutAnyFloor) & vbCrLf
        strSettings = strSettings & "Runemaker_autoLogoutCurrentFloor=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).autoLogoutCurrentFloor) & vbCrLf
        strSettings = strSettings & "Runemaker_autoLogoutOutOfRunes=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).autoLogoutOutOfRunes) & vbCrLf
        strSettings = strSettings & "Runemaker_autoWaste=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).autoWaste) & vbCrLf
        strSettings = strSettings & "Runemaker_firstActionMana=" & CStr(RuneMakerOptions(runemakerIDselected).firstActionMana) & vbCrLf
        strSettings = strSettings & "Runemaker_firstActionText=" & RuneMakerOptions(runemakerIDselected).firstActionText & vbCrLf
        strSettings = strSettings & "Runemaker_LowMana=" & CStr(RuneMakerOptions(runemakerIDselected).LowMana) & vbCrLf
        strSettings = strSettings & "Runemaker_ManaFluid=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).ManaFluid) & vbCrLf
        strSettings = strSettings & "Runemaker_msgSound=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).msgSound) & vbCrLf
        strSettings = strSettings & "Runemaker_msgSound2=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).msgSound2) & vbCrLf
        strSettings = strSettings & "Runemaker_secondActionMana=" & CStr(RuneMakerOptions(runemakerIDselected).secondActionMana) & vbCrLf
        strSettings = strSettings & "Runemaker_secondActionSoulpoints=" & CStr(RuneMakerOptions(runemakerIDselected).secondActionSoulpoints) & vbCrLf
        strSettings = strSettings & "Runemaker_secondActionText=" & RuneMakerOptions(runemakerIDselected).secondActionText & vbCrLf
        strSettings = strSettings & "Runemaker_activated=" & BooleanToUnifiedString(RuneMakerOptions(runemakerIDselected).activated) & vbCrLf
        strSettings = strSettings & "END_Runemaker=1" & vbCrLf
    End If
    
    'conds
    blnTemp = False
    For i = 1 To MAXCLIENTS
        If LCase(frmCondEvents.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
            frmCondEvents.cmbCharacter.ListIndex = i
            blnTemp = True
        End If
    Next i
    If blnTemp = True Then
        condEventsIDselected = frmCondEvents.cmbCharacter.ListIndex
        frmCondEvents.UpdateValues
        strSettings = strSettings & "BEGIN_CustomCondEvents=1" & vbCrLf
        For i = 1 To CustomCondEvents(condEventsIDselected).Number
            strSettings = strSettings & "CustomCondEvents_thing1=" & CustomCondEvents(condEventsIDselected).ev(i).thing1 & vbCrLf
            strSettings = strSettings & "CustomCondEvents_operator=" & CustomCondEvents(condEventsIDselected).ev(i).operator & vbCrLf
            strSettings = strSettings & "CustomCondEvents_thing2=" & CustomCondEvents(condEventsIDselected).ev(i).thing2 & vbCrLf
            strSettings = strSettings & "CustomCondEvents_delay=" & CustomCondEvents(condEventsIDselected).ev(i).delay & vbCrLf
            strSettings = strSettings & "CustomCondEvents_lock=" & CustomCondEvents(condEventsIDselected).ev(i).lock & vbCrLf
            strSettings = strSettings & "CustomCondEvents_keep=" & CustomCondEvents(condEventsIDselected).ev(i).keep & vbCrLf
            strSettings = strSettings & "CustomCondEvents_action=" & CustomCondEvents(condEventsIDselected).ev(i).action & vbCrLf
            strSettings = strSettings & "CustomCondEvents_ADD=1" & vbCrLf
        Next i
        strSettings = strSettings & "END_CustomCondEvents=1" & vbCrLf
    End If
    
    
    'trainer
    blnTemp = False
    For i = 1 To MAXCLIENTS
        If LCase(frmTrainer.cmbCharacter.List(i)) = LCase(CharacterName(idConnection)) Then
            frmTrainer.cmbCharacter.ListIndex = i
            blnTemp = True
        End If
    Next i
    If blnTemp = True Then
        trainerIDselected = frmTrainer.cmbCharacter.ListIndex
        frmTrainer.UpdateValues
        strSettings = strSettings & "BEGIN_Trainer=1" & vbCrLf
       
        For i = 0 To 8
           strSettings = strSettings & "Trainer_AllowedSides=" & CStr(i) & "," & BooleanToUnifiedString(TrainerOptions(trainerIDselected).AllowedSides(i)) & vbCrLf
        Next i
        strSettings = strSettings & "Trainer_idToAvoid=" & TrainerOptions(trainerIDselected).idToAvoid & vbCrLf
        strSettings = strSettings & "Trainer_maxitems=" & TrainerOptions(trainerIDselected).maxitems & vbCrLf
        strSettings = strSettings & "Trainer_misc_avoidID=" & TrainerOptions(trainerIDselected).misc_avoidID & vbCrLf
        strSettings = strSettings & "Trainer_misc_dance_14min=" & TrainerOptions(trainerIDselected).misc_dance_14min & vbCrLf
        strSettings = strSettings & "Trainer_misc_stoplowhp=" & TrainerOptions(trainerIDselected).misc_stoplowhp & vbCrLf
        strSettings = strSettings & "Trainer_spearDest=" & TrainerOptions(trainerIDselected).spearDest & vbCrLf
        strSettings = strSettings & "Trainer_spearID_b1=" & GoodHex(TrainerOptions(trainerIDselected).spearID_b1) & vbCrLf
        strSettings = strSettings & "Trainer_spearID_b2=" & GoodHex(TrainerOptions(trainerIDselected).spearID_b2) & vbCrLf
        strSettings = strSettings & "Trainer_stoplowhpHP=" & TrainerOptions(trainerIDselected).stoplowhpHP & vbCrLf
        
        For i = 1 To EQUIPMENT_SLOTS
          strSettings = strSettings & "Trainer_PlayerSlots_cheked=" & CStr(i) & "," & TrainerOptions(trainerIDselected).PlayerSlots(i).cheked & vbCrLf
          strSettings = strSettings & "Trainer_PlayerSlots_itemID_b1=" & CStr(i) & "," & GoodHex(TrainerOptions(trainerIDselected).PlayerSlots(i).itemID_b1) & vbCrLf
          strSettings = strSettings & "Trainer_PlayerSlots_itemID_b2=" & CStr(i) & "," & GoodHex(TrainerOptions(trainerIDselected).PlayerSlots(i).itemID_b2) & vbCrLf
          strSettings = strSettings & "Trainer_PlayerSlots_xvalue=" & CStr(i) & "," & TrainerOptions(trainerIDselected).PlayerSlots(i).xvalue & vbCrLf
        Next i
    
        strSettings = strSettings & "Trainer_enabled=" & TrainerOptions(trainerIDselected).enabled & vbCrLf
        
        strSettings = strSettings & "END_Trainer=1" & vbCrLf
    End If
    
    
    tmpRes = OverwriteOnPathFileSimple(myPath, strSettings)
    If tmpRes <> 0 Then
        aRes = GiveGMmessage(idConnection, "Unable to save settings at " & myPath & " - Got error " & CStr(tmpRes), "BlackdProxy")
        DoEvents
        Exit Sub
    End If
    ' update memory
    AddSettingsOfChar charName, strSettings
    
    
    ' show sucess message

    
    aRes = SendLogSystemMessageToClient(idConnection, "Sucesfully saved settings of " & charName & " ; They will autoload everytime you login this char.")
    DoEvents
    
    Exit Sub
goterr:
    If GameConnected(idConnection) = True Then
        aRes = GiveGMmessage(idConnection, "Unable to save settings for this character. Got unexpected error " & CStr(Err.Number), "BlackdProxy")
        DoEvents
    End If
End Sub

Public Sub PreloadAllCharSettingsFromHardDisk()
  Dim res As Long
  #If FinalMode Then
  On Error GoTo goterr
  #End If
  Dim strFileName As String
  Dim myPath As String
  Dim fn As Integer
  Dim fs As scripting.FileSystemObject
  Dim f As scripting.Folder
  Dim f1 As scripting.File
  Dim currentSettingPath As String
  Dim currentSettingThing As String
  Dim currentCharName As String
  Dim strLine As String
  AutoloadUsable = True
  Set SettingsOfChar = New scripting.Dictionary

  myPath = App.path
  If (Right$(myPath, 1) <> "\") And (Right$(myPath, 1) <> "/") Then
    myPath = myPath & "\" & CteAutoloadSubfolder & "\"
  Else
    myPath = myPath & CteAutoloadSubfolder & "\"
  End If
  AutoloadPath = myPath
  
  Set fs = New scripting.FileSystemObject
  If fs.FolderExists(myPath) = False Then
    fs.CreateFolder (myPath)
    DoEvents
    If fs.FolderExists(myPath) = False Then
        AutoloadUsable = False
        Exit Sub
    End If
  End If
  
  Set f = fs.GetFolder(myPath)
  For Each f1 In f.Files
    strFileName = f1.name
    If (Len(strFileName) > 4) Then
        If LCase(Right$(strFileName, 4)) = ".txt" Then
            currentSettingPath = myPath & strFileName
            currentSettingThing = ""
            currentCharName = Left$(strFileName, Len(strFileName) - 4)
            fn = FreeFile
            Open currentSettingPath For Input As #fn
            While Not EOF(fn)
                Line Input #fn, strLine
                If Trim$(strLine) <> "" Then
                    currentSettingThing = currentSettingThing & strLine & vbCrLf
                End If
            Wend
            Close #fn
            AddSettingsOfChar currentCharName, currentSettingThing
        End If
    End If
  Next
  Exit Sub
goterr:
  AutoloadUsable = False
  Exit Sub
End Sub

Public Sub AddSettingsOfChar(ByVal strChar As String, ByVal strSettings As String)
  On Error GoTo goterr
  ' add item to dictionary
  Dim res As Boolean
  If AutoloadUsable = True Then
    SettingsOfChar.item(LCase(strChar)) = strSettings
  End If
  Exit Sub
goterr:
  LogOnFile "errors.txt", "Get error at AddSettingsOfChar : " & Err.Description
End Sub

Public Function GetSettingsOfChar(ByVal strChar As String) As String
  On Error GoTo goterr
  ' get the IPandport from server name
  Dim aRes As String
  Dim res As Boolean
  If AutoloadUsable = True Then
    If SettingsOfChar.Exists(LCase(strChar)) = True Then
      GetSettingsOfChar = SettingsOfChar.item(LCase(strChar))
    Else
      GetSettingsOfChar = ""
    End If
  End If
  Exit Function
goterr:
  LogOnFile "errors.txt", "Got error at GetSettingsOfChar : " & Err.Description
  GetSettingsOfChar = ""
End Function
