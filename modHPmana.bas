Attribute VB_Name = "modHPmana"
#Const FinalMode = 1
Option Explicit
'...

Public Const cteHpManaSeparator = "^"
Public Type TypeHPmanaConfig
    charName As String
    hpVal As Byte
    hpACTION As String
    manaVal As Byte
    manaACTION As String
    baseHpVal As Byte
    baseManaVal As Byte
End Type

Public HPmanaConfig() As TypeHPmanaConfig
Public lastLootOrder() As Long
Public LastHealTime() As Long
Public LastCavebotTime() As Long
Public HPmanaRECAST As Long
Public HPmanaRECAST2 As Long
Public RunemakerChaos As Long
Public RunemakerChaos2 As Long
Public CavebotRECAST As Long
Public CavebotRECAST2 As Long
Public LimitRandomizator As Long

Private Sub parseConfigLine(strLine As String)
    On Error GoTo ignoreThis
    Dim strL As String
    Dim pos1 As Long
    Dim pos2 As Long
    Dim charName As String
    Dim hpVal As String
    Dim hpACTION As String
    Dim manaVal As String
    Dim manaACTION As String
    Dim sigPos As Long
    Dim b1 As Byte
    Dim b2 As Byte
    strL = Trim$(strLine)
    If strLine = "" Then
        Exit Sub
    Else
        pos1 = 1
        pos2 = InStr(pos1, strL, cteHpManaSeparator)
        charName = Mid$(strL, pos1, pos2 - pos1)
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strL, cteHpManaSeparator)
        hpVal = Mid$(strL, pos1, pos2 - pos1)
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strL, cteHpManaSeparator)
        hpACTION = Mid$(strL, pos1, pos2 - pos1)
        pos1 = pos2 + 1
        pos2 = InStr(pos1, strL, cteHpManaSeparator)
        manaVal = Mid$(strL, pos1, pos2 - pos1)
        pos1 = pos2 + 1
        pos2 = Len(strL) + 1
        manaACTION = Mid$(strL, pos1, pos2 - pos1)
        b1 = CByte(hpVal)
        b2 = CByte(manaVal)
        AddHPmanaSetting charName, b1, hpACTION, b2, manaACTION
    End If
ignoreThis:
    Exit Sub
End Sub

Public Sub UpdateHPmanaSetting(strCharname As String, bHP As Byte, strHPact As String, bMANA As Byte, strMANAact As String, idLine As Long)
    Dim ub As Long
    Dim sigPos As Long
    Dim i As Long
    ub = UBound(HPmanaConfig)
    If idLine <= ub Then
        HPmanaConfig(idLine).charName = strCharname
        HPmanaConfig(idLine).hpACTION = strHPact
        HPmanaConfig(idLine).hpVal = bHP
        HPmanaConfig(idLine).manaACTION = strMANAact
        HPmanaConfig(idLine).manaVal = bMANA
        HPmanaConfig(idLine).baseHpVal = bHP
        HPmanaConfig(idLine).baseManaVal = bMANA
    End If
End Sub


Public Sub AddHPmanaSetting(strCharname As String, bHP As Byte, strHPact As String, bMANA As Byte, strMANAact As String)
    Dim ub As Long
    Dim sigPos As Long
    Dim i As Long
    If strCharname <> "" Then
        ub = UBound(HPmanaConfig)
        sigPos = ub + 1
        ReDim Preserve HPmanaConfig(sigPos)
        HPmanaConfig(sigPos).charName = strCharname
        HPmanaConfig(sigPos).hpACTION = strHPact
        HPmanaConfig(sigPos).hpVal = bHP
        HPmanaConfig(sigPos).manaACTION = strMANAact
        HPmanaConfig(sigPos).manaVal = bMANA
        HPmanaConfig(sigPos).baseHpVal = bHP
        HPmanaConfig(sigPos).baseManaVal = bMANA
    End If
End Sub


Public Function LoadHPmanaConfig() As String
    On Error GoTo gotErr
  Dim fso As Scripting.FileSystemObject
  Set fso = New Scripting.FileSystemObject
  Dim fn As Integer
  Dim strLine As String
  If fso.FileExists(App.Path & "\HPmana.cfg") = False Then
    ReDim HPmanaConfig(0)
    Exit Function
  End If
  Set fso = Nothing
  ReDim HPmanaConfig(0)
  fn = FreeFile
  Open App.Path & "\HPmana.cfg" For Input As #fn
    While EOF(fn) = False
      Line Input #fn, strLine
      parseConfigLine strLine
    Wend
  Close fn
  DisplayLoadedHPmanaConfig
  LoadHPmanaConfig = ""
  Exit Function
gotErr:
  LoadHPmanaConfig = "(at load) error code " & CStr(Err.Number) & " : " & Err.Description
End Function


Public Function SaveHPmanaConfig() As String
    On Error GoTo gotErr
    Dim i As Long
    Dim ult As Long
    Dim strHD As String
    Dim strLine As String
    ult = UBound(HPmanaConfig)
    strHD = ""
    For i = 1 To ult
        strLine = HPmanaConfig(i).charName & cteHpManaSeparator & CStr(CLng(HPmanaConfig(i).baseHpVal)) & _
         cteHpManaSeparator & HPmanaConfig(i).hpACTION & cteHpManaSeparator & CStr(CLng(HPmanaConfig(i).baseManaVal)) & _
         cteHpManaSeparator & HPmanaConfig(i).manaACTION
        If strHD = "" Then
            strHD = strLine
        Else
            strHD = strHD & vbCrLf & strLine
        End If
    Next i
    OverwriteOnFile "HPmana.cfg", strHD
    SaveHPmanaConfig = ""
    Exit Function
gotErr:
    SaveHPmanaConfig = "(at save) error code " & CStr(Err.Number) & " : " & Err.Description
End Function

Public Sub DeleteAllSettings()
    ReDim HPmanaConfig(0)
End Sub

Public Sub deleteHPmanaSetting(l As Long)
    Dim i As Long
    Dim ub As Long
    If l > 0 Then
        ub = UBound(HPmanaConfig) - 1
        For i = l To ub
            HPmanaConfig(i).charName = HPmanaConfig(i + 1).charName
            HPmanaConfig(i).hpACTION = HPmanaConfig(i + 1).hpACTION
            HPmanaConfig(i).hpVal = HPmanaConfig(i + 1).hpVal
            HPmanaConfig(i).manaACTION = HPmanaConfig(i + 1).manaACTION
            HPmanaConfig(i).manaVal = HPmanaConfig(i + 1).manaVal
            HPmanaConfig(i).baseHpVal = HPmanaConfig(i + 1).baseHpVal
            HPmanaConfig(i).baseManaVal = HPmanaConfig(i + 1).baseManaVal
        Next i
        ReDim Preserve HPmanaConfig(ub)
    End If
End Sub

Public Sub DisplayLoadedHPmanaConfig()
    Dim lngAm As Long
    Dim i As Long
    lngAm = UBound(HPmanaConfig)
    frmHPmana.gridHPmana.Redraw = False
    frmHPmana.gridHPmana.Rows = lngAm + 1
    For i = 1 To lngAm
        frmHPmana.gridHPmana.Row = i
        frmHPmana.gridHPmana.TextMatrix(i, 1) = HPmanaConfig(i).charName
        frmHPmana.gridHPmana.Col = 1
        frmHPmana.gridHPmana.CellAlignment = flexAlignLeftCenter
        frmHPmana.gridHPmana.TextMatrix(i, 2) = CStr(CLng(HPmanaConfig(i).baseHpVal))
        frmHPmana.gridHPmana.Col = 2
        frmHPmana.gridHPmana.CellAlignment = flexAlignLeftCenter
        frmHPmana.gridHPmana.TextMatrix(i, 3) = HPmanaConfig(i).hpACTION
        frmHPmana.gridHPmana.Col = 3
        frmHPmana.gridHPmana.CellAlignment = flexAlignLeftCenter
        frmHPmana.gridHPmana.TextMatrix(i, 4) = CStr(CLng(HPmanaConfig(i).baseManaVal))
        frmHPmana.gridHPmana.Col = 4
        frmHPmana.gridHPmana.CellAlignment = flexAlignLeftCenter
        frmHPmana.gridHPmana.TextMatrix(i, 5) = HPmanaConfig(i).manaACTION
        frmHPmana.gridHPmana.Col = 5
        frmHPmana.gridHPmana.CellAlignment = flexAlignLeftCenter
    Next i

    frmHPmana.gridHPmana.Row = frmHPmana.gridHPmana.Rows - 1
    frmHPmana.gridHPmana.Col = 1
    frmHPmana.gridHPmana.ColSel = 5
 
    frmHPmana.gridHPmana.Redraw = True
End Sub

Public Function GotCustomHPsettings(idConnection As Integer) As Boolean
    On Error GoTo gotErr
    Dim i As Long
    Dim ult As Long
    Dim lstrchar As String
    Dim act As String
    lstrchar = LCase(CharacterName(idConnection))
    ult = UBound(HPmanaConfig)
    For i = 1 To ult
        If LCase(HPmanaConfig(i).charName) = lstrchar Then
            act = LCase(HPmanaConfig(i).hpACTION)
            If ((act <> "") And (act <> "no action")) Then
                GotCustomHPsettings = True
                Exit Function
            End If
        End If
    Next i
    GotCustomHPsettings = False
    Exit Function
gotErr:
    GotCustomHPsettings = False
End Function


Public Function GotCustomMANAsettings(idConnection As Integer) As Boolean
    On Error GoTo gotErr
    Dim i As Long
    Dim ult As Long
    Dim lstrchar As String
    Dim act As String
    lstrchar = LCase(CharacterName(idConnection))
    ult = UBound(HPmanaConfig)
    For i = 1 To ult
        If LCase(HPmanaConfig(i).charName) = lstrchar Then
            act = LCase(HPmanaConfig(i).manaACTION)
            If ((act <> "") And (act <> "no action")) Then
                GotCustomMANAsettings = True
                Exit Function
            End If
        End If
    Next i
    GotCustomMANAsettings = False
    Exit Function
gotErr:
    GotCustomMANAsettings = False
End Function


Public Function ChaotizeRechargeLevel(ByVal baseLevel As Long)
    Dim lngChaos As Long
    lngChaos = randomNumberBetween(baseLevel - LimitRandomizator, baseLevel + LimitRandomizator)
    If lngChaos > 100 Then
     lngChaos = 100
    End If
    If lngChaos < 1 Then
     lngChaos = 1
    End If
    ChaotizeRechargeLevel = lngChaos
End Function

