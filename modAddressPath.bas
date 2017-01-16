Attribute VB_Name = "modAddressPath"
#Const FinalMode = 1
Option Explicit

Public Type AddressPath
         baseModule As String
         baseAddress As Long
         lastJumpIndex As Long
         jump() As Long
End Type

Public Function AddressPathToString(ByRef adrPath As AddressPath) As String
    Dim res As String
    Dim i As Integer
    If (adrPath.baseModule = "") Then
        res = "&H" & Hex(adrPath.baseAddress)
        For i = 0 To adrPath.lastJumpIndex
        res = res & " > " & adrPath.jump(i)
        Next i
    Else
        res = """" & adrPath.baseModule & """" & " + " & Hex(adrPath.baseAddress)
        For i = 0 To adrPath.lastJumpIndex
            res = res & " > " & Hex(adrPath.jump(i))
        Next i
    End If
    AddressPathToString = res
End Function
Public Function ReadAddressPath(ByVal strRawAddressPath As String) As AddressPath
        Dim res As AddressPath
        Dim parts() As String
        Dim parts0() As String
        Dim lastPartIndex As Integer
        Dim i As Integer
        Dim modNameParts() As String
        Dim newMode As Boolean
        strRawAddressPath = Replace(strRawAddressPath, " ", "")
        res.baseAddress = &H0
        res.lastJumpIndex = -1
        ReDim res.jump(0)
        parts = Split(strRawAddressPath, ">")
        parts0 = Split(parts(0), "+")
        If UBound(parts0) = 0 Then
            newMode = False
            res.baseAddress = parts(0)
            res.baseModule = ""
        Else
            newMode = True
            res.baseAddress = CLng("&H" & parts0(1))
            modNameParts = Split(parts0(0), """")
            res.baseModule = modNameParts(1)
        End If
        lastPartIndex = UBound(parts)
        If lastPartIndex > 0 Then
            res.lastJumpIndex = lastPartIndex - 1
            ReDim res.jump(lastPartIndex - 1)
            For i = 1 To lastPartIndex
                If (newMode) Then
                    res.jump(i - 1) = CLng("&H" & parts(i))
                Else
                    res.jump(i - 1) = CLng(parts(i))
                End If
            Next i
        End If
        ReadAddressPath = res
        Exit Function

    End Function
    
    Public Function ReadCurrentAddressFLOAT(ByVal pid As Long, ByRef adrPath As AddressPath, Optional ByVal desiredErrorValue As Single = -1) As Single
        On Error GoTo goterr
        Dim rawval As Long
        Dim valLong As Long
        Dim valSingle As Single
        rawval = ReadCurrentAddress(pid, adrPath, -1, True)
        If (rawval = -1) Then
            ReadCurrentAddressFLOAT = desiredErrorValue
            Exit Function
        End If
        valSingle = Long2Float(rawval)
        ReadCurrentAddressFLOAT = valSingle
        Exit Function
goterr:
        ReadCurrentAddressFLOAT = desiredErrorValue
    End Function
    
    Public Function ReadCurrentAddressDOUBLE(ByVal pid As Long, ByRef adrPath As AddressPath, Optional ByVal desiredErrorValue As Long = -1) As Long
        On Error GoTo goterr
        Dim adr As Long
        Dim val8bytes As Double
        Dim valRounded As Long
        adr = ReadCurrentAddress(pid, adrPath, -1, False)
        If (adr = -1) Then
            ReadCurrentAddressDOUBLE = desiredErrorValue
            Exit Function
        End If
        val8bytes = QMemory_ReadDouble(pid, adr)
        valRounded = Math.Round(val8bytes)
        ReadCurrentAddressDOUBLE = valRounded
        Exit Function
goterr:
        ReadCurrentAddressDOUBLE = desiredErrorValue
    End Function
    
    Public Function ReadCurrentAddress(ByVal pid As Long, ByRef adrPath As AddressPath, Optional ByVal desiredErrorValue As Long = -1, Optional ByVal readFinalValue As Boolean = True) As Long
       Dim res As Long
        Dim realBase As Long
        Dim reqKey As String
        Dim i As Integer
        On Error GoTo goterr
        res = 0
        realBase = 0
        If adrPath.baseAddress = 0 Then
            ReadCurrentAddress = desiredErrorValue
            Exit Function
        End If
            If adrPath.baseModule = "" Then
                ' Old format
                realBase = adrPath.baseAddress
                If adrPath.lastJumpIndex = -1 Then
                    If (readFinalValue) Then
                        res = Memory_ReadLong(realBase, pid, False)
                    Else
                        res = realBase
                    End If
                    ReadCurrentAddress = res
                    Exit Function
                Else
                    res = Memory_ReadLong(realBase, pid, False)
                End If
            Else
                ' New format since Tibia 11
                If (moduleDictionary Is Nothing) Then
                    Set moduleDictionary = New Scripting.Dictionary
                End If
                reqKey = adrPath.baseModule & CStr(pid)
                If moduleDictionary.Exists(reqKey) Then
                    realBase = CLng(moduleDictionary(adrPath.baseModule & CStr(pid)))
                    realBase = realBase + adrPath.baseAddress
                    If adrPath.lastJumpIndex = -1 Then
                        If (readFinalValue) Then
                            res = QMemory_Read4Bytes(pid, realBase)
                        Else
                            res = realBase
                        End If
                        ReadCurrentAddress = res
                        Exit Function
                    Else
                        res = QMemory_Read4Bytes(pid, realBase)
                    End If
                Else
                    ' refresh all base addresses and region sizes
                    GetAllBaseAddressesAndRegionSizes tibiamainname, tibiaclassname
                    ' second try
                    If moduleDictionary.Exists(reqKey) Then
                        realBase = CLng(moduleDictionary(adrPath.baseModule & CStr(pid)))
                        realBase = realBase + adrPath.baseAddress
                        If adrPath.lastJumpIndex = -1 Then
                            If (readFinalValue) Then
                                res = QMemory_Read4Bytes(pid, realBase)
                            Else
                                res = realBase
                            End If
                            ReadCurrentAddress = res
                            Exit Function
                        Else
                            res = QMemory_Read4Bytes(pid, realBase)
                        End If
                    Else
                      ReadCurrentAddress = desiredErrorValue
                      Exit Function
                    End If
                End If

            End If
            If adrPath.lastJumpIndex >= 0 Then
                For i = 0 To adrPath.lastJumpIndex
                    ' follow the path of jumps
                    If (res <= 1000) Then ' Detects bad address
                        ReadCurrentAddress = desiredErrorValue
                        Exit Function
                    End If
                    res = res + adrPath.jump(i)
                    If i = adrPath.lastJumpIndex Then
                        If readFinalValue = True Then
                            res = QMemory_Read4Bytes(pid, res)
                        End If
                    Else
                        res = QMemory_Read4Bytes(pid, res)
                    End If
                Next i
            End If
            ReadCurrentAddress = res
            Exit Function
goterr:
Debug.Print "ReadCurrentAddress failure: " & Err.Description
            ReadCurrentAddress = desiredErrorValue
End Function
