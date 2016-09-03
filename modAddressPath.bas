Attribute VB_Name = "modAddressPath"
#Const FinalMode = 1
Option Explicit

Public Type AddressPath
         baseAddress As Long
         lastJumpIndex As Long
         jump() As Long
End Type
Public Function ReadAddressPath(ByVal strRawAddressPath As String) As AddressPath
        Dim res As AddressPath
        Dim parts() As String
        Dim lastPartIndex As Long
        Dim i As Long
        res.baseAddress = &H0
        res.lastJumpIndex = -1
        ReDim res.jump(0)
        parts = Split(strRawAddressPath, ">")
        res.baseAddress = parts(0)
        lastPartIndex = UBound(parts)
        If lastPartIndex > 0 Then
            res.lastJumpIndex = lastPartIndex - 1
            ReDim res.jump(lastPartIndex - 1)
            For i = 1 To lastPartIndex
                res.jump(i - 1) = parts(i)
            Next i
        End If
        ReadAddressPath = res
    End Function
    Public Function ReadCurrentAddress(ByVal pid As Long, ByRef adrPath As AddressPath, Optional ByVal desiredErrorValue As Long = -1, Optional ByVal readFinalValue As Boolean = True) As Long
        Dim res As Long
        Dim i As Long
        On Error GoTo gotErr
        res = 0
        If adrPath.baseAddress = 0 Then
            ReadCurrentAddress = desiredErrorValue
            Exit Function
        End If
  
            res = Memory_ReadLong(adrPath.baseAddress, pid, False)
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
                            res = Memory_ReadLong(res, pid, True)
                        End If
                    Else
                        res = Memory_ReadLong(res, pid, True)
                    End If
                Next i
            End If
            ReadCurrentAddress = res
            Exit Function
gotErr:
            ReadCurrentAddress = desiredErrorValue
End Function
