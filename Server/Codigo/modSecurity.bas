Attribute VB_Name = "modSecurity"
Option Explicit

'@Param a   Positive long
'@Param e   Positive long
'@Param m   Positive long in range  2 <= m <= Int(((65536 ^ 2) / 2 - 1) ^ (1 / 2)) + 1
Public Function AexpEmodM(ByVal a As Long, ByVal e As Long, ByVal m As Long) As Long
Dim res As Long

res = 1
a = a Mod m

Do While (e > 0) And (a <> 1) And (res <> 0)
    If e And &H1 Then
        res = (res * a) Mod m
    End If
    
    e = e \ 2
    a = (a * a) Mod m
Loop

AexpEmodM = res
End Function

Private Sub Swap(ByRef N1 As Byte, ByRef N2 As Byte)
Dim temp As Byte

temp = N1
N1 = N2
N2 = temp
End Sub

Private Sub RC4_KSA(ByRef S() As Byte, ByRef Key As String)
Dim i As Long
Dim j As Long
Dim KeyLen As Byte
Dim k() As Byte
Dim temp As Byte

If LenB(Key) = 0 Then Key = "NoKey"

KeyLen = Len(Key)
k() = StrConv(Key, vbFromUnicode)

For i = 0 To 255
    S(i) = i
Next i

For i = 0 To 255
    j = (j + S(i) + k(i Mod KeyLen)) And 255
    
    'This is more faster than Swap method
    temp = S(i)
    S(i) = S(j)
    S(j) = temp
    'Call Swap(s(i), s(j))
Next i
End Sub

Public Function RC4_EncryptString(ByRef Str As String, ByRef Key As String) As String
Dim S() As Byte

S() = StrConv(Str, vbFromUnicode)

Call RC4_EncryptByte(S(), Key)

RC4_EncryptString = StrConv(S(), vbUnicode)
End Function

Private Sub RC4_EncryptByte(ByRef Arr() As Byte, ByRef Key As String)
Dim S(255) As Byte
Dim i As Long
Dim j As Long
Dim k As Long
Dim temp As Byte

Call RC4_KSA(S(), Key)

For k = 0 To UBound(Arr)
    i = (i + 1) And 255
    j = (j + S(i)) And 255
    
    'This is more faster than Swap method
    temp = S(i)
    S(i) = S(j)
    S(j) = temp
    'Call Swap(s(i), s(j))
    
    Arr(k) = (S((CInt(S(i)) + CInt(S(j))) And 255) Xor Arr(k))
Next k
End Sub
