Attribute VB_Name = "modGeneral"
Option Explicit
 
' API Functions
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Public Declare Function GetLastError Lib "kernel32" () As Long
 
' API Data
Public Type SECURITY_ATTRIBUTES
    nLength                 As Long ' Specifies the length, in bytes, of this structure (12 bytes).
    lpSecurityDescriptor    As Long ' Points to a security descriptor for the object that controls the sharing of it.
    bInheritHandle          As Long ' Specifies whether the returned handle is inherited when a new process is created.
End Type
