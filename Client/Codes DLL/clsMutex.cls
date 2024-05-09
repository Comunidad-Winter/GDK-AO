VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMutex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
' API Constants
Private Const ERROR_ALREADY_EXISTS As Long = 183&
Private Const MUTEX_ALL_ACCESS     As Long = &H1F0001
 
' API Functions
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (lpMutexAttributes As SECURITY_ATTRIBUTES, ByVal bInitialOwner As Boolean, ByVal lpName As String) As Long
Private Declare Function OpenMutex Lib "kernel32" Alias "OpenMutexA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
 
' Member data
Private MutexDefaultSecurityStruct As SECURITY_ATTRIBUTES
Private MutexHandle     As Long
Private MutexName       As String
Private Sub Class_Initialize()
    ' Set default attributes
    MutexDefaultSecurityStruct.nLength = 12
    MutexDefaultSecurityStruct.lpSecurityDescriptor = 0
    MutexDefaultSecurityStruct.bInheritHandle = True
End Sub
 
Private Sub Class_Terminate()
    ' Destroy the mutex
    Call CloseHandle(MutexHandle)
End Sub
 
Public Function Create(ByVal Name As String) As Boolean
 
    ' Create the mutex
    MutexName = Name
   
    MutexHandle = OpenMutex(MUTEX_ALL_ACCESS, 0, MutexName)
   
    If Not MutexHandle Then
        MutexHandle = CreateMutex(MutexDefaultSecurityStruct, 0, MutexName)
    Else
        If GetLastError = ERROR_ALREADY_EXISTS Then
            Debug.Print "The mutex already exist"
            Create = False
            Exit Function
        End If
    End If
   
    Create = True
   
End Function
 
Public Sub DoUnLock()
    ' Release it
    Call ReleaseMutex(MutexHandle)
End Sub
Public Sub DoLock()
    ' Lock it
    Call WaitForSingleObject(MutexHandle, 0)
End Sub
 
