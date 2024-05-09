VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsThread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
 
' API Constants
Private Const WAIT_OBJECT_0    As Long = 0
Private Const CREATE_SUSPENDED As Long = &H4     ' The thread is created in a suspended state.
Private Const STACK_SIZE_PARAM As Long = &H10000 ' The dwStackSize parameter specifies the initial reserve size of the stack.
 
' API Functions
Private Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function ResumeThread Lib "kernel32" (ByVal hHandle As Long) As Long
 
' Member data
Private ThreadDefaultSecurityStruct As SECURITY_ATTRIBUTES
Private ThreadHandle    As Long
Private ThreadID        As Long
Private Sub Class_Initialize()
    ' Set default attributes
    ThreadDefaultSecurityStruct.nLength = 12
    ThreadDefaultSecurityStruct.lpSecurityDescriptor = 0
    ThreadDefaultSecurityStruct.bInheritHandle = True
End Sub
 
Private Sub Class_Terminate()
    ' Close the thread
    Call CloseHandle(ThreadHandle)
End Sub
Public Sub Create(ByVal PtrFunc As Long, ByVal PtrArg As Long)
    ' Create thread in a suspended state
    ThreadHandle = CreateThread(ThreadDefaultSecurityStruct, 0, PtrFunc, PtrArg, CREATE_SUSPENDED, ThreadID)
End Sub
Public Sub Kill()
    ' Kill the thread
    Call TerminateThread(ThreadHandle, 0)
End Sub
Public Sub DoResume()
    ' Resume the thread
    Call ResumeThread(ThreadHandle)
End Sub
Public Function WaitForDeath() As Boolean
    ' Wait
    WaitForDeath = (WaitForSingleObject(ThreadHandle, 2000) = WAIT_OBJECT_0)
End Function
 
