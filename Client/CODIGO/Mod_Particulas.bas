Attribute VB_Name = "Mod_Particulas"
Public particula(1 To 500) As Stream
Public TotalStreams As Long
Public Actual As Byte
 
Public Type Stream
    Name As String
    MapeZ As Integer
    VarZ As Integer
    MapX As Integer
    Mapy As Integer
    VarX As Single
    VarY As Single
    friction As Long
    AlphaInicial As Byte
    RedInicial As Byte
    GreenInicial As Byte
    BlueInicial As Byte
    AlphaFinal As Byte
    RedFinal As Byte
    GreenFinal As Byte
    BlueFinal As Byte
    NumOfParticles As Integer
    gravity As Single
    texture As Long
    Zize As Single
    Life As Integer
    angle As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    grav_strength As Long
    bounce_strength As Long
End Type
Sub CargarParticulas()
StreamFile = App.Path & "\Particles.ini"
TotalStreams = Val(General_Var_Get(StreamFile, "INIT", "Total"))
 
For loopc = 1 To TotalStreams
    particula(loopc).Name = General_Var_Get(StreamFile, Val(loopc), "Name")
    particula(loopc).VarX = General_Var_Get(StreamFile, Val(loopc), "VarX")
    particula(loopc).VarY = General_Var_Get(StreamFile, Val(loopc), "VarY")
    particula(loopc).VarZ = General_Var_Get(StreamFile, Val(loopc), "VarZ")
    particula(loopc).AlphaInicial = General_Var_Get(StreamFile, Val(loopc), "AlphaInicial")
    particula(loopc).RedInicial = General_Var_Get(StreamFile, Val(loopc), "RedInicial")
    particula(loopc).GreenInicial = General_Var_Get(StreamFile, Val(loopc), "GreenInicial")
    particula(loopc).BlueInicial = General_Var_Get(StreamFile, Val(loopc), "BlueInicial")
    particula(loopc).AlphaFinal = General_Var_Get(StreamFile, Val(loopc), "AlphaFinal")
    particula(loopc).RedFinal = General_Var_Get(StreamFile, Val(loopc), "RedFinal")
    particula(loopc).GreenFinal = General_Var_Get(StreamFile, Val(loopc), "GreenFinal")
    particula(loopc).BlueFinal = General_Var_Get(StreamFile, Val(loopc), "BlueFinal")
    particula(loopc).NumOfParticles = General_Var_Get(StreamFile, Val(loopc), "NumOfParticles")
    particula(loopc).gravity = General_Var_Get(StreamFile, Val(loopc), "Gravity")
    particula(loopc).texture = General_Var_Get(StreamFile, Val(loopc), "texture")
    particula(loopc).Zize = General_Var_Get(StreamFile, Val(loopc), "Zize")
    particula(loopc).Life = General_Var_Get(StreamFile, Val(loopc), "Life")
    particula(loopc).friction = General_Var_Get(StreamFile, Val(loopc), "Friction")
    particula(loopc).angle = General_Var_Get(StreamFile, Val(loopc), "Angle")
    particula(loopc).spin = General_Var_Get(StreamFile, Val(loopc), "Spin")
    particula(loopc).spin_speedL = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedL")
    particula(loopc).spin_speedH = General_Var_Get(StreamFile, Val(loopc), "Spin_SpeedH")
    particula(loopc).grav_strength = General_Var_Get(StreamFile, Val(loopc), "Grav_Strength")
    particula(loopc).bounce_strength = General_Var_Get(StreamFile, Val(loopc), "Bounce_Strength")
   
Next loopc
End Sub
 
Public Sub General_Var_Write(ByVal file As String, ByVal Main As String, ByVal var As String, ByVal value As String)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, var, value, file
End Sub
 
Public Function General_Var_Get(ByVal file As String, ByVal Main As String, ByVal var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim l As Long
    Dim Char As String
    Dim sSpaces As String 'Input that the program will retrieve
    Dim szReturn As String 'Default value if the string is not found
   
    szReturn = ""
   
    sSpaces = Space$(5000)
   
    getprivateprofilestring Main, var, szReturn, sSpaces, Len(sSpaces), file
   
    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function
