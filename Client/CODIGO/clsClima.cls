VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'// Autor: Emanuel Matías - Dunkan
'// Fecha: 12/04/2011
'// Contacto: emanuel.m@dunkancorp.com
'// Información: _
        Clase utilizada para calcular y cambiar el efecto de el clima. _
        Utilizando la hora de la PC
'// ESTÁ DE MAS DECIR QUE ESTA HARDCODEADO ! AJAJA
 
Option Explicit
 
Public Sub LoadState()
Dim m As Byte '// Mañana
Dim d As Byte '// Día
Dim T As Byte '// Tarde
Dim no As Byte '// Noche
Dim mn As Byte '// Media noche
Dim state As Byte

    For m = 5 To 10 '// Mañana
    If HorasClima = m Then
        state = 0
        CalculateRGB 203, 221, 228
        StateName = "Mañana"
    End If
    Next m
   
    For d = 11 To 15 '// Día
    If HorasClima = d Then
        state = 1
        CalculateRGB 244, 243, 246
        StateName = "Día"
    End If
    Next d
   
    For T = 16 To 19 '// Tarde
    If HorasClima = T Then
        state = 2
        CalculateRGB 223, 198, 194
        StateName = "Tarde"
    End If
    Next T
   
    For no = 20 To 23 '// Noche
    If HorasClima = no Then
        state = 3
        CalculateRGB 145, 147, 197
        StateName = "Noche"
    End If
    Next no
   
    For mn = 0 To 4  '// Media noche
    If HorasClima = mn Then
        state = 4
        CalculateRGB 97, 98, 177
        StateName = "Media Noche"
    End If
    Next mn

    Call DibujarTexto("Hora: " & HorasClima & ":" & MinutosClima & ":" & SegundosClima, 20, 10, , D3DColorARGB(200, 255, 255, 255))

End Sub

Private Sub CalculateRGB(r As Byte, g As Byte, b As Byte)
With RGBAlpha
    If .r > r Then
            .r = .r - 1
        ElseIf .r < r Then
            .r = .r + 1
        End If
        If .g > r Then
            .g = .r - 1
        ElseIf .g < r Then
            .g = .g + 1
        End If
        If .b > b Then
            .b = .b - 1
        ElseIf .b < b Then
            .b = .b + 1
        End If
        base_light = ARGB(.r, .g, .b, 255)
End With
End Sub


