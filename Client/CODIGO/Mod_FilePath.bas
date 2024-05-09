Attribute VB_Name = "Mod_FilePath"
'*********************************
'KuviK AO 3.0 2011
'Modified by ShaFTeR
'Manejamos los directorios desde un modulo y sacamos las function viejas
'Extract to VbGore
Option Explicit
 
Public DirGraficos As String
Public DirInit As String
Public DirMapas As String
Public DirMidi As String
Public DirFoto As String
Public DirWav As String
Public DirMp3 As String
 
Public Sub InitFilePaths()
'*****************************************************************
'Set the common file paths
'More info: http://www.vbgore.com/CommonCode.FilePa ... tFilePaths
'*****************************************************************
 
    DirGraficos = App.Path & "\Graficos\"
    DirInit = App.Path & "\Init\"
    DirMapas = App.Path & "\Mapas\"
    DirMidi = App.Path & "\Midi\"
    DirFoto = App.Path & "\MapasScreen\"
    DirWav = App.Path & "\Wav\"
    DirMp3 = App.Path & "\MP3\"
End Sub
 
