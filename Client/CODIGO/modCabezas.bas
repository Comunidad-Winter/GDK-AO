Attribute VB_Name = "modCabezas"
Option Explicit
Public MinEleccion As Integer, MaxEleccion As Integer
Public Actual As Integer
 
Sub DameOpciones()
 
Dim i As Integer
 
Select Case frmCrearPersonaje.lstGenero.List(frmCrearPersonaje.lstGenero.ListIndex)
   Case "Hombre"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = 1
                MaxEleccion = 30
                MinEleccion = 1
            Case "Elfo"
                Actual = 101
                MaxEleccion = 113
                MinEleccion = 101
            Case "Elfo Oscuro"
                Actual = 202
                MaxEleccion = 209
                MinEleccion = 202
            Case "Enano"
                Actual = 301
                MaxEleccion = 305
                MinEleccion = 301
            Case "Gnomo"
                Actual = 401
                MaxEleccion = 406
                MinEleccion = 401
            Case Else
                Actual = 30
                MaxEleccion = 30
                MinEleccion = 30
        End Select
   Case "Mujer"
        Select Case frmCrearPersonaje.lstRaza.List(frmCrearPersonaje.lstRaza.ListIndex)
            Case "Humano"
                Actual = 70
                MaxEleccion = 76
                MinEleccion = 70
            Case "Elfo"
                Actual = 170
                MaxEleccion = 176
                MinEleccion = 170
            Case "Elfo Oscuro"
                Actual = 270
                MaxEleccion = 280
                MinEleccion = 270
            Case "Gnomo"
                Actual = 470
                MaxEleccion = 474
                MinEleccion = 470
            Case "Enano"
                Actual = 370
                MaxEleccion = 373
                MinEleccion = 370
            Case Else
                Actual = 70
                MaxEleccion = 70
                MinEleccion = 70
        End Select
End Select
 
frmCrearPersonaje.HeadView.Cls
Call DrawGrhtoHdc(frmCrearPersonaje.HeadView.hdc, HeadData(Actual).Head(3).grhindex, 8, 5)
 
End Sub
Public Sub DrawGrhtoHdc(desthDC As Long, ByVal grh_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional transparent As Boolean = False)
 
    On Error Resume Next
   
    Dim file_path As String
    Dim src_x As Integer
    Dim src_y As Integer
    Dim src_width As Integer
    Dim src_height As Integer
    Dim hdcsrc As Long
    Dim MaskDC As Long
    Dim PrevObj As Long
    Dim PrevObj2 As Long
 
    If grh_index <= 0 Then Exit Sub
 
    If GrhData(grh_index).NumFrames <> 1 Then
        grh_index = GrhData(grh_index).Frames(1)
    End If
 
        file_path = App.Path & "\GRAFICOS\" & GrhData(grh_index).FileNum & ".bmp"
       
        src_x = GrhData(grh_index).sX
        src_y = GrhData(grh_index).sY
        src_width = GrhData(grh_index).pixelWidth
        src_height = GrhData(grh_index).pixelHeight
       
       
           
        hdcsrc = CreateCompatibleDC(desthDC)
         
        PrevObj = SelectObject(hdcsrc, LoadPicture(file_path))
       
 
        BitBlt desthDC, screen_x, screen_y, src_width, src_height, hdcsrc, src_x, src_y, vbSrcCopy
 
        DeleteDC hdcsrc
 
 
End Sub
 
