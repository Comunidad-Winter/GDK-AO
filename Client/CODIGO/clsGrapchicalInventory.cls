VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrapchicalInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 
Option Explicit
 
Dim InventoryOffset As Long             'Number of lines we scrolled down from topmost
Dim InvSelectedItem As Long             'Currently selected item
 
 
Dim UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory    'User's inventory

Dim WithEvents InventoryWindowNpc As PictureBox
Attribute InventoryWindowNpc.VB_VarHelpID = -1
Dim WithEvents InventoryWindowUser As PictureBox
Attribute InventoryWindowUser.VB_VarHelpID = -1

Dim WithEvents InventoryWindow As PictureBox    'Placeholder where to render the inventory
Attribute InventoryWindow.VB_VarHelpID = -1

Dim InvRect As RECT
Dim InvD3DRect As D3DRECT
Dim color(3) As Long

Private Sub Class_Initialize()
'***************************************************
'Last Modify Date: 18/05/11
'Author Modify: ShaFTeR
'Extracted to: GSZ-AO by •Parra
'***************************************************
    Set InventoryWindow = frmMain.picInv
    InventoryWindow.AutoRedraw = True
    InventoryWindow.ScaleMode = 3
    
    InvRect.Top = 0: InvD3DRect.Y1 = 0
    InvRect.Left = 0: InvD3DRect.X1 = 0
    InvRect.bottom = 160: InvD3DRect.Y2 = 160
    InvRect.Right = 160: InvD3DRect.X2 = 160
   
engine.Graphics_ResetColor color()
End Sub
 
Private Sub Class_Terminate()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'
'***************************************************
End Sub
 
Public Sub ScrollInventory(ByVal up As Boolean)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Scrolls the graphical inventory up / down
'***************************************************
    'Only move if there is something to show up /down
    If up Then
        If (InventoryOffset + (InventoryWindow.ScaleHeight \ TilePixelHeight)) * (InventoryWindow.ScaleWidth \ TilePixelWidth) < MAX_INVENTORY_SLOTS Then _
            InventoryOffset = InventoryOffset + 1
    Else
        If InventoryOffset > 0 Then _
            InventoryOffset = InventoryOffset - 1
    End If
End Sub
 
Public Sub SelectGold()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Sets the gold as the currently selected item
'***************************************************
    Dim prevSelectedItem As Long
   
    'Store preivously selected item
    prevSelectedItem = InvSelectedItem
   
    'Select the gold
    InvSelectedItem = FLAGORO
   
 
End Sub
 
Public Property Get SelectedItem() As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item index
'***************************************************
    SelectedItem = InvSelectedItem
End Property
 
Public Property Get MaxHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the max hit of the selected item
'***************************************************
    MaxHit = UserInventory(slot).MaxHit
End Property
 
Public Property Get MinHit(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the min hit of the selected item
'***************************************************
    MinHit = UserInventory(slot).MinHit
End Property
 
Public Property Get Def(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the defense of the selected item
'***************************************************
    Def = UserInventory(slot).Def
End Property
 
Public Property Get grhindex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the grh index of the selected item
'***************************************************
    grhindex = UserInventory(slot).grhindex
End Property
 
Public Property Get Valor(ByVal slot As Byte) As Single
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the value of the selected item
'***************************************************
    Valor = UserInventory(slot).Valor
End Property
 
Public Property Get Amount(ByVal slot As Byte) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's amount
'***************************************************
    If slot = FLAGORO Then
        Amount = UserGLD
    ElseIf slot >= LBound(UserInventory) And slot <= UBound(UserInventory) Then
        Amount = UserInventory(slot).Amount
    End If
End Property
 
Public Property Get OBJIndex(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object index
'***************************************************
    OBJIndex = UserInventory(slot).OBJIndex
End Property
 
Public Property Get OBJType(ByVal slot As Byte) As Integer
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's object type
'***************************************************
    OBJType = UserInventory(slot).OBJType
End Property
 
Public Property Get ItemName(ByVal slot As Byte) As String
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves the selected item's name
'***************************************************
    ItemName = UserInventory(slot).Name
End Property
 
Public Property Get Equipped(ByVal slot As Byte) As Boolean
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Retrieves True if the item at the given pos is eqiupped
'***************************************************
    Equipped = UserInventory(slot).Equipped
End Property
 
Public Sub SetItem(ByVal slot As Byte, ByVal eOBJIndex As Integer, ByVal eAmount As Integer, ByVal eEquipped As Byte, _
                        ByVal eGrhIndex As Integer, ByVal eObjType As Integer, ByVal eMaxHit As Integer, ByVal eMinHit As Integer, _
                        ByVal eDef As Integer, ByVal eValor As Single, ByVal eName As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/04/06
'Sets all data for a given inventory slot
'***************************************************
    If slot < 1 Or slot > MAX_INVENTORY_SLOTS Then Exit Sub
   
    With UserInventory(slot)
        .Amount = eAmount
        .Def = eDef
        .Equipped = eEquipped
        .grhindex = eGrhIndex
        .MaxHit = eMaxHit
        .MinHit = eMinHit
        .Name = eName
        .OBJIndex = eOBJIndex
        .OBJType = eObjType
        .Valor = eValor
    End With
   
End Sub
 
Private Function ClickItem(ByVal X As Long, ByVal Y As Long) As Long
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Selects the item clicked if it's valid and return's it's index
'***************************************************
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    temp_x = X \ 32
    temp_y = Y \ 32
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
   
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        'Make sure slot isn't empty
        If UserInventory(TempItem).grhindex Then
            ClickItem = TempItem
        Else
            ClickItem = 0
        End If
    End If
End Function
 
Public Sub UpdateInventory()
'*****************************************************
'Date: 05/09/09
'Author: Parra
'Description: Updates the inventory only when needed
'Last modify by ShaFTeR
'*****************************************************

DoEvents

    D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0, 0
    D3DDevice.BeginScene

    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer
    Dim s(3) As Long

s(0) = D3DColorXRGB(255, 255, 255)
s(1) = D3DColorXRGB(255, 255, 255)
s(2) = D3DColorXRGB(255, 255, 255)
s(3) = D3DColorXRGB(255, 255, 255)

Call engine.Device_Box_Textured_Render(19973, 0, 0, 160, 160, s, 0, 0)

    For i = 1 To UBound(UserInventory)
        If UserInventory(i).grhindex Then
            X = ((i - 1) Mod (InventoryWindow.Width / 32)) * 32 + 2
            Y = ((i - 1) \ (InventoryWindow.Width / 32)) * 32 + 2

            If InvSelectedItem = i Then
                Call engine.Device_Box_Textured_Render(19974, X, Y, 32, 32, s, 0, 0)
            End If

            engine.Device_Box_Textured_Render UserInventory(i).grhindex, X, Y, 32, 32, color(), 32, 32
            DibujarTexto UserInventory(i).Amount, Y, X, , D3DColorXRGB(255, 255, 255)

            If UserInventory(i).Equipped Then
                DibujarTexto "E", Y, X, , D3DColorXRGB(255, 0, 0)
            End If


        End If
    Next i

    D3DDevice.EndScene
    D3DDevice.Present InvRect, ByVal 0, frmMain.picInv.hWnd, ByVal 0

End Sub
 
Public Sub Initialize(ByRef InvPic As PictureBox, ByRef InvNpc As PictureBox, ByRef InvUser As PictureBox)
    '---------------------------------------------------------------------------------------
    ' Procedure : DrawInventoryNpc
    ' Author    : Midraks
    ' Date      : 11/04/2011
    '---------------------------------------------------------------------------------------
    Set InventoryWindow = InvPic
    Set InventoryWindowNpc = InvNpc
    Set InventoryWindowUser = InvUser
End Sub
 
Private Sub InventoryWindow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse move event of the inventory picture box
'Displays a ToolTip of the item under the mouse
'***************************************************
    Dim temp_x As Integer
    Dim temp_y As Integer
    Dim TempItem As Integer
   
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
   
    temp_x = X \ 32
    temp_y = Y \ 32
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindow.ScaleWidth \ 32) + 1
   
    If TempItem <= MAX_INVENTORY_SLOTS Then
        InventoryWindow.ToolTipText = UserInventory(TempItem).Name
    End If
End Sub
 
Private Sub InventoryWindow_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 27/07/04
'Implements the mouse up event of the inventory picture box
'Check outs which item was clicked
'***************************************************
    'Store previously selected item
    Dim prevSelItem As Long
   
    'Exit if it got outside the control's area
    If X < 0 Or Y < 0 Or X > InventoryWindow.Width Or Y > InventoryWindow.Height Then _
        Exit Sub
   
    prevSelItem = InvSelectedItem
 
    'Get the currently clickced item
    InvSelectedItem = ClickItem(CInt(X), CInt(Y))
End Sub

Sub DrawInventoryUser()
'---------------------------------------------------------------------------------------
' Procedure : DrawInventoryUser
' Author    : Midraks
' Date      : 11/04/2011
'---------------------------------------------------------------------------------------

    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer

    For i = 1 To UBound(UserInventory)
        If UserInventory(i).grhindex Then
            X = ((i - 1) Mod (InventoryWindowUser.Width / 32)) * 32 + 2
            Y = ((i - 1) \ (InventoryWindowUser.Width / 32)) * 32 + 2
            If InvSelectedItemUser = i Then
                Call engine.Draw_FilledBox(X, Y, 32, 32, D3DColorXRGB(0, 0, 0), D3DColorXRGB(255, 0, 0))
            End If
            Call engine.Draw_GrhIndex(UserInventory(i).grhindex, X, Y)
            Call engine.Text_Render_ext(UserInventory(i).Amount, Y, X, 40, 40, D3DColorXRGB(255, 255, 255))
        End If
    Next i

End Sub
 
Private Function ClickItemUser(ByVal X As Long, ByVal Y As Long) As Long
'---------------------------------------------------------------------------------------
' Procedure : ClickItemNpc
' Author    : Midraks
' Date      : 11/04/2011
'---------------------------------------------------------------------------------------
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    temp_x = X \ 32
    temp_y = Y \ 32
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindowUser.ScaleWidth \ 32) + 1
   
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        'Make sure slot isn't empty
        If UserInventory(TempItem).grhindex Then
            ClickItemUser = TempItem
        Else
            ClickItemUser = 0
        End If
    End If
End Function
 
Private Sub InventoryWindowUser_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------------------------
' Procedure : ClickItemNpc
' Author    : Midraks
' Date      : 11/04/2011
'---------------------------------------------------------------------------------------
    Dim prevSelItem As Long
   
    If X < 0 Or Y < 0 Or X > InventoryWindowUser.Width Or Y > InventoryWindowUser.Height Then _
    Exit Sub
   
    prevSelItem = InvSelectedItemUser
 
    InvSelectedItemUser = ClickItemUser(CInt(X), CInt(Y))
   
   
        With frmComerciar
        If InvSelectedItemUser <> 0 Then
        .Label1(0).Caption = UserInventory(InvSelectedItemUser).Name
        .Label1(1).Caption = UserInventory(InvSelectedItemUser).Valor
        .Label1(2).Caption = UserInventory(InvSelectedItemUser).Amount
        Select Case UserInventory(InvSelectedItemUser).OBJType
            Case 2
                .Label1(3).Caption = "Max Golpe:" & UserInventory(InvSelectedItemUser).MaxHit
                .Label1(4).Caption = "Min Golpe:" & UserInventory(InvSelectedItemUser).MinHit
                .Label1(3).Visible = True
                .Label1(4).Visible = True
            Case 3
                .Label1(3).Visible = False
                .Label1(4).Caption = "Defensa:" & UserInventory(InvSelectedItemUser).Def
                .Label1(4).Visible = True
        End Select
        End If
        End With
End Sub
 
Sub DrawInventoryNpc()
'---------------------------------------------------------------------------------------
' Procedure : DrawInventoryNpc
' Author    : Midraks
' Date      : 11/04/2011
'---------------------------------------------------------------------------------------
    Dim i As Byte
    Dim X As Integer
    Dim Y As Integer


    For i = 1 To UBound(NPCInventory)
        If NPCInventory(i).grhindex Then
            X = ((i - 1) Mod (InventoryWindowNpc.Width / 32)) * 32 + 2
            Y = ((i - 1) \ (InventoryWindowNpc.Width / 32)) * 32 + 2
            If InvSelectedItemNpc = i Then
                Call engine.Draw_FilledBox(X, Y, 32, 32, D3DColorXRGB(0, 0, 0), D3DColorXRGB(255, 0, 0))
            End If
            Call engine.Draw_GrhIndex(NPCInventory(i).grhindex, X, Y)
            Call engine.Text_Render_ext(NPCInventory(i).Amount, Y, X, 40, 40, D3DColorXRGB(255, 255, 255))
        End If
    Next i
    
End Sub
 
Private Function ClickItemNpc(ByVal X As Long, ByVal Y As Long) As Long
'---------------------------------------------------------------------------------------
' Procedure : ClickItemNpc
' Author    : Midraks
' Date      : 11/04/2011
'---------------------------------------------------------------------------------------
    Dim TempItem As Long
    Dim temp_x As Long
    Dim temp_y As Long
   
    temp_x = X \ 32
    temp_y = Y \ 32
   
    TempItem = temp_x + (temp_y + InventoryOffset) * (InventoryWindowNpc.ScaleWidth \ 32) + 1
   
    'Make sure it's within limits
    If TempItem <= MAX_INVENTORY_SLOTS Then
        'Make sure slot isn't empty
        If NPCInventory(TempItem).grhindex Then
            ClickItemNpc = TempItem
        Else
            ClickItemNpc = 0
        End If
    End If
End Function
 
Private Sub InventoryWindowNpc_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'---------------------------------------------------------------------------------------
' Procedure : ClickItemNpc
' Author    : Midraks
' Date      : 11/04/2011
'---------------------------------------------------------------------------------------
    Dim prevSelItem As Long
   
    If X < 0 Or Y < 0 Or X > InventoryWindowNpc.Width Or Y > InventoryWindowNpc.Height Then _
    Exit Sub
   
    prevSelItem = InvSelectedItemNpc
 
    InvSelectedItemNpc = ClickItemNpc(CInt(X), CInt(Y))
   
    With frmComerciar
    If InvSelectedItemNpc <> 0 Then
    .Label1(0).Caption = NPCInventory(InvSelectedItemNpc).Name
    .Label1(1).Caption = NPCInventory(InvSelectedItemNpc).Valor
    .Label1(2).Caption = NPCInventory(InvSelectedItemNpc).Amount
    Select Case NPCInventory(InvSelectedItemNpc).OBJType
        Case 2
            .Label1(3).Caption = "Max Golpe:" & NPCInventory(InvSelectedItemNpc).MaxHit
            .Label1(4).Caption = "Min Golpe:" & NPCInventory(InvSelectedItemNpc).MinHit
            .Label1(3).Visible = True
            .Label1(4).Visible = True
        Case 3
            .Label1(3).Visible = False
            .Label1(4).Caption = "Defensa:" & NPCInventory(InvSelectedItemNpc).Def
            .Label1(4).Visible = True
    End Select
    End If
    End With
   
End Sub

