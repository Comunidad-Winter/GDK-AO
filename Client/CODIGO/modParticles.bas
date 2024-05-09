Attribute VB_Name = "modParticles"
Option Explicit

Rem El numero de particulas totales
 Public NParticulas_General As Integer

' Pi...
Public Const ConstPI As Single = 3.14159
     
' To convert to Radians
Public Const RAD As Single = ConstPI / 180
     
' To convert to Degrees
Public Const DEG As Single = 180 / ConstPI
 
' Public StartTime
Public lngStartTime As Long
     
Public myTexture As Direct3DTexture8
     
' TL Vertex
Public Const D3DFVF_TLVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1 Or D3DFVF_SPECULAR)
   
'  V E R T E X   B U F F E R
 
Private m_vertBuff As Direct3DVertexBuffer8  'we assume this has been created
Private m_vertCount As Long                 'we assume this has been set
 
Private addr As Long                        'will holds the address the D3D
                                                   'managed memory
'Private verts() As D3DTLVERTEX                'array that we want to point to
                                                   'D3D managed memory
 
Private Type structParticle
 
    sngA            As Single
    sngR            As Single
    sngG            As Single
    sngB            As Single
    sngAlphaDecay   As Single
   
    sngSize         As Single
   
    sngX            As Single
    sngY            As Single
   
    sngXAccel       As Single
    sngYAccel       As Single
   
    sngXSpeed       As Single
    sngYSpeed       As Single
   
End Type
 
 
'ONLY ON GROUP
 
Private Type structGroupParticle
 
    ParticleCounts  As Long
    Particles()     As structParticle
    vertsPoints()   As D3DTLVERTEX
   
    ' POSITION GROUP & INFO
    myTextureGrh        As Long ' Contiene el indice del grh del grafico
   
    sngX                As Single
    sngY                As Single
    sngProgression      As Single
 
    lngFloat0           As Long
    lngFloat1           As Long
    lngFloatSize        As Long
   
    lngPreviousFrame    As Long
 
   
End Type
 
Public ParticleGroup() As structGroupParticle
 
Public Sub loadGroupParticle()
 
    ReDim Preserve ParticleGroup(1 To 2) As structGroupParticle

    ParticleGroup(1).ParticleCounts = 100
        ReLocate 1, -50, -50
    Begin 1
    
        ParticleGroup(2).ParticleCounts = 100
                ReLocate 2, -50, -50
    Begin 2
 
End Sub
Public Sub Begin(grIndex As Integer)
    '//We initialize our stuff here
    Dim i As Long
     
    With ParticleGroup(grIndex)
   
        .lngFloat0 = engine.Engine_FToDW(0)
        .lngFloat1 = engine.Engine_FToDW(1)
        .lngFloatSize = engine.Engine_FToDW(20) '//Size of our flame particles..
         
        ' Redim our particles to the particlecount
        ReDim .Particles(0 To .ParticleCounts)
         
        ' Redim vertices to the particle count
        ' Point sprites, so 1 per particle
        ReDim .vertsPoints(0 To .ParticleCounts)
             
        Set m_vertBuff = D3DDevice.CreateVertexBuffer(Len(.vertsPoints(0)) * .ParticleCounts, 0, D3DFVF_TLVERTEX, D3DPOOL_MANAGED)
       
        m_vertBuff.Lock 0, Len(.vertsPoints(0)) * .ParticleCounts, addr, 0
       
        DXLockArray8 m_vertBuff, addr, .vertsPoints
       
        ' Now generate all particles
For i = 0 To .ParticleCounts - 1
            Reset grIndex, i, 1
        Next i
        DXUnlockArray8 m_vertBuff, .vertsPoints
       
        m_vertBuff.Unlock
       
        ' Set initial time
        .lngPreviousFrame = GetTickCount()
   
    End With
   
End Sub
 
Public Sub Reset(grIndex As Integer, i As Long, ParticleDunk As Byte) ' Reset GROUP
Dim X As Single, Y As Single, Radio As Single
Dim Progression As Integer, Direction As Single
 
    With ParticleGroup(grIndex)
             
       Select Case ParticleDunk
       
            Case 1 ' Blue Ball
               
                X = .sngX: Y = .sngY
                ResetIt grIndex, i, X, Y, -20, (-1 * Rnd), 0.01, Rnd, 16
                ResetColor grIndex, i, 0.25, 0.25, 1, 1, 0.1 + (0.1 * Rnd)
           
            Case 2 ' Fire
               
                X = .sngX + (Rnd * 10)
                Y = .sngY
                 
                ' This is were we will reset individual particles.
                ResetIt grIndex, i, X, Y, -0.4 + (Rnd * 0.8), -0.5 - (Rnd * 0.4), 0, -(Rnd * 0.3), 2
                ResetColor grIndex, i, 1, 0.5, 0.2, 0.6 + (0.2 * Rnd), 0.01 + Rnd * 0.05
           
            Case 3 ' Smoke
               
                X = .sngX + 1 * Rnd + .sngX
                Y = .sngY * Rnd + .sngY
                ResetIt grIndex, i, X, Y, -(Rnd / 3 + 0.1), ((Rnd / 2) - 0.7) * 3, (Rnd - 0.5) / 200, (Rnd - 0.5) / 200, 20
                ResetColor grIndex, i, 0.8, 0.8, 0.8, 0.3, (Rnd * 0.005) + 0.005
           
            Case 4 ' Snow
           
                X = .sngX * Rnd
                Y = .sngY * Rnd
 
                ResetIt grIndex, i, X, Y, Rnd - 0.5, (Rnd + 0.3) * 4, 0, 0, ((Rnd + 0.3) * 4) * 3
                ResetColor grIndex, i, 1, 1, 1, 0.5, 0.02 * Rnd
           
            Case 5 ' MagicFire
           
                X = .sngX + (Rnd * 2): Y = .sngY
                ResetIt grIndex, i, X, Y, -0.4 + (Rnd * 0.8), -0.5 - (Rnd * 0.4), 0, -(Rnd * 0.3), 32
                ResetColor grIndex, i, 1, 0.5, 0.1, 0.7 + (0.2 * Rnd), 0.01 + Rnd * 0.05
           
            Case 6 ' LevelUp
           
                X = .sngX: Y = .sngY
                ResetIt grIndex, i, X, Y, Rnd * 1.5 - 0.75, Rnd * 1.5 - 0.75, Rnd * 4 - 2, Rnd * -4 + 2, 16
                ResetColor grIndex, i, 1, 0.5, 0.1, 1, 0.07 + Rnd * 0.01
           
            Case 7 ' LevelUp2
           
                X = .sngX: Y = .sngY
                ResetIt grIndex, i, X + (Rnd * 32 - 16), Y + (Rnd * 64 - 32), Rnd * 1 - 0.5, Rnd * 1 - 0.5, Rnd - 0.5, Rnd * -0.9 + 0.45, 16
                ResetColor grIndex, i, 0.1 + (Rnd * 0.1), 0.1 + (Rnd * 0.1), 0.8 + (Rnd * 0.3), 1, 0.07 + Rnd * 0.01
           
            Case 8 ' Heal
           
                X = .sngX: Y = .sngY
                ResetIt grIndex, i, X, Y, Rnd * 1.4 - 0.7, Rnd * -0.4 - 1.5, Rnd - 0.5, Rnd * -0.2 + 0.1, 16
                ResetColor grIndex, i, 0.2, 0.3, 0.9, 0.4, 0.01 + Rnd * 0.01
           
            Case 9 ' WormHole
               
                Dim lo As Integer
                Dim la As Integer
                Dim VarB(3) As Single
                For lo = 0 To 3
                    VarB(lo) = Rnd * 5
                    la = Int(Rnd * 8)
                    If la * 0.5 <> Int(la * 0.5) Then VarB(lo) = -(VarB(lo))
                Next lo
           
                Progression = Int(Rnd * 10)
                Radio = (i * 0.0125) * Progression
                X = .sngX + (Radio * Cos((i)))
                Y = .sngY + (Radio * Sin((i)))
           
                ResetIt grIndex, i, X, Y, VarB(0), VarB(1), VarB(2), VarB(3), 32
                ResetColor grIndex, i, 1, 0.6, 0.3, 1, 0.02 + Rnd * 0.3
           
            Case 10 ' Twirl
           
 
                Progression = Progression + Direction
                If Progression > 50 Then Direction = -1
                If Progression < -50 Then Direction = 1
 
                Y = .sngY - 10 + Progression * Cos((i * 0.01) + Progression * 2)
                X = .sngX + 10 + Progression * Sin((i * 0.01) + Progression * 2)
               
                ResetIt grIndex, i, X, Y, 1, 1, 0, 0, 16
                ResetColor grIndex, i, 1, 0.25, 0.25, 1, 0.6 + Rnd * 0.3
           
            Case 11 ' Flower
           
                Radio = Cos(2 * (i * 0.1)) * 50
                X = .sngX + Radio * Cos(i * 0.1)
                Y = .sngY + Radio * Sin(i * 0.1)
                ResetIt grIndex, i, X, Y, 1, 1, 0, 0, 16
                ResetColor grIndex, i, 1, 0.25, 0.1, 1, 0.3 + (0.2 * Rnd) + Rnd * 0.3
           
            Case 12 ' Galaxy
                Radio = Sin(20 / (i + 1)) * 60
                X = .sngX + (Radio * Cos((i)))
                Y = .sngY + (Radio * Sin((i)))
                ResetIt grIndex, i, X, Y, 0, 0, 0, 0, 16
                ResetColor grIndex, i, 0.2, 0.2, 0.6 + 0.4 * Rnd, 1, 0 + Rnd * 0.3
           
            Case 13 ' Heart
           
                Y = .sngY - 50 * Cos(i * 0.01 * 2) * Sqr(Abs(Sin(i * 0.01)))
                X = .sngX + 50 * Sin(i * 0.01 * 2) * Sqr(Abs(Cos(i * 0.01)))
                ResetIt grIndex, i, X, Y, 0, 0, 0, -(Rnd * 0.2), 16
                ResetColor grIndex, i, 1, 0.5, 0.2, 0.6 + (0.2 * Rnd), 0.01 + Rnd * 0.08
           
            Case 14 ' BlueExplotion
           
                X = .sngX + Cos(i) * 30
                Y = .sngY + Sin(i) * 20
                ResetIt grIndex, i, X, Y, 0, -5 * (Rnd * 0.5), 0, -5 * (Rnd * 0.5), 8
                ResetColor grIndex, i, 0.3, 0.6, 1, 1, 0.05 + (Rnd * 0.1)
           
            Case 15 ' GP
           
                Radio = 50 + Rnd * 15 * Cos(i * 3.5)
                X = .sngX + (Radio * Cos((i * 0.01428571428)))
                Y = .sngY + (Radio * Sin((i * 0.01428571428)))
                ResetIt grIndex, i, X, Y, 0, 0, 0, 0, 16
                ResetColor grIndex, i, 0.2, 0.8, 0.4, 0.5, 0# + Rnd * 0.1
           
            Case 16 ' BTwirl
           
 
                Progression = (Progression + Direction) * Rnd
                If Progression > 50 Then Direction = -1
                If Progression < -50 Then Direction = 1
 
                Y = .sngY - 10 + Progression * Cos((i * 0.01) + Progression * 2)
                X = .sngX + 10 + Progression * Sin((i * 0.01) + Progression * 2)
                ResetIt grIndex, i, X, Y, 1, 1, 0, 0, 8
                ResetColor grIndex, i, 0.25, 0.25, 1, 1, 0.1 + Rnd * 0.3 + Rnd * 0.3
           
            Case 17 ' BT
 
                Progression = Progression + Direction
                If Progression > 50 Then Direction = -1
                If Progression < -50 Then Direction = 1
 
                Y = .sngY - 10 + Progression * Cos((i * 0.01) + Progression * 2)
                X = .sngX + 10 + Progression * Sin((i * 0.01) + Progression * 2)
                ResetIt grIndex, i, X, Y, -10, -1 * Rnd, 0, Rnd, 8
                ResetColor grIndex, i, 0.25, 0.25, 1, 1, 0.1 + (0.1 * Rnd) + Rnd * 0.3
           
            Case 18 ' Atomic
           
                Radio = 10 + Sin(2 * (i * 0.1)) * 50
                X = .sngX + Radio * Cos(i * 0.033333)
                Y = .sngY + Radio * Sin(i * 0.033333)
                ResetIt grIndex, i, X, Y, 1, 1, 0, 0, 8
                ResetColor grIndex, i, 0.4, 0.25, 1, 1, 0.3 + (0.2 * Rnd) + Rnd * 0.3
               
            Case 19 ' Medit
           
                X = .sngX + Cos(i * Rnd) * 45 * Sin(i * Rnd)
                Y = .sngY
                ResetIt grIndex, i, X, Y, 1, 1, 0, -10, 30
                ResetColor grIndex, i, 0, 0.5, 0.1, 1, 0.01 + (0.2 * Rnd)
    End Select
       
    End With
End Sub
 
Public Sub Update(grIndex As Integer, ByVal nparti As Byte)
    Dim i As Long
    Dim sngElapsedTime As Single
    
    With ParticleGroup(grIndex)
     
        '//We calculate the time difference here
        sngElapsedTime = (GetTickCount() - .lngPreviousFrame) / 100
        .lngPreviousFrame = GetTickCount()
         
        For i = 0 To .ParticleCounts
            With .Particles(i)
                UpdateParticle grIndex, i, sngElapsedTime
                 
                '//If the particle is invisible, reset it again.
If .sngA <= 0 Then
                    Reset grIndex, i, nparti
                End If
               
                ParticleGroup(grIndex).vertsPoints(i).rhw = 1
                ParticleGroup(grIndex).vertsPoints(i).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                ParticleGroup(grIndex).vertsPoints(i).sX = .sngX
                ParticleGroup(grIndex).vertsPoints(i).sY = .sngY
               
            End With
           
        Next i
       
        D3DVertexBuffer8SetData m_vertBuff, 0, Len(.vertsPoints(0)) * .ParticleCounts, 0, .vertsPoints(0)
 
    End With
End Sub
 
Public Sub RenderParticle(grIndex As Integer)
    With D3DDevice
        '//Set the render states for using point sprites
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1 'True
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'True
        .SetRenderState D3DRS_POINTSIZE, ParticleGroup(grIndex).lngFloatSize
        .SetRenderState D3DRS_POINTSIZE_MIN, ParticleGroup(grIndex).lngFloat0
        .SetRenderState D3DRS_POINTSCALE_A, ParticleGroup(grIndex).lngFloat0
        .SetRenderState D3DRS_POINTSCALE_B, ParticleGroup(grIndex).lngFloat0
        .SetRenderState D3DRS_POINTSCALE_C, ParticleGroup(grIndex).lngFloat1
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        .SetRenderState D3DRS_ALPHABLENDENABLE, 1
         
        '//Set up the vertex shader
        .SetVertexShader D3DFVF_TLVERTEX
         
        '//Set our texture
        .SetTexture 0, myTexture
       
        .SetStreamSource 0, m_vertBuff, Len(ParticleGroup(grIndex).vertsPoints(0))
         
        '//And draw all our particles :D
        .DrawPrimitiveUP D3DPT_POINTLIST, ParticleGroup(grIndex).ParticleCounts, _
            ParticleGroup(grIndex).vertsPoints(0), Len(ParticleGroup(grIndex).vertsPoints(0))
         
        '//Reset states back for normal rendering
        .SetRenderState D3DRS_ALPHABLENDENABLE, 0
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 0 'False
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0 'False
       
    End With
   
   '        m_vertBuff.Lock 0, Len(verts(0)) * m_vertCount, addr, 0
 
   '        DXLockArray8 m_vertBuff, addr, verts
   '
   '        Dim i As Long
 
   '        For i = 0 To m_vertCount - 1
   '            verts(i).X = i 'or what ever you want to dow with the data
   '        Next
 
   '        DXUnlockArray8 m_vertBuff, verts
 
   '        m_vertBuff.Unlock
   
End Sub
 
' FUNCTIONS FOR ONLY GROUPS
 
Public Sub ReLocate(grIndex As Integer, sngNewX As Single, sngNewY As Single) ' RELOCATE GROUP
    ParticleGroup(grIndex).sngX = sngNewX
    ParticleGroup(grIndex).sngY = sngNewY
End Sub
 
' FUNCTIONS FOR ONLY PARTICLE
 
Public Sub ResetColor(grIndex As Integer, Particle As Long, sngRed As Single, sngGreen As Single, sngBlue As Single, sngAlpha As Single, sngDecay As Single)
    ' Reset color to the new values
    With ParticleGroup(grIndex).Particles(Particle)
        .sngR = sngRed
        .sngG = sngGreen
        .sngB = sngBlue
        .sngA = sngAlpha
        .sngAlphaDecay = sngDecay
    End With
End Sub
 
Public Sub ResetIt(grIndex As Integer, Particle As Long, X As Single, Y As Single, XSpeed As Single, YSpeed As Single, XAcc As Single, YAcc As Single, sngResetSize As Single)
   
    With ParticleGroup(grIndex).Particles(Particle)
        .sngX = X
        .sngY = Y
        .sngXSpeed = XSpeed
        .sngYSpeed = YSpeed
        .sngXAccel = XAcc
        .sngYAccel = YAcc
        .sngSize = sngResetSize
    End With
End Sub
 
Public Sub UpdateParticle(grIndex As Integer, Particle As Long, sngTime As Single)
   
    With ParticleGroup(grIndex).Particles(Particle)
        .sngX = .sngX + .sngXSpeed * sngTime
        .sngY = .sngY + .sngYSpeed * sngTime
     
        .sngXSpeed = .sngXSpeed + .sngXAccel * sngTime
        .sngYSpeed = .sngYSpeed + .sngYAccel * sngTime
     
        .sngA = .sngA - .sngAlphaDecay * sngTime
    End With
End Sub
 
Public Sub UpdateParticleGroup(grIndex As Integer, sngNewX As Single, sngNewY As Single)
    Dim i As Long
    Dim sngElapsedTime As Single
   
    With ParticleGroup(grIndex)
     
        ' We calculate the time difference here
        sngElapsedTime = (GetTickCount() - .lngPreviousFrame) / 100
        .lngPreviousFrame = GetTickCount()
       
        .sngX = sngNewX
        .sngY = sngNewY
         
        For i = 0 To .ParticleCounts
            With .Particles(i)
                UpdateParticle grIndex, i, sngElapsedTime
                 
                ' If the particle is invisible, reset it again.
If .sngA <= 0 Then
                    Reset grIndex, i, 1
                End If
               
                ParticleGroup(grIndex).vertsPoints(i).rhw = 1
                ParticleGroup(grIndex).vertsPoints(i).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                ParticleGroup(grIndex).vertsPoints(i).sX = .sngX
                ParticleGroup(grIndex).vertsPoints(i).sY = .sngY
               
            End With
           
        Next i
       
    End With
   
End Sub
 Public Sub DibujarPartiMapa(ByVal map As Integer, ByVal NParticula As Integer, ByVal X As Integer, ByVal Y As Integer, Optional ByVal luz As Byte = 0)
'**************************************************************
'Author: Lautaro Mei (lautaro.mei@hotmail.com)
'Last Modify Date: 30/03/2011
'Dibuja las particulas en el mapa
'**************************************************************
If UserMap = map Then
    MapData(X, Y).PartiIndex = NParticula
    
If luz = 1 Then
    engine.Light_Create X, Y
End If

Else
MapData(X, Y).PartiIndex = 0
End If
 End Sub
 
 Public Sub CargarPartis(ByVal NParticula As Integer, ByVal TipoParticula As Integer)
 '**************************************************************
'Author: Lautaro Mei (lautaro.mei@hotmail.com)
'Last Modify Date: 30/03/2011
'Actualiza las particulas
'**************************************************************
     Update NParticula, TipoParticula
     RenderParticle NParticula
 End Sub

Public Sub Parti_Render_Ir()
'**************************************************************
'Author: Lautaro Mei (lautaro.mei@hotmail.com)
'Last Modify Date: 30/03/2011
'Algoritmo que calcula y renderiza el movimiento de las particulas
'**************************************************************
If Parti_Ir = 1 Then

MapData(Parti_Cuenta, Parti_Cuenta_Y).PartiIndex = 0


If Parti_Xd > Parti_Cuenta Then

Parti_Cuenta = Parti_Cuenta + Sin(angle) * Parti_Vel * engine.timerElapsedTime
If Parti_Cuenta > Parti_Xd Then
Parti_Cuenta = Parti_Xd
Parti_Ir = 0
End If
Else
Parti_Cuenta = Parti_Cuenta + Sin(angle) * Parti_Vel * engine.timerElapsedTime
If Parti_Cuenta < Parti_Xd Then
Parti_Cuenta = Parti_Xd
Parti_Ir = 0
End If
End If

If Parti_Yd > Parti_Cuenta_Y Then
Parti_Cuenta_Y = Parti_Cuenta_Y - Cos(angle) * Parti_Vel * engine.timerElapsedTime

If Parti_Cuenta_Y > Parti_Yd Then
Parti_Cuenta_Y = Parti_Yd
Parti_Ir = 0
End If

Else
Parti_Cuenta_Y = Parti_Cuenta_Y - Cos(angle) * Parti_Vel * engine.timerElapsedTime

If Parti_Cuenta_Y < Parti_Yd Then
Parti_Cuenta_Y = Parti_Yd
Parti_Ir = 0
End If
End If

MapData(Parti_Cuenta, Parti_Cuenta_Y).PartiIndex = 2
Parti_X = Parti_Xd
Parti_Y = Parti_Yd
End If

End Sub
Public Sub Parti_Viajar(ByVal X As Single, ByVal Y As Single, ByVal Xd As Integer, ByVal Yd As Integer)
'**************************************************************
'Author: Lautaro Mei (lautaro.mei@hotmail.com)
'Last Modify Date: 30/03/2011
'Setea todas las variables y pone On el renderizado
'**************************************************************
MapData(Parti_X, Parti_Y).PartiIndex = 0
Parti_Cuenta = X
Parti_Cuenta_Y = Y
Parti_Xd = Xd
Parti_Yd = Yd
Parti_Ir = 1
angle = DegreeToRadian * Engine_GetAngle(X, Y, Xd, Yd)
End Sub

Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEn ... e_GetAngle
'************************************************************
Dim SideA As Single
Dim SideC As Single
 
    On Error GoTo ErrOut
 
    'Check for horizontal lines (90 or 270 degrees)
    If CenterY = TargetY Then
 
        'Check for going right (90 degrees)
        If CenterX < TargetX Then
            Engine_GetAngle = 90
 
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Check for horizontal lines (360 or 180 degrees)
    If CenterX = TargetX Then
 
        'Check for going up (360 degrees)
        If CenterY > TargetY Then
            Engine_GetAngle = 360
 
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If
 
        'Exit the function
        Exit Function
 
    End If
 
    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)
 
    'Side B = CenterY
 
    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)
 
    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)
    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583
 
    'If the angle is >180, subtract from 360
    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle
 
    'Exit function
 
Exit Function
 
    'Check for error
ErrOut:
 
    'Return a 0 saying there was an error
    Engine_GetAngle = 0
 
Exit Function
 
End Function

