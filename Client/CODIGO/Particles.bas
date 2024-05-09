Attribute VB_Name = "Particles"
Option Explicit

Public ParticleTexture(1 To 12) As Direct3DTexture8
Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180
Public Const RadianToDegree As Single = 57.2958279087977 '180 / Pi

Public LastTexture As Long
Public PixelOffsetX As Integer
Public PixelOffsetY As Integer

'Screen positioning
Public minY As Integer          'Start Y pos on current screen + tilebuffer
Public maxY As Integer          'End Y pos on current screen
Public minX As Integer          'Start X pos on current screen
Public maxX As Integer          'End X pos on current screen

Public ParticleOffsetX As Long
Public ParticleOffsetY As Long
Public LastOffsetX As Integer           'The last offset values stored, used to get the offset difference
Public LastOffsetY As Integer

Private Type Effect
    X As Single                 'Location of effect
    Y As Single
    GoToX As Single             'Location to move to
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    Used As Boolean             'If the effect is in use
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Direction As Integer        'Misc variable (depends on the effect)
    Particles() As Particle     'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    PartVertex() As TLVERTEX    'Used to point render particles
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindToChar As Integer       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
    BoundToMap As Byte          'If the effect is bound to the map or not (used only by the map editor)
End Type

Public NumEffects As Byte   'Maximum number of effects at once
Public Effect() As Effect   'List of all the active effects
 
'Constants With The Order Number For Each Effect
Public Const EffectNum_Fire As Byte = 1             'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow As Byte = 2             'Snow that covers the screen - weather effect
Public Const EffectNum_Heal As Byte = 3             'Healing effect that can bind to a character, ankhs float up and fade
Public Const EffectNum_Bless As Byte = 4            'Following three effects are same: create a circle around the central point
Public Const EffectNum_Protection As Byte = 5       ' (often the character) and makes the given particle on the perimeter
Public Const EffectNum_Strengthen As Byte = 6       ' which float up and fade out
Public Const EffectNum_Rain As Byte = 7             'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_EquationTemplate As Byte = 8 'Template for creating particle effects through equations - a page with some equations can be found here: http://www.vbgore.com/modules.php?name= ... opic&t=221" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
Public Const EffectNum_Waterfall As Byte = 9        'Waterfall effect
Public Const EffectNum_Summon As Byte = 10          'Summon effect
 
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal length As Long)
Public Declare Function timeGetTime Lib "winmm.dll" () As Long

Sub Engine_Init_ParticleEngine(Optional ByVal SkipToTextures As Boolean = False)
'*****************************************************************
'Loads all particles into memory - unlike normal textures, these stay in memory. This isn't
'done for any reason in particular, they just use so little memory since they are so small
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_ParticleEngine
'*****************************************************************
Dim i As Byte

    If Not SkipToTextures Then
    
        'Set the particles texture
        NumEffects = 25 'General_Var_Get(App.Path & "\Game.ini", "INIT", "NumEffects")
        ReDim Effect(1 To NumEffects)
    
    End If
    
    For i = 1 To UBound(ParticleTexture())
        If ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing
    Set ParticleTexture(i) = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "/Grh/" & "p" & i & ".png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)
Next i

    D3DDevice.SetRenderState D3DRS_POINTSIZE, engine.Engine_FToDW(2)
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    D3DDevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    D3DDevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0

End Sub

Function Effect_EquationTemplate_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: http://www.vbgore.com/Particle_effect_equations" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'More info: http://www.vbgore.com/CommonCode.Partic ... late_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_EquationTemplate_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_EquationTemplate  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_EquationTemplate_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_EquationTemplate_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... late_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim r As Single
   
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    r = (Index / 20) * exp(Index / Effect(EffectIndex).Progression Mod 3)
    X = r * Coseno(Index)
    Y = r * Seno(Index)
   
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_EquationTemplate_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... ate_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_EquationTemplate_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Function Effect_Bless_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... less_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Bless_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Bless     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Bless_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Bless_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Reset
'*****************************************************************
Dim A As Single
Dim X As Single
Dim Y As Single

    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Seno(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Coseno(A) * Effect(EffectIndex).Modifier)

    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)

End Sub
 
Private Sub Effect_Bless_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... ess_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Bless_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Function Effect_Fire_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Fire_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Fire_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Fire      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Fire_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Fire_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Seno((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Coseno((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
 
End Sub
 
Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... ire_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then
 
                    'Reset the particle
                    Effect_Fire_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Private Function Effect_FToDW(f As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_FToDW" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim Buf As D3DXBuffer
 
    'Converts a single into a long (Float to DWORD)
    Set Buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData Buf, 0, 4, 1, f
    D3DX.BufferGetData Buf, 0, 4, 1, Effect_FToDW
 
End Function
 
Function Effect_Heal_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Heal_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Heal_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Heal      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).X = X           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
   
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Heal_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Heal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Heal_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X - 10 + Rnd * 20, Effect(EffectIndex).Y - 10 + Rnd * 20, -Seno((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Coseno((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.8, 0.2, 0.2, 0.6 + (Rnd * 0.2), 0.01 + (Rnd * 0.5)
   
End Sub
 
Private Sub Effect_Heal_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... eal_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
Dim i As Integer
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
   
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then
 
                    'Reset the particle
                    Effect_Heal_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).color = 0
 
                End If
 
            Else
               
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
'*****************************************************************
'Kills (stops) a single effect or all effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Kill" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim loopc As Long
 
    'Check If To Kill All Effects
    If KillAll = True Then
 
        'Loop Through Every Effect
        For loopc = 1 To NumEffects
 
            'Stop The Effect
            Effect(loopc).Used = False
 
        Next
       
    Else
 
        'Stop The Selected Effect
        Effect(EffectIndex).Used = False
       
    End If
 
End Sub
 
Private Function Effect_NextOpenSlot() As Integer
'*****************************************************************
'Finds the next open effects index
'More info: http://www.vbgore.com/CommonCode.Partic ... xtOpenSlot" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
 
    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot
        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While Effect(EffectIndex).Used = True    'Check Next If Effect Is In Use
 
    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex
 
    'Clear the old information from the effect
    Erase Effect(EffectIndex).Particles()
    Erase Effect(EffectIndex).PartVertex()
    ZeroMemory Effect(EffectIndex), LenB(Effect(EffectIndex))
    Effect(EffectIndex).GoToX = -30000
    Effect(EffectIndex).GoToY = -30000
 
End Function
 
Function Effect_Protection_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... tion_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Protection_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Protection    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Protection_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Protection_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... tion_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim A As Single
Dim X As Single
Dim Y As Single
 
    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Seno(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Coseno(A) * Effect(EffectIndex).Modifier)
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_UpdateOffset(ByVal EffectIndex As Integer)
'***************************************************
'Update an effect's position if the screen has moved
'More info: http://www.vbgore.com/CommonCode.Partic ... dateOffset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'***************************************************
 
    Effect(EffectIndex).X = Effect(EffectIndex).X + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).Y = Effect(EffectIndex).Y + (LastOffsetY - ParticleOffsetY)
 
End Sub
 
Private Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)
 
'***************************************************
'Updates the binding of a particle effect to a target, if
'the effect is bound to a character
'More info: http://www.vbgore.com/CommonCode.Partic ... ateBinding" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'***************************************************
Dim TargetI As Integer
Dim TargetA As Single
 
    'Update position through character binding
    If Effect(EffectIndex).BindToChar > 0 Then
 
        'Store the character index
        TargetI = Effect(EffectIndex).BindToChar
 
        'Check for a valid binding index
        If TargetI > LastChar Then
            Effect(EffectIndex).BindToChar = 0
            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub
            End If
        ElseIf charlist(TargetI).Active = 0 Then
            Effect(EffectIndex).BindToChar = 0
            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub
            End If
        Else
 
            'Calculate the X and Y positions
            Effect(EffectIndex).GoToX = Engine_TPtoSPX(charlist(Effect(EffectIndex).BindToChar).Pos.X) + 16
            Effect(EffectIndex).GoToY = Engine_TPtoSPY(charlist(Effect(EffectIndex).BindToChar).Pos.Y)
 
        End If
 
    End If
 
    'Move to the new position if needed
    If Effect(EffectIndex).GoToX > -30000 Or Effect(EffectIndex).GoToY > -30000 Then
        If Effect(EffectIndex).GoToX <> Effect(EffectIndex).X Or Effect(EffectIndex).GoToY <> Effect(EffectIndex).Y Then
 
            'Calculate the angle
            TargetA = Engine_GetAngle(Effect(EffectIndex).X, Effect(EffectIndex).Y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
 
            'Update the position of the effect
            Effect(EffectIndex).X = Effect(EffectIndex).X - Seno(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
            Effect(EffectIndex).Y = Effect(EffectIndex).Y + Coseno(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
 
            'Check if the effect is close enough to the target to just stick it at the target
            If Effect(EffectIndex).GoToX > -30000 Then
                If Abs(Effect(EffectIndex).X - Effect(EffectIndex).GoToX) < 6 Then Effect(EffectIndex).X = Effect(EffectIndex).GoToX
            End If
            If Effect(EffectIndex).GoToY > -30000 Then
                If Abs(Effect(EffectIndex).Y - Effect(EffectIndex).GoToY) < 6 Then Effect(EffectIndex).Y = Effect(EffectIndex).GoToY
            End If
 
            'Check if the position of the effect is equal to that of the target
            If Effect(EffectIndex).X = Effect(EffectIndex).GoToX Then
                If Effect(EffectIndex).Y = Effect(EffectIndex).GoToY Then
 
                    'For some effects, if the position is reached, we want to end the effect
                    If Effect(EffectIndex).KillWhenAtTarget Then
                        Effect(EffectIndex).BindToChar = 0
                        Effect(EffectIndex).Progression = 0
                        Effect(EffectIndex).GoToX = Effect(EffectIndex).X
                        Effect(EffectIndex).GoToY = Effect(EffectIndex).Y
                    End If
                    Exit Sub    'The effect is at the right position, don't update
 
                End If
            End If
 
        End If
    End If
 
End Sub
 
Private Sub Effect_Protection_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... ion_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Protection_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Public Sub Effect_Render(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... ect_Render" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    'Check if we have the device
    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
 
    'Set the render state for the size of the particle
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize
   
    'Set the render state to point blitting
    If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
   
    'Set the last texture to a random number to force the engine to reload the texture
    LastTexture = -65489
 
    'Set the texture
    D3DDevice.SetTexture 0, ParticleTexture(Effect(EffectIndex).Gfx)
 
    'Draw all the particles at once
    D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), Len(Effect(EffectIndex).PartVertex(0))
 
    'Reset the render state back to normal
    If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 
End Sub
 
Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Snow_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Snow_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Snow      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Snow_Reset EffectIndex, loopc, 1
    Next loopc
 
    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Snow_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    If FirstReset = 1 Then
   
 
        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (PixelOffsetX + 400)), Rnd * (PixelOffsetY + 50), Rnd * 5, 5 + Rnd * 3, 0, 0
 
    Else
 
        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (PixelOffsetX + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (PixelOffsetY + 50)
        If Effect(EffectIndex).Particles(Index).sngX > PixelOffsetX Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (PixelOffsetY + 50)
        If Effect(EffectIndex).Particles(Index).sngY > PixelOffsetY Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (PixelOffsetX + 50)
 
    End If
 
    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.8, 0
 
End Sub
 
Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... now_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check if particle is in use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if to reset the particle
            If Effect(EffectIndex).Particles(loopc).sngX < -200 Then Effect(EffectIndex).Particles(loopc).sngA = 0
            If Effect(EffectIndex).Particles(loopc).sngX > (PixelOffsetX + 200) Then Effect(EffectIndex).Particles(loopc).sngA = 0
            If Effect(EffectIndex).Particles(loopc).sngY > (PixelOffsetY + 200) Then Effect(EffectIndex).Particles(loopc).sngA = 0
 
            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Reset the particle
                Effect_Snow_Reset EffectIndex, loopc
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Function Effect_Strengthen_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... then_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Strengthen_Begin = EffectIndex
 
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Strengthen    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Strengthen_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Strengthen_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... then_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim A As Single
Dim X As Single
Dim Y As Single
 
    'Get the positions
    A = Rnd * 360 * DegreeToRadian
    X = Effect(EffectIndex).X - (Seno(A) * Effect(EffectIndex).Modifier)
    Y = Effect(EffectIndex).Y + (Coseno(A) * Effect(EffectIndex).Modifier)
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt X, Y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_Strengthen_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... hen_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check if particle is in use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update the particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Strengthen_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Sub Effect_UpdateAll()
'*****************************************************************
'Updates all of the effects and renders them
'More info: http://www.vbgore.com/CommonCode.Partic ... _UpdateAll" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim loopc As Long
 
    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub
 
    'Set the render state for the particle effects
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
 
    'Update every effect in use
    For loopc = 1 To NumEffects
 
        'Make sure the effect is in use
        If Effect(loopc).Used Then
       
            'Update the effect position if the screen has moved
            Effect_UpdateOffset loopc
       
            'Update the effect position if it is binded
            Effect_UpdateBinding loopc
 
            'Find out which effect is selected, then update it
            If Effect(loopc).EffectNum = EffectNum_Fire Then Effect_Fire_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Snow Then Effect_Snow_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Heal Then Effect_Heal_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Bless Then Effect_Bless_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Protection Then Effect_Protection_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Strengthen Then Effect_Strengthen_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Rain Then Effect_Rain_Update loopc
            If Effect(loopc).EffectNum = EffectNum_EquationTemplate Then Effect_EquationTemplate_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update loopc
            If Effect(loopc).EffectNum = EffectNum_Summon Then Effect_Summon_Update loopc
           
            'Render the effect
            Effect_Render loopc, False
 
        End If
 
    Next
   
    'Set the render state back for normal rendering
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 
End Sub
 
Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Rain_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Rain_Begin = EffectIndex
 
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Rain      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Rain_Reset EffectIndex, loopc, 1
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... Rain_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    If FirstReset = 1 Then
 
        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (PixelOffsetX + 400)), Rnd * (PixelOffsetY + 50), Rnd * 5, 25 + Rnd * 12, 0, 0
 
    Else
 
        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (PixelOffsetY + 16)
        If Effect(EffectIndex).Particles(Index).sngX > PixelOffsetX Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (PixelOffsetY + 50)
        If Effect(EffectIndex).Particles(Index).sngY > PixelOffsetY Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (PixelOffsetX + 50)
 
    End If
 
    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.4, 0
 
End Sub
 
Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... ain_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check if the particle is in use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update the particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if to reset the particle
            If Effect(EffectIndex).Particles(loopc).sngX < -200 Then Effect(EffectIndex).Particles(loopc).sngA = 0
            If Effect(EffectIndex).Particles(loopc).sngX > (PixelOffsetX + 200) Then Effect(EffectIndex).Particles(loopc).sngA = 0
            If Effect(EffectIndex).Particles(loopc).sngY > (PixelOffsetY + 200) Then Effect(EffectIndex).Particles(loopc).sngA = 0
 
            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Reset the particle
                Effect_Rain_Reset EffectIndex, loopc
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Public Sub Effect_Begin(ByVal EffectIndex As Integer, ByVal X As Single, ByVal Y As Single, ByVal GfxIndex As Byte, ByVal Particles As Byte, Optional ByVal Direction As Single = 180, Optional ByVal BindToMap As Boolean = False)
'*****************************************************************
'A very simplistic form of initialization for particle effects
'Should only be used for starting map-based effects
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim RetNum As Byte
 
    Select Case EffectIndex
        Case EffectNum_Fire
            RetNum = Effect_Fire_Begin(X, Y, GfxIndex, Particles, Direction, 1)
        Case EffectNum_Waterfall
            RetNum = Effect_Waterfall_Begin(X, Y, GfxIndex, Particles)
    End Select
   
    'Bind the effect to the map if needed
    If BindToMap Then Effect(RetNum).BoundToMap = 1
   
End Sub
 
Function Effect_Waterfall_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... fall_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Waterfall_Begin = EffectIndex
 
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Waterfall     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).Used = True             'Enabled the effect
    Effect(EffectIndex).X = X                   'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Waterfall_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... fall_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
 
    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 130), 0, 8 + (Rnd * 6), 0, 0
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + (Rnd * 60), Effect(EffectIndex).Y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 0
    End If
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0
   
End Sub
 
Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... all_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
'Lorwik> Particulas de vbgore http://www.****.net" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go through the particle loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
   
        With Effect(EffectIndex).Particles(loopc)
   
            'Check if the particle is in use
            If .Used Then
   
                'Update The Particle
                .UpdateParticle ElapsedTime
 
                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).Y + 140) Or (.sngA = 0) Then
   
                    'Reset the particle
                    Effect_Waterfall_Reset EffectIndex, loopc
   
                Else
 
                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(loopc).X = .sngX
                    Effect(EffectIndex).PartVertex(loopc).Y = .sngY
   
                End If
   
            End If
           
        End With
 
    Next loopc
 
End Sub
 
Function Effect_Summon_Begin(ByVal X As Single, ByVal Y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 0) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... mmon_Begin" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim EffectIndex As Integer
Dim loopc As Long
'Lorwik> Particulas de vbgore http://www.****.net" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Summon_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Summon    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).Used = True                     'Enable the effect
    Effect(EffectIndex).X = X                           'Set the effect's X coordinate
    Effect(EffectIndex).Y = Y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For loopc = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(loopc) = New Particle
        Effect(EffectIndex).Particles(loopc).Used = True
        Effect(EffectIndex).PartVertex(loopc).rhw = 1
        Effect_Summon_Reset EffectIndex, loopc
    Next loopc
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Summon_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... mmon_Reset" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim X As Single
Dim Y As Single
Dim r As Single
   
    If Effect(EffectIndex).Progression > 1000 Then
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 1.4
    Else
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.5
    End If
    r = (Index / 30) * exp(Index / Effect(EffectIndex).Progression)
    X = r * Coseno(Index)
    Y = r * Seno(Index)
   
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).X + X, Effect(EffectIndex).Y + Y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0, Rnd, 0, 0.9, 0.2 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_Summon_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Partic ... mon_Update" class="postlink" rel="nofollow" onClick="window.open(this.href);return false;
'*****************************************************************
Dim ElapsedTime As Single
Dim loopc As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go Through The Particle Loop
    For loopc = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(loopc).Used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(loopc).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(loopc).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 1800 Then
 
                    'Reset the particle
                    Effect_Summon_Reset EffectIndex, loopc
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(loopc).Used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).Used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(loopc).color = 0
 
                End If
 
            Else
           
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(loopc).color = D3DColorMake(Effect(EffectIndex).Particles(loopc).sngR, Effect(EffectIndex).Particles(loopc).sngG, Effect(EffectIndex).Particles(loopc).sngB, Effect(EffectIndex).Particles(loopc).sngA)
                Effect(EffectIndex).PartVertex(loopc).X = Effect(EffectIndex).Particles(loopc).sngX
                Effect(EffectIndex).PartVertex(loopc).Y = Effect(EffectIndex).Particles(loopc).sngY
 
            End If
 
        End If
 
    Next loopc
 
End Sub
 
Public Function Engine_PixelPosX(ByVal X As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    Engine_PixelPosX = (X - 1) * 32
End Function

Public Function Engine_PixelPosY(ByVal Y As Integer) As Integer
'*****************************************************************
'Converts a tile position to a screen position
'*****************************************************************
    Engine_PixelPosY = (Y - 1) * 32
End Function
Public Function Engine_TPtoSPX(ByVal X As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPX
'************************************************************
    Engine_TPtoSPX = Engine_PixelPosX(X - (UserPos.X - AddtoUserPos.X) - engine.HalfWindowTileWidth - TileBufferSize) + engine.OffsetCounterX - 288 + TileBufferOffset

End Function

Public Function Engine_TPtoSPY(ByVal Y As Byte) As Long
'************************************************************
'Tile Position to Screen Position
'Takes the tile position and returns the pixel location on the screen
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_TPtoSPY
'************************************************************

    Engine_TPtoSPY = Engine_PixelPosY(Y - (UserPos.Y - AddtoUserPos.Y) - engine.HalfWindowTileWidth - TileBufferSize) + engine.OffsetCounterY - 288 + TileBufferOffset

End Function


Public Function Engine_GetAngle(ByVal CenterX As Integer, ByVal CenterY As Integer, ByVal TargetX As Integer, ByVal TargetY As Integer) As Single
'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle
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
