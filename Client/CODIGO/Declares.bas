Attribute VB_Name = "Mod_Declaraciones"
Option Explicit

Public HorasClima As Byte
Public MinutosClima As Byte
Public SegundosClima As Byte

Public aConnectAb As Byte
Public bConnectAb As Byte

Public StateName As String

Public mode As Boolean
Public temp_rgb(3) As Long
Public LuzGrh(3) As Long
Public AlphaY As Byte
 
'Desvanecimientos
Public Type tDesv
    Act As Byte
    AlphaX As Byte
End Type
Public TechoDesv As tDesv

Public InvSelectedItemNpc As Long
Public InvSelectedItemUser As Long
'////////////////////////////////////

Public Const CurServerIp As String = "127.0.0.1"
Public Const CurServerPort As Integer = 7666

'Sets a Grh animation to loop indefinitely.
Public Const INFINITE_LOOPS As Integer = -1

'Objetos p�blicos
Public DialogosClanes As New clsGuildDlg
Public Dialogos As New cDialogos
Public Audio As New clsAudio
Public Inventario As New clsGrapchicalInventory
Public SurfaceDB As clsSurfaceManDynDX8   'No va new porque es unainterfaz, el new se pone al decidir que clase de objeto es

'Sonidos
Public Const SND_CLICK As String = "click.Wav"
Public Const SND_PASOS1 As String = "23.Wav"
Public Const SND_PASOS2 As String = "24.Wav"
Public Const SND_NAVEGANDO As String = "50.wav"
Public Const SND_OVER As String = "click2.Wav"
Public Const SND_DICE As String = "cupdice.Wav"
Public Const SND_LLUVIAINEND As String = "lluviainend.wav"
Public Const SND_LLUVIAOUTEND As String = "lluviaoutend.wav"

'Musica
Public Const Mp3_Inicio As Byte = 1

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public ColoresPJ(0 To 50) As tColor

Public currentMidi As Long

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

Public NoRes As Boolean 'no cambiar la resolucion

Public RainBufferIndex As Long
Public FogataBufferIndex As Long

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87

Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

Public UsaMacro As Boolean
Public CnTd As Byte
Public SecuenciaMacroHechizos As Byte



'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]


Public Tips() As String * 255
Public Const LoopAdEternum = 999

'Direcciones
Public Enum E_Heading
    NORTH = 1
    EAST = 2
    SOUTH = 3
    WEST = 4
End Enum

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20
Public Const MAX_NPC_INVENTORY_SLOTS = 50
Public Const MAXHECHI = 35

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521

Public Enum Skills
     Suerte = 1
     Magia = 2
     Robar = 3
     Tacticas = 4
     Armas = 5
     Meditar = 6
     Apu�alar = 7
     Ocultarse = 8
     Supervivencia = 9
     Talar = 10
     Comerciar = 11
     Defensa = 12
     Pesca = 13
     Mineria = 14
     Carpinteria = 15
     Herreria = 16
     Liderazgo = 17 ' NOTA: Solia decir "Curacion"
     Domar = 18
     Proyectiles = 19
     Wresterling = 20
     Navegacion = 21
End Enum

Public Const FundirMetal As Integer = 88

'
' Mensajes
'
' MENSAJE_*  --> Mensajes de texto que se muestran en el cuadro de texto
'

Public Const MENSAJE_CRIATURA_FALLA_GOLPE As String = "La criatura fallo el golpe!!!"
Public Const MENSAJE_CRIATURA_MATADO As String = "La criatura te ha matado!!!"
Public Const MENSAJE_RECHAZO_ATAQUE_ESCUDO As String = "Has rechazado el ataque con el escudo!!!"
Public Const MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO  As String = "El usuario rechazo el ataque con su escudo!!!"
Public Const MENSAJE_FALLADO_GOLPE As String = "Has fallado el golpe!!!"
Public Const MENSAJE_SEGURO_ACTIVADO As String = ">>SEGURO ACTIVADO<<"
Public Const MENSAJE_SEGURO_DESACTIVADO As String = ">>SEGURO DESACTIVADO<<"
Public Const MENSAJE_PIERDE_NOBLEZA As String = "��Has perdido puntaje de nobleza y ganado puntaje de criminalidad!! Si sigues ayudando a criminales te convertir�s en uno de ellos y ser�s perseguido por las tropas de las ciudades."

Public Const MENSAJE_GOLPE_CABEZA As String = "��La criatura te ha pegado en la cabeza por "
Public Const MENSAJE_GOLPE_BRAZO_IZQ As String = "��La criatura te ha pegado el brazo izquierdo por "
Public Const MENSAJE_GOLPE_BRAZO_DER As String = "��La criatura te ha pegado el brazo derecho por "
Public Const MENSAJE_GOLPE_PIERNA_IZQ As String = "��La criatura te ha pegado la pierna izquierda por "
Public Const MENSAJE_GOLPE_PIERNA_DER As String = "��La criatura te ha pegado la pierna derecha por "
Public Const MENSAJE_GOLPE_TORSO  As String = "��La criatura te ha pegado en el torso por "

' MENSAJE_[12]: Aparecen antes y despues del valor de los mensajes anteriores (MENSAJE_GOLPE_*)
Public Const MENSAJE_1 As String = "��"
Public Const MENSAJE_2 As String = "!!"

Public Const MENSAJE_GOLPE_CRIATURA_1 As String = "��Le has pegado a la criatura por "

Public Const MENSAJE_ATAQUE_FALLO As String = " te ataco y fallo!!"

Public Const MENSAJE_RECIVE_IMPACTO_CABEZA As String = " te ha pegado en la cabeza por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ As String = " te ha pegado el brazo izquierdo por "
Public Const MENSAJE_RECIVE_IMPACTO_BRAZO_DER As String = " te ha pegado el brazo derecho por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ As String = " te ha pegado la pierna izquierda por "
Public Const MENSAJE_RECIVE_IMPACTO_PIERNA_DER As String = " te ha pegado la pierna derecha por "
Public Const MENSAJE_RECIVE_IMPACTO_TORSO As String = " te ha pegado en el torso por "

Public Const MENSAJE_PRODUCE_IMPACTO_1 As String = "��Le has pegado a "
Public Const MENSAJE_PRODUCE_IMPACTO_CABEZA As String = " en la cabeza por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ As String = " en el brazo izquierdo por "
Public Const MENSAJE_PRODUCE_IMPACTO_BRAZO_DER As String = " en el brazo derecho por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ As String = " en la pierna izquierda por "
Public Const MENSAJE_PRODUCE_IMPACTO_PIERNA_DER As String = " en la pierna derecha por "
Public Const MENSAJE_PRODUCE_IMPACTO_TORSO As String = " en el torso por "

Public Const MENSAJE_TRABAJO_MAGIA As String = "Haz click sobre el objetivo..."
Public Const MENSAJE_TRABAJO_PESCA As String = "Haz click sobre el sitio donde quieres pescar..."
Public Const MENSAJE_TRABAJO_ROBAR As String = "Haz click sobre la victima..."
Public Const MENSAJE_TRABAJO_TALAR As String = "Haz click sobre el �rbol..."
Public Const MENSAJE_TRABAJO_MINERIA As String = "Haz click sobre el yacimiento..."
Public Const MENSAJE_TRABAJO_FUNDIRMETAL As String = "Haz click sobre la fragua..."
Public Const MENSAJE_TRABAJO_PROYECTILES As String = "Haz click sobre la victima..."

Public Const MENSAJE_ENTRAR_PARTY_1 As String = "Si deseas entrar en una party con "
Public Const MENSAJE_ENTRAR_PARTY_2 As String = ", escribe /entrarparty"

Public Const MENSAJE_NENE As String = "Cantidad de NPCs: "

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    grhindex As Integer
    '[Alejo]: tipo de datos ahora es Long
    Amount As Long
    '[/Alejo]
    Equipped As Byte
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
End Type

Type NpCinV
    OBJIndex As Integer
    Name As String
    grhindex As Integer
    Amount As Integer
    Valor As Long
    OBJType As Integer
    Def As Integer
    MaxHit As Integer
    MinHit As Integer
    C1 As String
    C2 As String
    C3 As String
    C4 As String
    C5 As String
    C6 As String
    C7 As String
    
End Type

Type tReputacion 'Fama del usuario
    NobleRep As Long
    BurguesRep As Long
    PlebeRep As Long
    LadronesRep As Long
    BandidoRep As Long
    AsesinoRep As Long
    
    Promedio As Long
End Type

Type tEstadisticasUsu
    CiudadanosMatados As Long
    CriminalesMatados As Long
    UsuariosMatados As Long
    NpcsMatados As Long
    Clase As String
    PenaCarcel As Long
End Type

Public Nombres As Boolean

Public MixedKey As Long

'User status vars
Global OtroInventario(1 To MAX_INVENTORY_SLOTS) As Inventory

Public UserHechizos(1 To MAXHECHI) As Integer

Public NPCInventory(1 To MAX_NPC_INVENTORY_SLOTS) As NpCinV
Public NPCInvDim As Integer
Public UserMeditar As Boolean
Public UserName As String
Public UserPassword As String
Public UserMaxHP As Integer
Public UserMinHP As Integer
Public UserMaxMAN As Integer
Public UserMinMAN As Integer
Public UserMaxSTA As Integer
Public UserMinSTA As Integer
Public UserGLD As Long
Public UserLvl As Integer
Public UserPort As Integer
Public UserServerIP As String
Public UserCanAttack As Integer
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserEstadisticas As tEstadisticasUsu
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public IScombate As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As String

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
'<-------------------------NUEVO-------------------------->

Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String

Public Const NUMCIUDADES As Byte = 3
Public Const NUMSKILLS As Byte = 21
Public Const NUMATRIBUTOS As Byte = 5
Public Const NUMCLASES As Byte = 16
Public Const NUMRAZAS As Byte = 5

Public UserSkills(1 To NUMSKILLS) As Integer
Public SkillsNames(1 To NUMSKILLS) As String

Public UserAtributos(1 To NUMATRIBUTOS) As Integer
Public AtributosNames(1 To NUMATRIBUTOS) As String

Public Ciudades(1 To NUMCIUDADES) As String
Public CityDesc(1 To NUMCIUDADES) As String

Public ListaRazas(1 To NUMRAZAS) As String
Public ListaClases(1 To NUMCLASES) As String

Public Musica As Boolean
Public Sound As Boolean
Public Mp3Music As Boolean

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public logged As Boolean
Public NoPuedeUsar As Boolean

'Barrin 30/9/03
Public UserPuedeRefrescar As Boolean

Public UsingSkill As Integer

Public Enum E_MODO
    Normal = 1
    BorrarPj = 2
    CrearNuevoPj = 3
    Dados = 4
    RecuperarPass = 5
End Enum

Public EstadoLogin As E_MODO
   
Public Enum FxMeditar
'    FXMEDITARCHICO = 4
'    FXMEDITARMEDIANO = 5
'    FXMEDITARGRANDE = 6
'    FXMEDITARXGRANDE = 16
    CHICO = 4
    MEDIANO = 5
    GRANDE = 6
    XGRANDE = 16
End Enum


'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public stxtbuffercmsg As String 'Holds temp raw data from server
Public Connected As Boolean 'True when connected to server
Public UserMap As Integer

'String contants
Public Const ENDC As String * 1 = vbNullChar    'Endline character for talking with server
Public Const ENDL As String * 2 = vbCrLf        'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends

'
'********** FUNCIONES API ***********
'

Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal nVirtKey As Long) As Integer

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Lista de cabezas
Public Type tIndiceCabeza
    Head(1 To 4) As Integer
End Type

Public Type tIndiceCuerpo
    Body(1 To 4) As Integer
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Type tIndiceFx
    Animacion As Integer
    OffsetX As Integer
    OffsetY As Integer
End Type

'Old fashion BitBlt function
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long

'ECModule.dll
Public Declare Function Coseno Lib "ECModule.dll" (ByVal Number As Double) As Double
Public Declare Function Seno Lib "ECModule.dll" (ByVal Number As Double) As Double
Public Declare Function Tangente Lib "ECModule.dll" (ByVal Number As Double) As Double

Private Type ClientOptions 'Terminar
    UsaLifeBar As Boolean
End Type

Public ClientOptions As ClientOptions
