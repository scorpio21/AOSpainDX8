Attribute VB_Name = "Mod_Declaraciones"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Marquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 numero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Codigo Postal 1900
'Pablo Ignacio Marquez


Option Explicit

Public RawServersList As String

Public Type tServerInfo
    Ip As String
    Puerto As Integer
    Desc As String
    PassRecPort As Integer
End Type

Public ServersLst() As tServerInfo
Public ServersRecibidos As Boolean

Public CurServer As Integer

Public CreandoClan As Boolean
Public ClanName As String
Public Site As String

Public UserCiego As Boolean
Public UserEstupido As Boolean

'[Alejo-21-5]
Public Type tConfEnviada
    ModoPaquetes As Long
    IntOk As Long 'tiempo de espera entre M
    IntOkCantPak As Long 'Cantidad de paketes a enviar para esperar un ok
End Type

Public ConfEnv As tConfEnviada

Public Type tFlagOk
    NumMagico As Long
    cant As Long
End Type

Public FlagOk As tFlagOk
Public PingInicio As Long, Ping As Long

'Timers de GetTickCount
Public Const tAt = 2000
Public Const tUs = 600

Public Const bCabeza = 1
Public Const bPiernaIzquierda = 2
Public Const bPiernaDerecha = 3
Public Const bBrazoDerecho = 4
Public Const bBrazoIzquierdo = 5
Public Const bTorso = 6

Public Const PrimerBodyBarco = 84
Public Const UltimoBodyBarco = 87


Public Dialogos As New cDialogos
Public Sound As clsSoundEngine  ' Motor de audio DX8 - inicializado en General.bas
Public NumEscudosAnims As Integer

Public ArmasHerrero(0 To 100) As Integer
Public ArmadurasHerrero(0 To 100) As Integer
Public ObjCarpintero(0 To 100) As Integer

'[KEVIN]
Public Const MAX_BANCOINVENTORY_SLOTS = 40
Public UserBancoInventory(1 To MAX_BANCOINVENTORY_SLOTS) As Inventory
'[/KEVIN]


Public Tips() As String * 255
Public Const LoopAdEternum = 999

Public Const NUMCIUDADES = 3

'Direcciones
'Public Const NORTH = 1
'Public Const EAST = 2
'Public Const SOUTH = 3
'Public Const WEST = 4

'Objetos
Public Const MAX_INVENTORY_OBJS = 10000
Public Const MAX_INVENTORY_SLOTS = 20
Public Const MAX_NPC_INVENTORY_SLOTS = 50
Public Const MAXHECHI = 35

Public Const NUMSKILLS = 22 '21+1 por la resistencia magica [Efestos]
Public Const NUMATRIBUTOS = 5
Public Const NUMCLASES = 20 '16+4 gladiador,arquero,chaman y aldeano [Neptuno]
Public Const NUMRAZAS = 7 '5+2 orcos y hobbits [Neptuno]

Public Const MAXSKILLPOINTS = 100

Public Const FLAGORO = 777

Public Const FOgata = 1521


'-- Skills: reemplazados por Enum eSkill (ver abajo) para compatibilidad DX8
'Public Const Suerte = 1
'Public Const Magia = 2
'Public Const Robar = 3
'Public Const Tacticas = 4
'Public Const Armas = 5
'Public Const Meditar = 6
'Public Const Apuñalar = 7
'Public Const Ocultarse = 8
'Public Const Supervivencia = 9
'Public Const Talar = 10
'Public Const Comerciar = 11
'Public Const Defensa = 12
'Public Const Pesca = 13
'Public Const Mineria = 14
'Public Const Carpinteria = 15
'Public Const Herreria = 16
'Public Const Curacion = 17
'Public Const Domar = 18
'Public Const Proyectiles = 19
'Public Const Wresterling = 20
'Public Const Navegacion = 21
'Public Const FundirMetal = 88
'Public Const Magia = 2  '-- DX8: en eSkill enum
'Public Const Robar = 3  '-- DX8: en eSkill enum
'Public Const Tacticas = 4  '-- DX8: en eSkill enum
'Public Const Armas = 5  '-- DX8: en eSkill enum
'Public Const Meditar = 6  '-- DX8: en eSkill enum
'Public Const Apuñalar = 7  '-- DX8: en eSkill enum
'Public Const Ocultarse = 8  '-- DX8: en eSkill enum
'Public Const Supervivencia = 9  '-- DX8: en eSkill enum
'Public Const Talar = 10  '-- DX8: en eSkill enum
'Public Const Comerciar = 11  '-- DX8: en eSkill enum
'Public Const Defensa = 12  '-- DX8: en eSkill enum
'Public Const Pesca = 13  '-- DX8: en eSkill enum
'Public Const Mineria = 14  '-- DX8: en eSkill enum
'Public Const Carpinteria = 15  '-- DX8: en eSkill enum
'Public Const Herreria = 16  '-- DX8: en eSkill enum
'Public Const Curacion = 17  '-- DX8: en eSkill enum
'Public Const Domar = 18  '-- DX8: en eSkill enum
'Public Const Proyectiles = 19  '-- DX8: en eSkill enum
'Public Const Wresterling = 20  '-- DX8: en eSkill enum
'Public Const Navegacion = 21  '-- DX8: en eSkill enum

'Public Const FundirMetal = 88  '-- DX8: en eSkill enum

'Inventario
Type Inventory
    OBJIndex As Integer
    Name As String
    GrhIndex As Integer
    '[Alejo]: tipo de datos ahora es Long
    amount As Long
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
    GrhIndex As Integer
    amount As Integer
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

Public ListaRazas() As String
Public ListaClases() As String

Public Nombres As Boolean

Public MixedKey As Long

'User status vars
Public UserInventory(1 To MAX_INVENTORY_SLOTS) As Inventory
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
Public UserPuedeRefrescar As Boolean
Public UserEstado As Byte '0 = Vivo & 1 = Muerto
Public UserPasarNivel As Long
Public UserExp As Long
Public UserReputacion As tReputacion
Public UserDescansar As Boolean
Public tipf As String
Public PrimeraVez As Boolean
Public FPSFLAG As Boolean
Public pausa As Boolean
Public IScombate As Boolean
Public ISseguro As Boolean
Public UserParalizado As Boolean
Public UserNavegando As Boolean
Public UserHogar As String

'<-------------------------NUEVO-------------------------->
Public Comerciando As Boolean
Public MD5HushYo As String * 16
Public MostrarIndexNombre As Integer
Public uPos As WorldPos
'<-------------------------NUEVO-------------------------->

Public UserClase As String
Public UserSexo As String
Public UserRaza As String
Public UserEmail As String

Public UserSkills() As Integer
Public SkillsNames() As String

Public UserAtributos() As Integer
Public AtributosNames() As String

Public Ciudades() As String
Public CityDesc() As String

Public Musica As Byte
Public Fx As Byte

Public SkillPoints As Integer
Public Alocados As Integer
Public flags() As Integer
Public Oscuridad As Integer
Public logged As Boolean
Public NoPuedeUsar As Boolean

Public UsingSkill As Integer


'Server stuff
Public RequestPosTimer As Integer 'Used in main loop
Public stxtbuffer As String 'Holds temp raw data from server
Public SendNewChar As Boolean 'Used during login
Public Connected As Boolean 'True when connected to server
Public DownloadingMap As Boolean 'Currently downloading a map from server
Public UserMap As Integer

'String contants
Public ENDC As String 'Endline character for talking with server
Public ENDL As String 'Holds the Endline character for textboxes

'Control
Public prgRun As Boolean 'When true the program ends
Public finpres As Boolean

Public IPdelServidor As String
Public PuertoDelServidor As String

'********** FUNCIONES API ***********
Public Declare Function GetTickCount Lib "kernel32" () As Long

'para escribir y leer variables
Public Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Public Declare Function getprivateprofilestring Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'Teclado
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
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
    offsetX As Integer
    offsetY As Integer
End Type

'�Carga los txt de la web?
Public DescargarTxt(4) As Boolean

'Posicion en un mapa (movido desde TileEngine.bas para romper dependencia circular)
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type
'=== Enums DX8 TileEngine (agregados desde Aodrag9) ===
'Direcciones
Public Enum E_Heading
    SOUTH = 3
    NORTH = 1
    WEST = 4
    EAST = 2
End Enum

Public Enum TipoPaso
    CONST_BOSQUE = 1
    CONST_NIEVE = 2
    CONST_CABALLO = 3
    CONST_DUNGEON = 4
    CONST_PISO = 5
    CONST_DESIERTO = 6
    CONST_PESADO = 7
End Enum

Public Type tPaso
    CantPasos As Byte
    Wav() As Integer
End Type

Public Const NUM_PASOS As Byte = 7

Public Enum eMoveType
    Inventory = 1
    Bank = 2
    SpellsI = 3
End Enum

'=== Tipos y constantes DX8/Aodrag9 - requeridos por TileEngine.bas y clsSoundEngine ===

Public Enum E_SISTEMA_MUSICA
    CONST_DESHABILITADA = 0
    CONST_MP3 = 1
    CONST_MIDI = 2
End Enum

Private Type tOption
    NoRes As Byte
    BaseTecho As Byte
    bCursores As Byte
    MovEscritura As Byte
    URLCON As Byte
    NamePlayers As Byte
    PrimeraVez As Byte
    GuildNews As Byte
    VSynC As Byte
    VProcessing As Byte
    MusicVolume As Long
    HechizosClasicos As Byte
    Ambient As Byte
    AmbientVol As Long
    Audio As Byte
    FxNavega As Long
    InvertirSonido As Byte
    FXVolume As Long
    sMusica As E_SISTEMA_MUSICA
    BloqCruceta As Byte
End Type

Public Opciones As tOption

Public Const GRH_FOGATA As Integer = 1521

Public Const MUS_Inicio As String = "6"
Public Const MUS_CrearPersonaje As String = "7"
Public Const MUS_VolverInicio As String = "53"

' Constantes de sonido numericas para clsSoundEngine / TileEngine (DX8)
Public Const SND_CLICK As Byte = 190
Public Const SND_NAVEGANDO As Byte = 50
Public Const SND_OVER As Byte = 0
Public Const SND_DICE As Byte = 188
Public Const SND_FUEGO As Byte = 79
Public Const SND_PASOS1 As Byte = 23
Public Const SND_PASOS2 As Byte = 24

'=== Types de color para TileEngine.bas (DX8) ===
Public Type RGB
    r As Long
    g As Long
    b As Long
End Type

Public Type RGBClimax
    r As Byte
    g As Byte
    b As Byte
    A As Byte
End Type

'=== Tipos/Enums de Aodrag9 Declares.bas agregados para completar DX8 ===
Public Type Servidores
    Nombre As String
    Ip As String
    Puerto As Integer
End Type

Private Type tMapaConnect
    Map As Byte
    X As Byte
    Y As Byte
End Type

Public Type tColor
    r As Byte
    g As Byte
    b As Byte
End Type

Public Enum eClass
    Mago = 1
    Clerigo = 2
    Guerrero = 3
    Asesino = 4
    Ladron = 5
    Bardo = 6
    Druida = 7
    Bandido = 8
    Paladin = 9
    Cazador = 10
    Pescador = 11
    Herrero = 12
    Leñador = 13
    Minero = 14
    Carpintero = 15
    Pirata = 16
    Gladiador = 17
    Arquero = 18
    Chaman = 19
    Aldeano = 21
End Enum

Public Enum eAtributos
    Fuerza = 1
    Agilidad = 2
    Inteligencia = 3
    Energia = 4
    Constitucion = 5
End Enum

Public Enum PlayerType
    User = &H1
    Consejero = &H2
    SemiDios = &H4
    Dios = &H8
    Admin = &H10
    RoleMaster = &H20
    ChaosCouncil = &H40
    RoyalCouncil = &H80
End Enum

Public Enum eObjType
    otUseOnce = 1
    otWeapon = 2
    otArmadura = 3
    otArboles = 4
    otGuita = 5
    otPuertas = 6
    otContenedores = 7
    otCarteles = 8
    otLlaves = 9
    otForos = 10
    otPociones = 11
    otBebidas = 13
    otLe�a = 14
    otFogata = 15
    otescudo = 16
    otcasco = 17
    otAnillo = 18
    otTeleport = 19
    otYacimiento = 22
    otMinerales = 23
    otPergaminos = 24
    otInstrumentos = 26
    otYunque = 27
    otFragua = 28
    otBarcos = 31
    otFlechas = 32
    otBotellaVacia = 33
    otBotellaLlena = 34
    otManchas = 35          'No se usa
    otCualquiera = 1000
End Enum

Public Enum E_MODO
    Normal = 1
    CrearNuevoPj = 2
    Dados = 3
    LoginCuenta = 4
    BorrandoPJ = 5
End Enum

Public Type pjs
    NamePJ As String
    LvlPJ As Integer
    ClasePJ As eClass
    
    Acuerpo As Integer
    rcvHead As Integer
    rcvCasco As Integer
    rcvShield As Integer
    rcvWeapon As Integer
    rcvRaza As Integer
    PJLogged As Byte
    mapa As String
End Type

Public Type acc
    Name As String
    Pass As String
    Email As String
    preg As String
    resp As String
   
    CantPJ As Byte
    pjs(1 To 8) As pjs
End Type

Public Enum FxMeditar
    CHICO = 46
    MEDIANO = 3
    GRANDE = 49
    XGRANDE = 40
    XXGRANDECIU = 34
    XXGRANDECRI = 35
End Enum

Public Enum eClanType
    ct_RoyalArmy
    ct_Evil
    ct_Neutral
    ct_GM
    ct_Legal
    ct_Criminal
End Enum

Public Enum eEditOptions
    eo_Gold = 1
    eo_Experience
    eo_Body
    eo_Head
    eo_CiticensKilled
    eo_CriminalsKilled
    eo_Level
    eo_Class
    eo_Skills
    eo_SkillPointsLeft
    eo_Nobleza
    eo_Asesino
    eo_Sex
    eo_Raza
    eo_addGold
    eo_Speed
End Enum

Public Enum eTrigger
    Nada = 0
    BAJOTECHO = 1
    trigger_2 = 2
    POSINVALIDA = 3
    ZONASEGURA = 4
    ANTIPIQUETE = 5
    ZONAPELEA = 6
End Enum

Public Type tHead
    Texture As Integer
    startX As Integer
    startY As Integer
End Type

Public Type tIndiceAtaque
    Body(1 To 4) As Long
    HeadOffsetX As Integer
    HeadOffsetY As Integer
End Type

Public Enum eCursorState
    cur_Normal = 0
    cur_Action
    cur_Wait
    cur_Npc
    cur_Npc_Hostile
    cur_User
    cur_User_Danger
    cur_Obj
End Enum

Public Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    id As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    Angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    
    Speed As Single
    life_counter As Long

    Radio As Integer
End Type


Public Enum eSkill
    Equitacion = 1
    Magia = 2
    Robar = 3
    Tacticas = 4
    Armas = 5
    Meditar = 6
    Apuñalar = 7
    Ocultarse = 8
    Supervivencia = 9
    Talar = 10
    Comerciar = 11
    Defensa = 12
    Pesca = 13
    Mineria = 14
    Carpinteria = 15
    Herreria = 16
    Liderazgo = 17
    Domar = 18
    Proyectiles = 19
    Wresterling = 20
    Navegacion = 21
    Resistencia = 22  '[Efestos]
    FundirMetal = 88
End Enum

'=== Variables globales de motor DX8 (agregadas para compatibilidad TileEngine) ===
Public UserCharIndex As Integer  ' Indice del personaje del usuario en CharList()

