# MEMORIA — Cliente Argentum Online DX8 (aoscorpio2) `I:\AospainOri`

## IMPORTANTE (no olvidar)
- **aospain NO usa compresión todavía.** NO implantar compresión por ahora. Existe material de referencia (modCompression, `GameIni tCabecera`, compresor AODrag) pero queda en reserva para futuro.
- Objetivo actual: **compilar limpio** `vb6 /make` y obtener `I:\AospainOri\cli\aoscorpio2.exe`.

## Compilación (comando)
- Borrar el log ANTES de cada run: `C:\Users\sonsc\AppData\Local\Temp\opencode\vb6err.txt`.
- `& "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make I:\AospainOri\cli\Client.vbp /out C:\Users\sonsc\AppData\Local\Temp\opencode\vb6err.txt`
- Los errores de VB6 pueden reportar línea con **desfase (+2 aprox)** respecto a la línea real.
- VB6 reporta el PRIMER error del módulo y para → errores encadenados, corregir uno a uno.

## ESTADO ACTUAL (corte de hoy) ✅ COMPILACIÓN COMPLETA — B2 en curso (Task 1 hecha)
- **`vb6 /make` EXITOSO (ExitCode 0)**: log = "La generación de 'aoscorpio2.exe' ha tenido éxito." y **`I:\AospainOri\cli\aoscorpio2.exe` regenerado** (tras `5437b74`, 08/08/2026).
- B1 cerrado (migración DX7→DX8). B2 (sistema de cuentas): Task 1 (protocolo cliente) COMPLETA y revisada; quedan Task 2 (conexión/formularios) y Task 3 (server /SALIR) — ver sección PRÓXIMO PASO abajo.

## CORRECCIONES HECHAS (sesión de hoy y previas)
1. `TCP.bas:397` `Case "CC"` → `Call MakeChar(..., 0)` (10º arg `Ataque` = 0, TileEngine.bas:613).
2. `TCP.bas:420-434` `Case "CP"` → `.Head = Val(...)`, `.Casco = tempint`, `.Body = BodyData(...)` (charlist.Head/.Casco son **Integer/IDs** en el activo, no HeadData).
3. TileEngine.bas: añadidas `UserMaxAGU/UserMinAGU/UserMaxHAM/UserMinHAM` (usadas en TCP.bas:766 "EHYS"; antes solo existían en OLD).
4. TileEngine.bas:677,847 → `Dialogos.RemoveDialog` → `Dialogos.QuitarDialogo` (método DX8 de `cDialogos.cls:270`).
5. TileEngine.bas:849,2493 → `MinLimiteX/Y/MaxLimiteX/Y` (nunca definidas) → `XMinMapSize/XMaxMapSize/YMinMapSize/YMaxMapSize` (=1/200).
6. Declares.bas (activo) → añadidas `Public Const CASPER_HEAD As Integer = 500` y `FRAGATA_FANTASMAL As Integer = 87` (valores de `ee\Cliente\CODIGO\Declares.bas:135-136`; usadas en TileEngine.bas:964).
7. General.bas:1360 → `Function HayAgua(ByVal X As Integer, ByVal Y As Integer)` (antes ByRef; se llamaba con `UserPos.X` que es Long — Position en Declares.bas:364).
8. TileEngine.bas `ActualizarMiniMapa` → eliminadas las líneas `frmMain.UserM.Left/Top` (**UserM no existe** en frmMain DX8); se conserva el cálculo de `MinimapMaxX/MaxY` (usado por DibujarMiniMapa:1102+).
9. General.bas (activo) → añadida `Public Function ARGB(r,g,b,A)` copiada de `ee\Cliente\CODIGO\General.bas:1127-1146` (necesaria en TileEngine.bas:1211; los colores se usan como D3D AARRGGBB en Geometry_Create_Box).
10. TileEngine.bas → añadido `battle_mode As Boolean` a `Type MapInfo` (200) y `MapDat.battle_mode` → `MapInfo.battle_mode` (linea ~1310). `MapDat` nunca existió (bug de referencia). battle_mode queda siempre False (no hay fuente de datos).
11. TileEngine.bas → añadidos tipos y arrays vacíos para features muertas de Irongete (nunca declarados ni en preHeading):
    - `Private Type tZona` (x1,x2,y1,y2,Grh) + `Private ZonaList(-1 To -1) As tZona` (usado ~1445).
    - `Private Type tEfecto` (EfectoIndex, beneficioso, Grh) + `Private EfectoList(-1 To -1) As tEfecto` (usado ~1577).
    - El bucle `For ... To UBound(...)` con array fijo vacío (-1) no itera → sin error de runtime.
12. TileEngine.bas ~267 → `Public ColoresPJ(0 To 50) As Long` + `Public Sub InitColoresPJ()` (default vbWhite; 49 = RGB(0,255,255); 50 = RGB(255,0,0)) + `Call InitColoresPJ` al inicio de InitTileEngine (~2126). No existe `Colores.ini` en el activo ni en referencias ee/dx.
13. TileEngine.bas ~262 → `Public Const OFFSET_HEAD As Integer = 34` (usado en línea 1799; no estaba definida).
14. cDialogos.cls:197 → `Update_Dialog_Pos(ByVal X As Integer, ByVal Y As Integer, ByVal Index As Integer)` (antes ByRef; el llamador pasa expresiones Long). También TileEngine.bas:1799 `UpdateDialogPos` → `Update_Dialog_Pos` (nombre real).
15. Declares.bas:476 → `Private Type tMapaConnect` → `Public Type tMapaConnect` + `Dim MapaConnect As tMapaConnect` local en `RenderConnect` (1835; nunca se asigna → (0,0,0)).
16. TileEngine.bas `RenderConnect`: **comentados** el bloque preview `If Not frmCuenta.ListPJ.ListIndex < 0 Then With Cuenta.pjs(...)` (~1885-1891) y el bloque `If frmCuenta.Visible = True Then` de consejos (~1906-1909). `frmCuenta`/`Cuenta` no existen en el activo.
17. TileEngine.bas: las 10 asignaciones `frmCargando.Status.Caption =` → `.Text =` (Status es **RichTextBox** RICHTX32; `.Caption` no existe). Líneas ~2176,2180,2184,2188,2192,2196,2200,2204,2208,2213.
18. frmCargando.frm: añadido `Public Sub progresoConDelay(ByVal porcentaje As Integer)` (solo `DoEvents`; el formulario activo no tiene `imgProgress` — el de la referencia ee sí). El `progress` también.
19. TileEngine.bas `DrawSpells` (2536): eliminada la llamada muerta `Spells.DrawSpells` (renderiza a `frmMain.picSpell`).
20. General.bas: añadida `Public Function ColorToDX8(ByVal long_color As Long) As Long` (de `ee\Cliente\CODIGO\General.bas:1531-1548`; RGB hex → `D3DColorXRGB`). Necesaria en `clsDX8Font.cls:89` (`Text_Render_Special`).
21. clsTexManager.cls:131: `D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, ...)` → `mD3D.CreateTextureFromFileInMemoryEx(device, ...)` (privados del class; NO existen globals `D3DX`/`D3DDevice`).
22. clsBufferMan.cls:432,453: `Call CopyMemory(...)` → `Call CopyMem(...)` (la clase declara `Private Declare Sub CopyMem ... Alias "RtlMoveMemory"` en la 74; `CopyMemory` solo es Private en clsDX8Font/Bass.bas → invisible).
23. clsSoundEngine.cls:1067,1105: `General_Distance_Get` no existía en ningún sitio (ni en ee/Aodrag9). Añadida a General.bas (de `K:\Descargas\aaoo\Aodrag9\codigo\General.bas:1709-1718`, distancia Manhattan `Abs(x1-x2)+Abs(y1-y2)`).
24. frmRenderConnect.frm: `Audio.General_Set_Wav(SND_WAV_CLICK)` → `PlayWaveDS(SND_WAV_CLICK)` (Mod_Wav está en el proyecto; `Audio` NO existe, el motor es `Sound`). btnConsejo_Click protegido: `If UBound(ListaConsejos) >= 1 Then ...`.
25. Declares.bas: añadidas `Public Consejos(1 To 100) As String`, `Public ListaConsejos() As String`, `Public Form_Caption As String` (antes solo `ConsejoSeleccionado`).
26. General.bas: añadidos `Public Sub Auto_Drag(ByVal hwnd As Long)` (ReleaseCapture + SendMessage WM_NCLBUTTONDOWN/HTCAPTION) y `Public Sub CloseClient()` (Sound.Engine_DeInitialize + End). `Form_Caption = "AOSpain v" & App.Major & "." & App.Minor & " Beta: 1"` fijado en `Sub Main` (junto a `frmConnect.version`).
27. APIdeclaraciones.bas: añadidos `ReleaseCapture`, `SendMessage`, `WM_NCLBUTTONDOWN = &HA1`, `HTCAPTION = 2`.

## LOADERS IMPLEMENTADOS (TileEngine.bas, sección «Loaders migrados» al final ~3380+)
- `LoadGrhData` (doble pasada: halla maxGrh → dimensiona `GrhData(1..maxGrh)` → lee registros; estáticos rellenan FileNum/SX/SY/pixelW/H/TileW/H/Frames(1)), `CargarCuerpos`, `CargarCabezas`, `CargarCascos` (estos dos derivan `heads()/Cascos().Texture=FileNum,startX=SX,startY=SY` desde el grh 1), `CargarFxs`, `CargarAtaques` (tolera ausencia), `LoadMiniMap` (no-op), `CargarParticulas` (no-op).
- `IniPath = App.Path & "\Init\"` añadido en InitTileEngine.
- Declaraciones añadidas (~TileEngine.bas:535): `GrhCount`, `NumHeads`, `NumCascos`, `NumCuerpos`, `NumFxs`, `NumAtaques`, `heads() As tHead`, `Cascos() As tHead`.
- Sin `Ataques.ind`/`minimap.dat`/`particulas.ini` en el cliente → CargarAtaques tolera ausencia; MiniMap/Particulas no-op.

## FORMATO `Init\Graficos.ind` (CONFIRMADO por parseo, EOF exacto 180693 bytes)
- `tCabecera` (263B) + 5×Int16 + registros: `Grh` Int16, `NumFrames` Int16; si `nf>1`: `nf×Int16` frames + Int16 speed; si no: FileNum Int16 + SX Int16 + SY Int16 + pixelWidth Int16 + pixelHeight Int16; lista termina con `Grh<=0`.
- Resultado: 12597 registros, `maxGrh=19625`.
- Misma cabecera exacta que `k:\Descargas\aaoo\dx\Cliente\Init\Graficos.ind` (`Argentum Online by Noland Studios...`, bytes 255-267 idénticos).
- `Cabezas.ind` (5465B) = cabecera + 650 cabezas × 8B (tIndiceCabeza Integer×4). `Personajes.ind` (3613B) = 279 cuerpos × 12B. `Fxs.ind` = 20 fxs. `Cascos.ind` = 38 cascos.
- Los loaders usan Int16 (NO el loader de dx que mezcla Long).

## REFERENCIAS EXTERNAS
- `k:\Descargas\aaoo\dx\Cliente` — cliente DX8 de referencia (mismo formato); `dx\Cliente\codigo\TileEngine.bas:937` LoadGrhData (pero su lectura Long no alinea; usar Int16).
- `k:\Descargas\aaoo\ee\Cliente\CODIGO\Mod_Carga.bas` — loaders canónicos con `MiCabecera` (CargarAtaques:30, CargarCascos:70, LoadGrhData:283).
- `ee\Cliente\CODIGO\Declares.bas` — CASPER_HEAD/FRAGATA_FANTASMAL (135-136), `Public Cuenta As acc` (594).
- `ee\Cliente\CODIGO\General.bas` — `Function ARGB` (1127-1146).
- Backup DX7: `I:\AospainOri\cli_backup_dx7_20260728_1510\codigo\TCP.bas:426` (Case "CP" antiguo, HeadData).
- En activo: `Type MapInfo` (TileEngine.bas:200), `Position` (Declares.bas:364 X/Y As Long), `tHead` (Declares.bas:642), `tIndiceAtaque` (648 Body(1..4) As Long), `charlist.Body As BodyData` (TileEngine.bas:128).
- `Client.vbp` sin `clsGraphicalInventory`, sin `Mod_Carga`/`clsIniReader`; forms actuales: frmCantidad, frmCargando, frmComerciar, frmConnect, frmMain, frmOldPersonaje, frmSkills3, frmRenderConnect, frmMSG, frmForo, FrmEstadisticas, frmtip, frmPres, frmMensaje, frmCrearPersonaje, frmBorrar, frmRecuperar, frmPasswd, frmEntrenador, frmSpawnList, ...

## PRÓXIMO PASO (mañana) — B2 en curso

### ESTADO B2 (corte 08/08/2026) — Task 1 COMPLETA + fixes post-review, compilación EXIT 0
Commits en `main` desde BASE `159e7c5`:
- `47ce97e` — Task 1: protocolo de cuentas en el cliente (cli/codigo/TCP.bas).
- `5437b74` — fixes post-review: PASSCL orophin antes de NLOGIN de cuenta + field 38 = `ModValCoDe` (valcode real); bloque `[B2 ACCOUNT]` GDI de TileEngine.bas movido al tope del módulo.

**Lección crítica (verificada contra el server Ori):** el NLOGIN de cuentas EXIGE `PASSCL=orophin` ANTES (Servidor TCP.bas:1715) y field 38 == `ValidarLoginMSG(ValCoDe)` del handshake VAL (1735). El server dx (referencia `I:\Aospain1.0-dx`) NO tiene esos gates — por eso el plan original ("SIN PASSCL", `Login(0)`) era incorrecto. El handler `VAL` del cli guarda ahora el valcode en el global `Public ModValCoDe As Integer` (cli TCP.bas, bloque 3-char) y `Case CrearNuevoPj` lo usa para field 38.
**Benignos (NO arreglar):** ALOGIN field 4 truncado (ConnectAccount solo lee fields 1-2); VAL solo llega una vez.

### LO QUE FALTA DE B2 (mañana)
1. **Task 2** (brief `I:\AospainOri\.superpowers\sdd\memoria.md\task-2-brief.md`):
   - `frmAccLogin.cmdLogin_Click` → conectar socket antes de ALOGIN (`If Not frmMain.Socket1.Connected Then ... HostName=CurServerIp/RemotePort=CurServerPort/Connect`), mismo en cmdCrear_Click. OJO no conectar dos veces; orden: conectar → setear EstadoLogin.
   - `frmCuent.EntrarAlMundo` (1281-1327) → solo desconectar+reconectar si `Not Connected`; si ya está conectado, mandar `OOLOGI` directo.
   - `frmCrearPersonaje.boton_Click(0)` (975-1009) → rama cuenta: `If EstadoLogin = LoginAccount Or EstadoLogin = Dados Or frmCuent.Visible Then EstadoLogin = CrearNuevoPj: Call Login(0) (si Connected, si no frmMensaje)`; si no → `frmPasswd.Show vbModal` (clásico intacto).
   - Verificar que `codigo\Cuentas\DrawPJenPicture.bas` está como Module en `cli\Client.vbp`.
2. **Task 3** (server): `Case "/SALIR"` (2203-2210) → tras `FINOK`, si `UserList(...).Accounted <> ""` llamar `EnviarListaPJs(UserIndex, Accounted)` y NO cerrar socket.
3. Compilar limpio final (`vb6 /make`, borrar log antes) y commit B2.

### Recordatorios
- Flujo clásico (PASSCL+OLOGIN/NLOGIN, frmOldPersonaje/frmPasswd/frmBorrar/frmRecuperar PASSRECO) debe quedar 100% intacto.
- Compilar: borrar `C:\Users\sonsc\AppData\Local\Temp\opencode\vb6err.txt` antes de cada `vb6 /make`.
- Reference servers: `k:\Descargas\aaoo\dx\Cliente` (DX8), `ee\Cliente` (canónico), `I:\Aospain1.0-dx\Cliente_DX8` (cuentas).
