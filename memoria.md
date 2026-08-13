# MEMORIA — Cliente Argentum Online DX8 (aoscorpio2) `I:\AospainOri`

## IMPORTANTE (no olvidar)
- **aospain NO usa compresión todavía.** NO implantar compresión por ahora. Existe material de referencia (modCompression, `GameIni tCabecera`, compresor AODrag) pero queda en reserva para futuro.
- Objetivo actual: **compilar limpio** `vb6 /make` y obtener `I:\AospainOri\cli\aoscorpio2.exe`.

## Compilación (comando)
- Borrar el log ANTES de cada run: `C:\Users\sonsc\AppData\Local\Temp\opencode\vb6err.txt`.
- `& "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make I:\AospainOri\cli\Client.vbp /out C:\Users\sonsc\AppData\Local\Temp\opencode\vb6err.txt`
- Los errores de VB6 pueden reportar línea con **desfase (+2 aprox)** respecto a la línea real.
- VB6 reporta el PRIMER error del módulo y para → errores encadenados, corregir uno a uno.

## ESTADO ACTUAL (corte 13/08/2026) 🚧 B3 EN CURSO — Replicar `frmConnect` del ref (Cliente_DX8) para login directo + 1 botón por acción
- B1 (migración DX7→DX8) ✅ y B2 (sistema de cuentas: Tasks 1-3) ✅ CERRADOS y VERIFICADOS en runtime por el usuario (12/08/2026).
- **Nueva tarea B3 (decidida por el usuario con pregunta):** reemplazar el `frmConnect` actual (menú `Image1(0-7)` + botón `cmdAccLogin` que abría `frmAccLogin` modal = bug 2) por el `frmConnect` del ref `I:\Aospain1.0-dx\Cliente_DX8\codigo\frmConnect.frm`:
  1. **Login directo** (cuenta+pass en la pantalla de conexión).
  2. **Un botón por acción**: Conectar / Crear Cuenta / Recuperar / Eliminar Cuenta.
  3. **Servidor destino = el mío** (`I:\AospainOri\Servidor`), que ya tiene ALOGIN/NACCNT/INIAC/RECCUU/BRCU.
- Idioma SIEMPRE en español.

## HALLAZGOS DE LA INVESTIGACIÓN (13/08/2026, para ejecutar B3)
### Fondo idéntico → coordenadas del ref válidas
- `I:\AospainOri\cli\Graficos\Conectar.jpg` y `I:\Aospain1.0-dx\Cliente_DX8\Graficos\Conectar.jpg` → **mismo MD5 `94EB6E11B02D30DF81C767434B70AF99`** (263227 bytes). Se puede copiar el form ref tal cual visualmente.

### Ref `frmConnect.frm` (Cliente_DX8, 408 líneas, ScaleMode=0 User, ScaleWidth=799, ScaleHeight=471.094, Picture `frx:000C`)
- Controles clave (twips): `NameTxt` (5055,3401), `PasswordTxt` (5055,4700, PasswordChar=*), `Conectar` Image Index=1 (6240,8040), `CrearPersonaje` Image (4320,7920), `RecuperarCuenta` Image (2280,8040), `EliminarCuenta` Image (8160,7920), `version` Label, `IPTxt`/`PortTxt`/`DescTxt` (ocultos, Ctrl+I los togglea), `lst_servers`, `Text1`/`Text2`.
- **`Conectar_Click(Index)`**: si Index≠1 Exit; valida campos; limpia socket previo; `nombrecuent=NameTxt`, `passcuent=PasswordTxt`, `UserPassword=passcuent` (**texto plano — OJO, adaptar a MD5**); `EstadoLogin=LoginAccount`; `Me.MousePointer=99`; `Socket1.Connect`. **NO llama Login()** → lo hace el handler VAL.
- **`CrearPersonaje_Click`** (crear cuenta, fix del 24038): si no conectado → `EstadoLogin=CrearAccount`, `Socket1.HostAddress=CurServerIp`, `RemotePort=CurServerPort`, `Socket1.Connect`, `DoEvents`; luego `frmCrearAccount.Show vbModal, Me`.
- **`RecuperarCuenta_Click`**: `frmRecuperarpj.Show vbModal, Me` — **frmRecuperarpj NO existe en mi cliente** → usar `frmRecuperarCuenta.Show vbModal, Me`.
- **`EliminarCuenta_Click`**: 2×InputBox (cuenta; luego "CONTRASEÑA, CORREO y RESPUESTA separados por comas") → conecta si no → `SendData("BRCU" & sAccount & "," & sConfirm)`. **FORMATO ref (2 campos) ≠ MI server (5 campos) → adaptar.**
- `Form_KeyDown` ESC: usa `LiberarObjetosDX` (**no existe en mi cliente** → usar `DeinitTileEngine` como mi frmConnect actual) + `SaveGameini` + `prgRun=False` + `UnloadAllForms`.
- `Form_KeyUp` Ctrl+I: toggle `PortTxt.Visible`/`IPTxt.Visible`.
- `Form_Load`: `EngineRun=False`; `PortTxt=Config_Inicio.Puerto`.
- `Form_Activate`: `CurServer<>0`→IP/port de `ServersLst(CurServer)`; si no de `IPdelServidor`/`PuertoDelServidor`; `MkDir Web`.
- `lst_servers_Click`/`CargarLst`: setean CurServer y textos.

### Mi cliente (estado actual)
- `frmConnect.frm` actual (564 líneas): ScaleMode=3 Pixel 800x600, Picture cargada en Form_Load (`LoadPicture Graficos\Conectar.jpg`), menú `Image1(0-7)`, `cmdAccLogin_Click`→`frmAccLogin.Show vbModal` (**bug 2 a eliminar**), `imgGetPass_Click`→`frmRecuperarCuenta.Show vbModal`, `lst_servers`+`text1` (news), `Command1`, `imgServEspana/Argentina`.
- **`TCP.bas` NO tiene `Case "HLQ"`** (falta éxito real de crear cuenta). El ref Cliente_DX8 SÍ lo tiene (TCP.bas:786-792 → `MsgBox "La cuenta fue creada con éxito..."`).
- `TCP.bas Login()` (1086-1151): `If EstadoLogin >= Normal And EstadoLogin <= RecuperarPass Then` con Cases **Normal/LoginAccount/BorrarPj/CrearNuevoPj** (CrearNuevoPj manda PASSCL+NLOGIN con field38=`ModValCoDe`); si no, OLOGIN/NLOGIN según `SendNewChar`. **CrearAccount (4) NO tiene Case** → cae en el `Select` sin match → `Exit Sub` sin enviar (comportamiento deseado para evitar NACCNT prematuro con campos vacíos).
- `TCP.bas EnviarLoginCuenta` (1062-1068): `nombrecuent=sNombre`; `UserPassword=MD5String(sPass)`; `MD5HushYo=UserPassword`; `EstadoLogin=LoginAccount`; `Login(0)`.
- `TCP.bas EnviarCrearCuenta` (1070-1072): solo `frmCrearAccount.Show vbModal, frmAccLogin`.
- `TCP.bas VAL` (499-509): `bK=ReadField(1)`; `bO=100`; `ModValCoDe=ValidarLoginMSG(ReadField(2))`; `Login(ModValCoDe)`. (Si `frmBorrar.Visible` → BORR directo).
- `TCP.bas ERR` (724-730): MsgBox; `If Not frmCrearPersonaje.Visible And EstadoLogin<>LoginAccount And EstadoLogin<>Normal Then Socket1.Disconnect`.
- `TCP.bas INIAC` (839-846) / `ADDPJ` (847-862): muestran frmCuent con lista de PJs.
- `TCP.bas SendData` (1042-1060): `bK=GenCrC(bK,sdData)`; `bO=bO+1`; `~bK&ENDC`; `Socket1.Write`.
- `frmMain.frm Socket1_Connect` (1094-1123): `MixedKey` desde IP; `Second.Enabled=True`; **si `frmCrearPersonaje.Visible` o `Not frmRecuperar.Visible` → `SendData("gIvEmEvAlcOde")`; si no → `PASSRECO` con frmRecuperar**. (No tiene ramas por EstadoLogin como el ref, pero MI server exige gIvEmEvAlcOde como primer paquete → **mantener este patrón**, NO copiar las ramas del ref que no envían gIvEmEvAlcOde).
- `frmCrearAccount.frm` (mi cliente, Cuentas\, 278 líneas): controles `Nombre,Pass,RePass,Mail,Mail2,pregunta,respuesta,Check1,Image1,Image2`. `Image2_Click` valida y `SendData("NACCNT" & nombre & "," & Pass & "," & Mail & "," & pregunta & "," & respuesta)` + `Unload Me` + **`MsgBox` falso "La cuenta fue creada con éxito." (líneas 266-269, A ELIMINAR)**. El ref idéntico pero SIN el MsgBox falso.
- `frmRecuperarCuenta.frm` (Cuentas\, 174 líneas): `Siguiente_Click` conecta (`HostName=CurServerIp, RemotePort=CurServerPort, Connect`) y `SendData("RECCUU" & txtNombre & "," & txtMail)`; `Recuperar_Click` → `"REECUU" & txtNombre & "," & txtRespuesta`.
- `frmAccLogin.frm` (128 líneas): `cmdLogin` conecta si no y `EnviarLoginCuenta`; `cmdCrear` muestra email, conecta y `EnviarCrearCuenta`.

### Mi servidor (Servidor\Codigo\Modulos\TCP.bas)
- Handshake (1643-1651): `gIvEmEvAlcOde` → `ValCoDe=RandomNumber(20000,32000)`, `RandKey=RandomNumber(0,99999)`, `PrevCRC=RandKey`, `PacketNumber=100`, `SendData("VAL" & RandKey & "," & ValCoDe)`.
- `ConnectAccount` (626-638): `Password <> GetVar(.act, name, "password")` → `ERR "Password incorrecto."` + CloseSocket. **El .act guarda `password=MD5String(Password)` (CreateAccount:697) → el cliente DEBE mandar el MD5 en ALOGIN field2.**
- `EnviarListaPJs` (640-660): `INIAC[name],[numPjs+1]` si `TienePjs`, si no `INIAC0`; luego `ADDPJ` por PJ.
- `CreateAccount` (682-726): crea `.act` (password=MD5String, mail, Pregunta, Respuesta, ban=0, [PJS] NumPjs=0, PJ1-8=""); `DoEvents`; **`CloseSocket(UserIndex)` (716); `SendData("HLQ")` (718) — HLQ DESPUÉS de cerrar (igual que ref server 461-463)**.
- `ALOGIN` (1841-1854): valida AsciiValidos + CuentaExiste → `ConnectAccount(field1, field2)`.
- `NACCNT` (1856-1864): 5 fields → `CreateAccount`.
- `OOLOGI` (1825-1839): PersonajeExiste → BANCheck → `ConnectUser`.
- `BRCU` (1967-2013): **5 campos `[cuenta],[Nombre],[Pass],[Mail],[Respuesta]`**; valida `bValidationData=UCase(GetVar("Nombre"))`, `MD5String(bPass)=realPass`, `bMail=realMail`, `bRespuesta=realRespuesta`; borra el último PJ del .act + su .chr.
  - ⚠️ **OJO:** `CreateAccount` NO escribe "Nombre" en el .act → `GetVar("Nombre")` devuelve "" y el check de campo2 fallaría para cualquier nombre no vacío. **PENDIENTE DE CLARIFICAR/PROBAR con el usuario antes de dar por bueno el botón EliminarCuenta.**
  - El ref server BRCU (2016-2069) usa OTRO formato ([cuenta]+[pass,mail,respuesta] juntos, sin MD5) → **no sirve de referencia para mi server; adaptar al formato de MI server.**

## PLAN DE EJECUCIÓN B3 (decisión tomada en sesión 13/08/2026)
1. **Copiar `frmConnect.frm` + `frmConnect.frx` del ref** (`I:\Aospain1.0-dx\Cliente_DX8\codigo\`) a `I:\AospainOri\cli\codigo\`, con adaptaciones en el código:
   - `RecuperarCuenta_Click` → `frmRecuperarCuenta.Show vbModal, Me`.
   - `Conectar_Click` → **`UserPassword = MD5String(passcuent)` y `MD5HushYo = UserPassword`** (mi server guarda MD5). No llamar Login() directo; que el VAL handler dispare ALOGIN (flujo ya probado en B2).
   - `CrearPersonaje_Click` → mantener (conecta primero con `EstadoLogin=CrearAccount`, muestra `frmCrearAccount`).
   - `EliminarCuenta_Click` → adaptar al formato de MI server (5 campos) — **revisar el tema "Nombre" del server antes/después**.
   - `Form_KeyDown` ESC → `LiberarObjetosDX` → `DeinitTileEngine` (como mi frmConnect actual).
   - `Form_KeyUp` Ctrl+I → mantener toggle.
   - Cuidar `frmConnect.frx` ref (408043 B, contiene Picture+iconos+ItemData lst_servers) — copiarlo junto con el .frm (mi frx actual es de 36 B).
2. **`TCP.bas`:** añadir `Case "HLQ"` (MsgBox "La cuenta fue creada con éxito." — éxito REAL vía server). **NO añadir Case CrearAccount a Login()** (evita NACCNT prematuro con campos vacíos; el envío lo hace `frmCrearAccount.Image2_Click`).
3. **`frmCrearAccount.frm Image2_Click`:** eliminar el `MsgBox` falso de éxito (dejar `Unload Me`; la confirmación llega por HLQ/ERR).
4. **`frmMain.frm Socket1_Connect`:** **NO copiar las ramas del ref** (mi server exige `gIvEmEvAlcOde` primero); mantener patrón actual. Verificar que para `EstadoLogin=CrearAccount` fluya bien (envía gIvEmEvAlcOde, VAL fija bK, luego Image2_Click envía NACCNT con socket ya conectado → **fix del 24038**).
5. Compilar cliente (`vb6 /make`) y, si se toca, servidor. Verificar "ha tenido éxito".
6. Commit.

## Referencias externas a controles de frmConnect que DEBEN preservarse al copiar el form ref
- `General.bas:860/871/881` → `frmConnect.PortTxt` / `frmConnect.IPTxt` (CInt). El ref los tiene como TextBox ✅.
- `General.bas:940` → `frmConnect.version = "v" & ...` (ref tiene Label version ✅).
- `frmSkills2.frm:297-298` → `frmConnect.IPTxt.Text` / `frmConnect.PortTxt.Text` (TextBox ✅).
- `frmMain.frm:1132,1171,1184` → `frmConnect.Visible/MousePointer/Show`.
- `frmCuent.frm:1344` → `frmConnect.Show`; `frmOldPersonaje.frm:155` → `frmConnect.MousePointer=11`.
- `TCP.bas:134,137,845` → `frmConnect.Visible` / `Unload frmConnect`.
- `frmCrearPersonaje.frm:1033` → `frmConnect.Picture = LoadPicture(...conectar.jpg)` ✅.
- `frmCrearCaracter.frm:181` → `frmConnect.FONDO.Picture` ⚠️ **frmCrearCaracter NO está en Client.vbp** (no se compila) → ignorar.

## FORMATO `Init\Graficos.ind` (CONFIRMADO por parseo, EOF exacto 180693 bytes)
- `tCabecera` (263B) + 5×Int16 + registros: `Grh` Int16, `NumFrames` Int16; si `nf>1`: `nf×Int16` frames + Int16 speed; si no: FileNum Int16 + SX Int16 + SY Int16 + pixelWidth Int16 + pixelHeight Int16; lista termina con `Grh<=0`.
- Resultado: 12597 registros, `maxGrh=19625`. Misma cabecera exacta que `k:\Descargas\aaoo\dx\Cliente\Init\Graficos.ind`.
- `Cabezas.ind` (5465B) = 650 cabezas × 8B. `Personajes.ind` (3613B) = 279 cuerpos × 12B. `Fxs.ind` = 20 fxs. `Cascos.ind` = 38 cascos. Loaders usan Int16.

## REFERENCIAS EXTERNAS
- `I:\Aospain1.0-dx\Cliente_DX8\codigo\frmConnect.frm` + `.frx` — **plantilla B3** (login directo, botones por acción, fix 24038).
- `I:\Aospain1.0-dx\Cliente_DX8\codigo\frmMain.frm` 1077-1120 — `Socket1_Connect` con ramas E_MODO (NO copiar; solo informativo).
- `I:\Aospain1.0-dx\Cliente_DX8\codigo\TCP.bas` 786-806 (`Case HLQ`, `Case ERR`), 1214-1217 (`Case CrearAccount`→NACCNT — NO usar en Login, ver plan).
- `I:\Aospain1.0-dx\Cliente_DX8\codigo\Cuentas\frmCrearAccount.frm` — igual al mío pero sin MsgBox falso (267-269).
- `I:\Aospain1.0-dx\Servidor\Codigo\Modulos\TCP.bas` 425-471 (CreateAccount HLQ), 2016-2069 (BRCU formato ref).
- `k:\Descargas\aaoo\dx\Cliente` — cliente DX8 de referencia; `ee\Cliente\CODIGO` — canónico.
- Backup DX7: `I:\AospainOri\cli_backup_dx7_20260728_1510\codigo\TCP.bas:426`.
- SDD ledger: `I:\AospainOri\.superpowers\sdd\memoria.md\progress.md`.

## Recordatorios
- Flujo clásico (PASSCL+OLOGIN/NLOGIN, frmOldPersonaje/frmPasswd/frmBorrar/frmRecuperar PASSRECO) queda intacto.
- Compilar: borrar `C:\Users\sonsc\AppData\Local\Temp\opencode\vb6err.txt` antes de cada `vb6 /make`.
- Skill "memory" instalado global en `C:\Users\sonsc\.agents\skills\memory\SKILL.md`.
