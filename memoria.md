# MEMORIA — Cliente Argentum Online DX8 (aoscorpio2) `I:\AospainOri`

## IMPORTANTE (no olvidar)
- **aospain NO usa compresión todavía.** NO implantar compresión por ahora. Existe material de referencia (modCompression, `GameIni tCabecera`, compresor AODrag) pero queda en reserva para futuro.
- **ENCODING: los .frm/.bas son CP1252 + CRLF.** VB6 no carga archivos UTF-8/LF. Si se edita con herramientas que re-escriben el archivo, verificar bytes: acentos = 1 byte (E9=é, F3=ó, E1=á), saltos = `0D 0A`. Un diff que "toca" TODAS las líneas con acentos = corrupción de encoding.
- Objetivo: mantener **compilación limpia** `vb6 /make` → `I:\AospainOri\cli\aoscorpio2.exe` y `I:\AospainOri\Servidor\AOSpainServ1.0.4-.exe`.

## Compilación (comando)
- Borrar el log ANTES de cada run. Cliente: `C:\Users\sonsc\AppData\Local\Temp\opencode\vb6err.txt`; servidor: `C:\Users\sonsc\AppData\Local\Temp\opencode\vb6srv.txt`.
- `& "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make I:\AospainOri\cli\Client.vbp /out <log>` y lo mismo con `I:\AospainOri\Servidor\SERVER.VBP`.
- Éxito = `La generación de '...' ha tenido éxito.` VB6 reporta PRIMER error del módulo y para; desfase de línea aprox ±2.
- **OJO:** `SERVER.VBP` tiene `AutoIncrementVer=1` → cada compilación sube `RevisionVer` y deja el .vbp "modificado". Revertir con `git checkout -- Servidor/SERVER.VBP` si no queremos incluirlo.
- **Para editar .frm/.bas de forma segura:** leer/escribir con `[System.Text.Encoding]::GetEncoding(1252)` desde PowerShell y usar `"`r`n"` como salto. NO usar el tool de edit/write sobre estos archivos (escribe UTF-8/LF).

## ESTADO ACTUAL (corte 13/08/2026) ✅ B3 CERRADO y COMMITEADO
- B1 (migración DX7→DX8) ✅ y B2 (account system Tasks 1-3) ✅ VERIFICADOS en runtime por el usuario.
- **B3 COMPLETO** (frmConnect del ref: login directo + 1 botón por acción) — commit `cd9ae46`. Compila cliente y servidor. **PENDIENTE de verificación en runtime (probar cuenta nueva + HLQ, login directo MD5, recuperar, eliminar).**

## LO QUE SE HIZO EN B3 (verificado, commit `cd9ae46`)
- `cli\codigo\frmConnect.frm` + `frmConnect.frx` (408043 B, Picture frx:000C) copiados y reescritos del ref (escritos en UTF-8/LF → **corregidos a CP1252+CRLF** para que VB6 cargue).
- `Conectar_Click(Index)` → valida, limpia socket previo (`Disconnect`+`Cleanup`+DoEvents), `nombrecuent/passcuent`, **`UserPassword=MD5String(passcuent)`, `MD5HushYo=UserPassword`**, `EstadoLogin=LoginAccount`, conecta a `CurServerIp/CurServerPort`. NO llama Login() (lo dispara VAL).
- `CrearPersonaje_Click` → si no conectado: `EstadoLogin=CrearAccount`, conecta, `DoEvents`; luego `frmCrearAccount.Show vbModal, Me` (**fix 24038**, envío NACCNT con socket vivo).
- `RecuperarCuenta_Click` → `frmRecuperarCuenta.Show vbModal, Me` (nota: ref usaba frmRecuperarpj, inexistente).
- `EliminarCuenta_Click` → 4×InputBox (account, pass, mail, respuesta), espera conexión máx 3s (patrón GetTickCount de frmCuent:1316-1325), `SendData("BRCU" & sAccount & ",," & sPass & "," & sMail & "," & sRespuesta)` — **campo 2 (Nombre) VACÍO** (validado: CreateAccount server NO escribe "Nombre" en el .act → `GetVar("Nombre")=""` y `""=""` pasa, TCP.bas:1986).
- `Form_KeyDown` ESC → `DeinitTileEngine` (LiberarObjetosDX no existe) + `SaveGameini` + `prgRun=False` + `UnloadAllForms`.
- `Form_KeyUp` Ctrl+I togglea `PortTxt.Visible/IPTxt.Visible`. `Form_Load`: `EngineRun=False`, `PortTxt=Config_Inicio.Puerto`. `Form_Activate`: IP/port de `ServersLst(CurServer)` o `IPdelServidor`/`PuertoDelServidor`, `CargarLst`, MkDir Web.
- **`TCP.bas` cliente añadió `Case "HLQ"`** (~724, antes de ERR): `MsgBox "La cuenta fue creada con éxito."` — confirmación REAL vía server.
- **`frmCrearAccount.frm Image2_Click`**: ELIMINADO el MsgBox falso de éxito ("La cuenta fue creada con éxito."). Queda `SendData("NACCNT"...)` + `Unload Me`. (Al editarlo se corrompió `términos` EF BF BD → arreglado a E9).
- **Server `CreateAccount`** (TCP.bas ~682-726): **fix del HLQ perdido** — `CloseSocket` ponía `ConnID=-1` (763) + `Cleanup`/`Unload Socket2`/`ResetUserSlot` antes del `SendData(ToIndex)` (que solo escribe si `ConnID>-1`, 913). Ahora: **`SendData("HLQ")` ANTES de cerrar** + espera 300ms (`GetTickCount` loop con DoEvents) + `CloseSocket`.
- **`Declares.bas` servidor**: añadido `Public Declare Function GetTickCount Lib "kernel32" () As Long` (tras sndPlaySound, ~linea 1221).
- **GOTCHA de encoding en servidor:** en sesión previa el `TCP.bas` del servidor quedó re-encodificado a UTF-8/`?` (237 líneas de diff falso, acentos→bytes 3F). Se RESTAURÓ con `git checkout` y se re-aplicó el fix con PowerShell en CP1252 → diff actual limpio (solo 11 líneas). **Los EF BF BD que queden (72 en TCP.bas cliente, 1 en server TCP.bas, 15 en Declares.bas) son preexistentes del repo, NO tocar.**

## FLUJO DE LOGIN (mi cliente + mi server, NO tocar)
- `Socket1_Connect` (frmMain.frm:1094-1123): envía `gIvEmEvAlcOde` (si frmCrearPersonaje.Visible o Not frmRecuperar.Visible) o `PASSRECO` (frmRecuperar.Visible). **MI server exige gIvEmEvAlcOde primero → NO copiar ramas E_MODO del ref.**
- Server handshake (1643-1651): `gIvEmEvAlcOde` → `SendData("VAL" & RandKey & "," & ValCoDe)`.
- Cliente `Case "VAL"` (TCP.bas:499-509): `bK=ReadField(1)`, `bO=100`, `ModValCoDe=ValidarLoginMSG(ReadField(2))`, `Login(ModValCoDe)`. (Si frmBorrar.Visible → BORR directo).
- `TCP.bas Login()` (1090-1155): si `EstadoLogin` 1..7 → Select Case **Normal/LoginAccount/BorrarPj/CrearNuevoPj**; **CrearAccount (4) sin Case → Exit Sub sin enviar** (deseado: evita NACCNT prematuro). Si EstadoLogin fuera de rango (0) → cae a OLOGIN/NLOGIN según `SendNewChar` (❗ cuando BRCU conecta sin setear EstadoLogin, el VAL dispara `Login(0)` y envía OOLOGI/OLOGIN con UserName vacío → server responde ERR "no existe" pero NO cierra → BRCU igual se procesa. Es ruido benigno heredado del ref, no daña).
- `ALOGIN` (field1=cuenta, field2=MD5) → `ConnectAccount` (626-638) compara contra `GetVar(.act,"password")` = MD5String → el cliente DEBE mandar MD5 (verificado B3).
- `.act` se crea con: `password=MD5String`, `mail`, `Pregunta`, `Respuesta`, `ban=0`, `[PJS] NumPjs=0 + PJ1-8=""`, **sin clave "Nombre"**.

## Next steps (próxima sesión)
1. **Verificación runtime** por el usuario (arrancar server + cliente con `aoscorpio2.exe`):
   - Crear cuenta nueva → esperar MsgBox REAL "La cuenta fue creada con éxito." (HLQ) y SIN MsgBox falso del form.
   - Login directo con esa cuenta (ALOGIN con MD5) → debe entrar al alta de personaje (INIPJ flow).
   - Recuperar cuenta (frmRecuperarCuenta RECCUU/REECUU).
   - Eliminar cuenta (BRCU con campo2 vacío, pass por MD5).
2. Si algo falla en runtime: revisar el flujo exacto indicado arriba; NO alterar Socket1_Connect ni el orden de handshake.
3. Opcional pendiente de decidir: `TCP.bas EnviarCrearCuenta` (1070-1072) quedó huérfana (abre frmCrearAccount con owner frmAccLogin) — no molesta (frmAccLogin sigue en el vbp).

## Work history (commits)
- `cd9ae46` — B3: frmConnect login directo con MD5, confirmación HLQ real, fix HLQ antes de CloseSocket en servidor (6 archivos).
- `6bf7a41` — docs: estado B3 + plan migración frmConnect.
- `dcd0458` — feat: imgGetPass → frmRecuperarCuenta.
- `97c5639` / `50a4587` — docs B2 verificado / cerrado.
- `40e4819` — feat: servidor re-envía lista PJs tras /SALIR.

## REFERENCIAS EXTERNAS
- `I:\Aospain1.0-dx\Cliente_DX8\codigo\frmConnect.frm` + `.frx` — plantilla B3 ya copiada/adaptada.
- `I:\Aospain1.0-dx\Cliente_DX8\codigo\TCP.bas` 786-806 (HLQ/ERR), 1214-1217 (Case CrearAccount→NACCNT — NO usar en Login).
- `I:\Aospain1.0-dx\Cliente_DX8\codigo\Cuentas\frmCrearAccount.frm` — igual al mío pero sin MsgBox falso.
- `I:\Aospain1.0-dx\Servidor\Codigo\Modulos\TCP.bas` 425-471 (CreateAccount HLQ), 2016-2069 (BRCU formato ref, distinto del mío).
- `k:\Descargas\aaoo\dx\Cliente` — cliente DX8 ref; `ee\Cliente\CODIGO` — canónico. Backup DX7: `I:\AospainOri\cli_backup_dx7_20260728_1510`.
- SDD ledger: `I:\AospainOri\.superpowers\sdd\memoria.md\progress.md`.

## Recordatorios
- Flujo clásico (PASSCL+OLOGIN/NLOGIN, frmOldPersonaje/frmPasswd/frmBorrar/frmRecuperar PASSRECO) intacto.
- Compilar: borrar el log ANTES de cada `vb6 /make`.
- Skill "memory" global en `C:\Users\sonsc\.agents\skills\memory\SKILL.md`.