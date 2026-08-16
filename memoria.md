# MEMORIA - Proyecto AospainOri (Cliente AO DX8)

## Objetivo
Migrar cliente AO de DX7 a DX8, manteniendo funcionalidad idéntica.

## Compilación
- Borrar log antes de cada run: `Remove-Item vb6err.txt; Remove-Item errores.log`
- Comando: `& "C:\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.EXE" /make Client.vbp /out vb6err.txt`
- Éxito = `La generacion de '...' ha tenido exito.`
- VBP: `I:\AospainOri\cli\Client.vbp`
- Encoding: .frm/.bas = CP1252+CRLF. Usar PowerShell para edits, NUNCA la tool de edit/write.

## Referencias
- **Cliente funcional (DX7):** `k:\Descargas\aaoo\Aodrag9\codigo\TileEngine.bas`
- **Cliente DX7 original:** `k:\Argentum\AOspain1.0\cli\`
- **Los .ind son IDÉNTICOS** entre todos los clientes

## Estado Actual ✅
- **EXE funciona** — no crashea al login
- **Inventario funcional** — DrawGrhtoHdc con BMP header offset=14 + color table 8-bit
- **Personaje visible** — DrawHead con hardcoded 27x32 (como funcional)
- **Recordar cuenta** funcional
- **Mod_Cripto.bas** agregado
- **Error logging** a errores.log

## Pendiente ❌
- **Cabeza se ve mal** — DrawHead hardcoded 27x32 pero los GRH reales son 16x16 (BMPs 64x16). La textura se recorta mal. Necesita adaptación o texturas correctas.
- **Cuerpo pequeño** — GRH 4582 animado, primer frame pw=26 ph=46. Puede ser correcto o puede estar escalado mal.
- **Mapa: black area abajo** — puede ser por screenminY/screenmaxY o por el tile buffer
- **Velocidad movimiento** — engineBaseSpeed puede necesitar ajuste

## Lecciones Clave
1. **ScreenX/ScreenY en RenderScreen:** NO inicializar a 1. El funcional los deja en 0 (default VB6). Inicializar a 1 desplaza el suelo 32px y causa crash.
2. **DrawHead:** El funcional usa hardcoded textureX2=27, textureY2=32 con offset -3/-3. No usar GRH dimensions (causa problemas con texturas pequeñas).
3. **Error handling:** On Error Resume Next en RenderScreen (como funcional). On Error GoTo puede causar problemas con Direct3D.
4. **DrawGrhtoHdc:** BMP file header = 14 bytes (no 2). Para 8-bit BMPs incluir color table (256×4=1024 bytes).
5. **Error handlers:** MsgBox+End en EXE cierra la app. Cambiar a LogError+ExitSub.
6. **Comparar con funcional** antes de hacer cambios grandes.

## Archivos Relevantes
- `cli\codigo\TileEngine.bas` — RenderScreen, CharRender, DrawHead, DDrawTransGrhtoSurface, DrawGrhtoHdc, LoadGrhData, CargarCabezas
- `cli\codigo\Modulo_DibujarInventario.bas` — DibujarInv (FUNCIONA)
- `cli\codigo\General.bas` — SwitchMap, DirGraficos
- `cli\codigo\Declares.bas` — Type char (Head As Integer, Casco As Integer)
- `cli\codigo\Mod_ErrorLOG.bas` — LogError
