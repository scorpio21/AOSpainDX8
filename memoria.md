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
- **Cliente DX8 referencia:** `i:\Aospain1.0-dx\Cliente_DX8\` ← NUEVO - mismo bug en DrawPJenPicture
- **Los .ind son IDÉNTICOS** entre todos los clientes

## Estado Actual ✅
- **EXE funciona** — no crashea al login
- **Inventario funcional** — DrawGrhtoHdc con BMP header offset=14 + color table 8-bit
- **Personaje visible** — DrawHead con hardcoded 27x32 (como funcional)
- **Recordar cuenta** funcional
- **Mod_Cripto.bas** agregado
- **Error logging** a errores.log

---

## Sesión 2026-08-16 — Fixes Aplicados ✅

### 1. Anti-Fade de mapa (negro al entrar)
**Archivo:** `TileEngine.bas` línea ~1270  
**Problema:** `ColorClimax` (struct RGBClimax) se inicializa a 0 por defecto en VB6 → mapa arranca NEGRO y tarda ~255 frames en iluminarse.  
**Fix:** En `InitColor()`, después de `AmbientColor.A = 255`, se inicializa `ColorClimax` al color diurno y se propaga a `base_light` y `TechoColor(0..3)`:
```vb
ColorClimax.r = 230 : ColorClimax.g = 200 : ColorClimax.b = 200 : ColorClimax.A = 255
base_light = ARGB(ColorClimax.r, ColorClimax.g, ColorClimax.b, ColorClimax.A)
For i = 0 To 3 : TechoColor(i) = base_light : Next i
```

### 2. Velocidad excesiva de movimiento
**Archivo:** `General.bas` línea ~1091  
**Problema:** `engineBaseSpeed = 0.019` → con 55ms/frame: `timerTicksPerFrame = 1.045` → personaje completa 1 tile en ~3 frames (muy rápido).  
**Fix:** `0.019` → `0.009`

### 3. Nombres de personajes invisibles
**Archivo:** `General.bas` (tras la línea de `Opciones.Audio`)  
**Problema:** `tOption.NamePlayers As Byte` arranca en `0` → `CharRender` no dibuja nombres.  
**Fix:** Añadida línea: `Opciones.NamePlayers = 1`

### 4. Spam de errores INACT en errores.log
**Archivo:** `TileEngine.bas` línea ~1998 (en `DDrawGrhtoSurface`)  
**Problema:** GRHs 5757-5759 tienen `NumFrames=0` en el .ind → `LogError "INACT..."` se llama miles de veces.  
**Note:** `DDrawTransGrhtoSurface` ya hace `Exit Sub` silencioso. Solo `DDrawGrhtoSurface` logueaba.  
**Fix:** Cambiado `LogError "INACT..." : Exit Sub` → `Exit Sub` (silencioso).

### 5. Cuerpos sin dibujar en pantalla de cuentas
**Archivo:** `codigo\Cuentas\DrawPJenPicture.bas`  
**Problema:** Archivo en disco tenía las ~185 líneas centrales **corrupto** (dos columnas de código mezcladas, probablemente por un editor que no respetó CRLF). La función `DibujaPJ` llamaba a `GrhRenderToHdc` con parámetros incorrectos.  
**Fix:** Archivo reescrito completamente con el código limpio (usuario lo proporcionó). Mejoras vs el original:
- Añadidos bounds-checks: `If Body > 0 And Body <= UBound(BodyData)` antes de cada acceso
- `dibujaban` implementado correctamente con `frmCuent.PJ(Index).Line(...)`
- `GrhRenderToHdc` existe en `CargasInit.bas` línea 464 — usa BitBlt con cache de Picture por FileNum

---

## Estructura del Motor — Sesión 2026-08-16

### Split de TileEngine.bas ✅
**Motivación:** TileEngine.bas tenía ~3841 líneas. Separado en dos módulos.

| Archivo | Líneas | Contenido |
|---|---|---|
| `TileEngine.bas` | ~3346 | Motor render: RenderScreen, CharRender, DrawHead, DDrawGrhtoSurface, DDrawTransGrhtoSurface, InitTileEngine, DirectXInit, MakeChar, partículas |
| `CargasInit.bas` (NUEVO) | ~501 | Carga de datos: DrawGrhtoHdc, LoadGrhData, CargarCuerpos, CargarCabezas, CargarCascos, CargarFxs, CargarAtaques, GrhRenderToHdc |

**Client.vbp actualizado:** línea 23 = `Module=CargasInit; CODIGO\CargasInit.bas`

### GrhRenderToHdc (CargasInit.bas línea 464)
- Usa `GetGrhPictureForHdc()` — cache de objetos Picture VB6 por FileNum
- Hace `BitBlt` desde MemDC al HDC del PictureBox
- Llaman: `DrawPJenPicture.DibujaPJ`, `frmCrearPersonaje`

---

## Pendiente ❌
- **Cabeza se ve mal** — DrawHead hardcoded 27x32 pero GRH reales son 16x16 (BMPs 64x16). Textura recortada mal. Ver si es correcto con el funcional.
- **Negro en mapa (parcial)** — GRHs 5757-5759 INACTIVOS en capa 1 → tiles sin dibujar. Posiblemente problema en Graficos.ind. Silenciado el spam pero los tiles siguen negros.
- **Heading=0 en personaje** — Si el servidor envía heading=0, `Walk(0)` crashea (array 1-based). Fix: En `MakeChar` validar `If Heading < 1 Then Heading = 3`.
- **EXE crash** — Comparar dependencias OCX/DLL. Revisar `App.Path` y orden de init antes del primer `On Error`.

## Lecciones Clave
1. **ScreenX/ScreenY en RenderScreen:** NO inicializar a 1. El funcional los deja en 0 (default VB6).
2. **DrawHead:** El funcional usa hardcoded textureX2=27, textureY2=32. No usar GRH dimensions.
3. **Error handling:** On Error Resume Next en RenderScreen (como funcional).
4. **DrawGrhtoHdc:** BMP file header = 14 bytes (no 2). Para 8-bit BMPs incluir color table (256×4).
5. **Error handlers:** MsgBox+End en EXE cierra la app. Cambiar a LogError+ExitSub.
6. **Corruption display:** PowerShell tiene un bug de display con líneas > ~100 chars → parecen "corruptas". El archivo en disco está bien. Usar Notepad++ para confirmar.
7. **Encoding PowerShell:** Siempre usar `[System.IO.File]::ReadAllLines($path, $enc)` + `$enc = GetEncoding(1252)`. Nunca `Get-Content -Encoding UTF8`.

## Archivos Relevantes
- `cli\codigo\TileEngine.bas` — RenderScreen, CharRender, DrawHead, DDrawTransGrhtoSurface, InitTileEngine, DirectXInit
- `cli\codigo\CargasInit.bas` — DrawGrhtoHdc, LoadGrhData, CargarCuerpos/Cabezas/Cascos, GrhRenderToHdc ← NUEVO
- `cli\codigo\Cuentas\DrawPJenPicture.bas` — render de cuerpos en pantalla de cuentas (REPARADO)
- `cli\codigo\Cuentas\frmCuent.frm` — pantalla selección de personaje (10 slots PJ(0..9))
- `cli\codigo\Modulo_DibujarInventario.bas` — DibujarInv (FUNCIONA)
- `cli\codigo\General.bas` — SwitchMap, DirGraficos, InitTileEngine call, Opciones init
- `cli\codigo\Declares.bas` — Type char, tOption, RGBClimax, E_Heading
- `cli\codigo\Mod_ErrorLOG.bas` — LogError → errores.log
