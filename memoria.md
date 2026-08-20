# Memoria del Proyecto - AospainOri

## Objective
Resolver el problema de texturas incorrectas en el inventario del cliente AO. Los objetos newbie (manzana roja, botella de agua, daga, vestimentas, pociones) muestran texturas incorrectas (alas) en lugar de las correctas.

**Directorio de trabajo:** `I:\AospainOri`

## Compile/verify commands
- **Compilación:** Abrir el proyecto VB6 en `I:\AospainOri\cli\` y compilar (F5 o Make Project)
- **Ejecución:** Ejecutar el cliente, iniciar sesión, abrir inventario (F1)
- **Verificación:** Observar los objetos newbie en slots 1-6 del inventario
- **Logs:** Los logs se generan en `I:\AospainOri\cli\Logs\`
  - `errores.log` - errores generales
  - `Inventario.log` - datos del inventario
  - `DrawGrhtoHdc.log` - detalles de renderizado GRH

## Current state

### ✅ HECHO (Verified)
- **Datos técnicos correctos:** 
  - `OBJ.DAT` (servidor): GrhIndex correctos para OBJ 467,468,460,463,461,462
  - `Graficos.ind` (cliente): Coordenadas correctas confirmadas con fix_grh_ind.py
  - `3.bmp` (cliente): Archivo correcto y completo (640x742, 8-bit con paleta)
- **AO GRH Editor:** Confirma que las imágenes están en las coordenadas correctas en 3.bmp
- **Logging:** Sistema de logging agregado en `DrawGrhtoHdc` y `Modulo_DibujarInventario`
- **Identificación del problema:** El problema está en `DrawGrhtoHdc` en `CargasInit.bas` línea 302-314

### ⚠️ ACTIVO (In Progress)
- **Investigación del renderizado BMP:** `DrawGrhtoHdc` usa `StretchDIBits` para BMPs de 8-bit
- **Cálculo de offset:** Coordenadas (sx,sy) se calculan correctamente pero no se usan correctamente en el buffer

### ❌ BLOQUEADO (Blocked)
- **Corrección de offset causó crash:** Usar `Buffer(pixelOffset)` en lugar de `Buffer(PixelDataOffset)` causó que el cliente se cerrara
- **Modificación de BITMAPINFOHEADER falló:** Cambiar biWidth/biHeight causó que `StretchDIBits` retornara 0 (sin renderizado)

## Work history
- **BASE:** Código original sin modificaciones
- **Cambios en `CargasInit.bas`:**
  - Línea 270-291: Agregado logging detallado con cálculo de rowSize y pixelOffset
  - Línea 245-261: Logging de BMP (header, paleta)
  - Línea 302-314: Intentos de corrección (revertidos)
- **Cambios en `Modulo_DibujarInventario.bas`:**
  - Línea 123-127: Logging de coordenadas de origen/destino
- **Cambios en `Mod_ErrorLOG.bas`:**
  - Línea 127: Habilitado Log_Inventario
- **Commit ca18dc7:** "Add logging system for inventory texture debugging" (2026-08-21)

## Git configuration
- **Email para commits:** 7824081+scorpio21@users.noreply.github.com
- **Repo:** https://github.com/scorpio21/AOSpainDX8.git

## Next steps
1. **Enfoque alternativo para renderizado BMP de 8-bit:**
   - Investigar si hay que usar funciones GDI diferentes para BMPs con paleta
   - Considerar extracción manual de píxeles de la región específica
   - Revisar si hay ejemplos de código similar en el proyecto (DX8 vs GDI32)

2. **Investigar sistemas alternativos:**
   - Revisar `clsGraphicalInventory.cls` (DirectX 8) como posible solución
   - Comparar cómo renderiza DX8 vs GDI32

3. **Solución temporal (si es posible):**
   - Considerar reconstruir/reemplazar el archivo 3.bmp con uno que tenga las coordenadas esperadas por el código actual

## Known gotchas
- **Archivo 3.bmp es correcto:** Confirmado visualmente que tiene las imágenes correctas en las coordenadas correctas
- **BMP es 8-bit con paleta:** BitsPerPixel=8, Compression=0, ColorsUsed=255
- **rowSize=640:** No hay padding en las filas (640 es múltiplo de 4)
- **CalcOffset vs PixelDataOffset:** 
  - PixelDataOffset=1074 (inicio de datos de píxeles)
  - CalcOffset=22130 para GRH 506 (posición correcta en el buffer)
  - Usar CalcOffset causó crash del cliente
- **StretchDIBits:**
  - Con código original: result=742 (funciona pero texturas incorrectas)
  - Con BITMAPINFOHEADER modificado: result=0 (sin renderizado)
- **Dos sistemas de inventario:**
  - `Modulo_DibujarInventario.bas` + `DrawGrhtoHdc` (GDI32) - SISTEMA ACTIVO
  - `clsGraphicalInventory.cls` (DirectX 8) - SISTEMA INACTIVO
- **Archivo fix_grh_ind.py:** Script útil para verificar coordenadas en Graficos.ind

## Key files
- `I:\AospainOri\cli\codigo\CargasInit.bas` - Función `DrawGrhtoHdc` (línea 33-318)
- `I:\AospainOri\cli\codigo\Modulo_DibujarInventario.bas` - Función `DibujarInv` (línea 1-132)
- `I:\AospainOri\cli\Init\Graficos.ind` - Índice de gráficos del cliente
- `I:\AospainOri\Servidor\Dat\OBJ.DAT` - Definición de objetos del servidor
- `I:\AospainOri\cli\Graficos\3.bmp` - Atlas de sprites de objetos (640x742, 8-bit)
- `I:\AospainOri\fix_grh_ind.py` - Script de verificación de Graficos.ind