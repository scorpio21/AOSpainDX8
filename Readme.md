# AOSpainDX8 🐉
**Servidor clásico de Argentum Online — Motor DirectX 8 (VB6)**

AOSpainDX8 es la evolución del histórico servidor **AoSpain** (2003–2014), migrado de DirectX 7 a **DirectX 8** para compatibilidad moderna. Este repositorio contiene tanto el **cliente** como el **servidor** del juego, completamente en Visual Basic 6.

---

## 🎯 Objetivo del proyecto

- Migrar el motor gráfico y de audio de **DX7 → DX8** manteniendo toda la lógica de juego original.
- Preservar el **protocolo cliente↔servidor** sin cambios (valores numéricos de skills, clases, direcciones, etc.).
- Crear una base limpia, documentada y compilable para desarrollo futuro.
- Modernizar el stack sin romper la esencia del AO clásico.

---

## ✅ Estado de la migración DX7 → DX8

| Componente | Estado |
|---|---|
| `TileEngine.bas` — Motor gráfico | ✅ Migrado a DX8 (`Direct3DDevice8`) |
| `clsSoundEngine.cls` — Motor de audio | ✅ Integrado (`DirectSound8` + `DirectMusic8`) |
| `Mod_Wav.bas` — Efectos de sonido | ✅ Reescrito con `sndPlaySound` (Win32) |
| `MoDuLo_MIDI.bas` — Música | ✅ Delegado a `clsSoundEngine` |
| `clsBufferMan.cls` — Buffer de sonido | ✅ DX8 (`DirectSoundSecondaryBuffer8`) |
| `modRenderValue.bas` — Render de daño | ✅ Integrado |
| `client.vbp` — Referencias | ✅ `DX8VB.DLL` + `quartz.dll` |
| `Declares.bas` — Tipos y enums | ✅ `E_Heading`, `eSkill`, `eClass`, `tOption`, etc. |
| Protocolo cliente↔servidor | ✅ Preservado (valores numéricos originales) |
| Compilación limpia | 🔄 En progreso |

---

## 📁 Estructura del repositorio

```
AOSpainDX8/
├── cli/                    # Cliente DX8 (Visual Basic 6)
│   ├── client.vbp          # Proyecto VB6 del cliente
│   ├── codigo/             # Fuentes del cliente
│   │   ├── TileEngine.bas  # Motor gráfico DX8
│   │   ├── clsSoundEngine.cls
│   │   ├── Declares.bas
│   │   ├── General.bas
│   │   └── ...
│   ├── DX8VB.DLL           # Runtime DirectX 8 para VB6
│   ├── quartz.dll          # DirectShow (audio MP3)
│   └── zlib.dll            # Compresión
│
├── Servidor/               # Servidor de juego (Visual Basic 6)
│   ├── SERVER.VBP
│   └── Codigo/
│       └── Modulos/
│           ├── Declares.bas
│           ├── TCP.bas
│           ├── GameLogic.bas
│           └── ...
│
├── Readme.md
└── License.txt
```

---

## � Requisitos para compilar

### Cliente
- **Visual Basic 6** (SP6)
- `DX8VB.DLL` en la carpeta `cli/` (incluido)
- `quartz.dll` en la carpeta `cli/` (incluido)

### Servidor
- **Visual Basic 6** (SP6)

---

## 🎮 Clases de personaje (protocolo)

| # | Clase | # | Clase |
|---|---|---|---|
| 1 | Mago | 11 | Pescador |
| 2 | Clérigo | 12 | Herrero |
| 3 | Guerrero | 13 | Leñador |
| 4 | Asesino | 14 | Minero |
| 5 | Ladrón | 15 | Carpintero |
| 6 | Bardo | 16 | Pirata |
| 7 | Druida | 17 | Gladiador |
| 8 | Bandido | 18 | Arquero |
| 9 | Paladín | 19 | Chamán |
| 10 | Cazador | 21 | Aldeano |

---

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Puedes aportar:
- Correcciones de bugs
- Mejoras de código
- Documentación
- Herramientas auxiliares

Abre un **Issue** o un **Pull Request**.

---

## 📄 Licencia

Este proyecto se distribuye bajo licencia **GPL v2** (heredada del código original de Argentum Online).  
Copyright (C) 2002 Marquez Pablo Ignacio, Otto Perez, Aaron Perkins.

---

## 📬 Contacto

Repositorio: [https://github.com/scorpio21/AOSpainDX8](https://github.com/scorpio21/AOSpainDX8)

deepwiki: [https://deepwiki.com/scorpio21/AOSpainDX8](https://deepwiki.com/scorpio21/AOSpainDX8)
