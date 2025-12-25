# Changelog - APS Viewer Visual

## Versión 1.0.8.0 (Actual)

### ✨ Nuevas Características

#### Animación de Carga BIMETRYC
- **Nuevo módulo**: `loading-animation.ts`
  - Animación completa del logo BIMETRYC convertida de React a TypeScript vanilla
  - Efectos visuales incluidos:
    - Grid overlay animado
    - Sweep overlay (barrido de luz)
    - Anillos pulsantes (exterior e interior)
    - Línea de escaneo vertical
    - Logo SVG flotante con animación suave
    - Texto con fade-in

- **Duración**: Exactamente 4 segundos (4000ms)
- **Activación automática**:
  - Cuando se detecta un nuevo URN (cambio de modelo)
  - Cuando el visor se reinicia (cambio de página en Power BI)
  - Antes de inicializar el visor en todos los casos

#### Mejoras en Gestión del Visor
- **Nuevo método**: `destroyViewer()`
  - Limpia correctamente el visor anterior antes de crear uno nuevo
  - Maneja la limpieza de recursos (selección, theming, listeners)
  - Resetea el estado del visual correctamente
  - Previene memory leaks

- **Detección mejorada de reinicios**:
  - Detecta cuando el visor se reinicia (viewer es null pero hay URN)
  - Muestra la animación también en reinicios
  - Maneja correctamente los cambios de página en Power BI

#### Manejo de Errores Mejorado
- La animación se oculta automáticamente si:
  - Falla la obtención del token
  - Falla la inicialización del visor
- Logs mejorados para debugging

### 🔧 Cambios Técnicos

- **tsconfig.json**: Actualizado para incluir todos los archivos TypeScript en `src/`
- **pbiviz.json**: Versión actualizada a 1.0.8.0
- **Imports**: Agregado import del módulo `loading-animation`

### 📝 Archivos Modificados

1. `src/visual.ts`
   - Agregado import de `loading-animation`
   - Agregadas propiedades para gestión de animación
   - Agregado método `destroyViewer()`
   - Modificada lógica de detección de URN/reinicio
   - Modificado `initializeViewer()` para mostrar animación

2. `src/loading-animation.ts` (NUEVO)
   - Módulo completo de animación
   - Función `showLoadingAnimation()` con duración configurable
   - Función `hideLoadingAnimation()` para ocultar manualmente
   - Generación de SVG del logo con IDs únicos
   - Inyección automática de keyframes CSS

3. `tsconfig.json`
   - Cambiado de `files` a `include` para incluir todos los archivos .ts

4. `pbiviz.json`
   - Versión actualizada a 1.0.8.0

### 🐛 Correcciones

- Corrección en limpieza del contenedor al destruir el visor
- Mejor manejo de reinicios del visor al cambiar de página

---

## Versión 1.0.7.0 (Anterior)

- Versión base con funcionalidad completa del visor
- Integración con Autodesk Platform Services
- Coloreo categórico dinámico
- Filtrado bidireccional (Power BI ↔ Viewer)
- Extensión de ghosting personalizada
