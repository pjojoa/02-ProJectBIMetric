# Instrucciones para Compilar y Empaquetar el Visual

## 📋 Requisitos Previos

1. **Node.js**: Versión 14.x o superior
2. **npm**: Incluido con Node.js
3. **Power BI Visuals Tools**: Instalado globalmente
   ```bash
   npm install -g powerbi-visuals-tools
   ```

## 🔨 Pasos para Compilar y Empaquetar

### 1. Navegar a la carpeta del visual
```bash
cd "d:\03 Project BIMetric\aps-powerbi-tools-develop\visuals\aps-viewer-visual"
```

### 2. Instalar dependencias (si es la primera vez)
```bash
npm install
```

### 3. Verificar que no hay errores de compilación
```bash
npm run lint
```

### 4. Compilar y empaquetar el visual
```bash
npm run package
```

Este comando:
- Compila todos los archivos TypeScript
- Genera el archivo `.pbiviz` en la carpeta `dist/`
- El archivo se llamará: `aps_viewer_visual_a4f2990a03324cf79eb44f982719df44.1.0.8.0.pbiviz`

### 5. Ubicación del archivo empaquetado
El archivo `.pbiviz` estará en:
```
dist/aps_viewer_visual_a4f2990a03324cf79eb44f982719df44.1.0.8.0.pbiviz
```

## 📦 Importar en Power BI Desktop

### Opción 1: Desde Power BI Desktop

1. Abre Power BI Desktop
2. Ve a la pestaña **Visualizaciones**
3. Haz clic en los **tres puntos** (`...`) en la parte inferior
4. Selecciona **Importar un visual desde un archivo**
5. Navega a la carpeta `dist/` y selecciona el archivo `.pbiviz`
6. El visual aparecerá en el panel de visualizaciones

### Opción 2: Desde el Editor de Visuales

1. Abre Power BI Desktop
2. Ve a **Archivo** → **Opciones y configuración** → **Opciones**
3. En la sección **Seguridad**, habilita **Desarrollador de visuales**
4. En la pestaña **Visualizaciones**, aparecerá un icono de **Desarrollador**
5. Arrastra el visual al lienzo
6. Configura el endpoint de token en las opciones de formato

## ✅ Verificación Post-Instalación

### Probar la Animación de Carga

1. **Carga inicial**:
   - Agrega el visual a un reporte
   - Configura el endpoint de token
   - Arrastra una columna con URN al campo "Urn"
   - Deberías ver la animación BIMETRYC por 4 segundos

2. **Cambio de URN**:
   - Cambia el URN en los datos
   - La animación debería aparecer nuevamente

3. **Cambio de página**:
   - Cambia de página en el reporte de Power BI
   - El visor se reiniciará y la animación debería aparecer

### Configuración Requerida

1. **Access Token Endpoint**: 
   - Ve a **Formato** → **Viewer Runtime**
   - Configura la URL del endpoint (ej: `https://zero2-projectbimetric.onrender.com/token`)

2. **Campos de Datos**:
   - **Urn**: Columna con el URN del modelo
   - **ExternalIds**: Columna con los IDs externos de los elementos
   - **Color** (opcional): Columna con categorías para coloreo

## 🐛 Solución de Problemas

### Error: "Cannot find module 'loading-animation'"
- **Solución**: Verifica que el archivo `src/loading-animation.ts` existe
- Verifica que `tsconfig.json` incluye todos los archivos en `src/`

### La animación no aparece
- **Verifica**: Que el contenedor tenga `position: relative`
- **Verifica**: Que no haya errores en la consola del navegador (F12)
- **Verifica**: Que el URN esté correctamente configurado

### El visual no compila
- **Solución**: Ejecuta `npm install` nuevamente
- **Solución**: Verifica que todas las dependencias estén instaladas
- **Solución**: Ejecuta `npm run lint` para ver errores específicos

### El archivo .pbiviz no se genera
- **Solución**: Verifica que no haya errores de TypeScript
- **Solución**: Verifica que `pbiviz` esté instalado globalmente
- **Solución**: Intenta ejecutar `pbiviz package` directamente

## 📝 Notas Importantes

- **Versión**: El visual está en la versión **1.0.8.0**
- **Duración de animación**: Exactamente **4 segundos** (4000ms)
- **Compatibilidad**: Power BI API 5.4.0
- **Navegadores**: Compatible con Chrome, Edge, Firefox (versiones recientes)

## 🔄 Actualizar el Visual Existente

Si ya tienes el visual importado en Power BI:

1. Elimina el visual anterior del reporte
2. Importa la nueva versión (1.0.8.0)
3. Reconfigura los campos y opciones

O simplemente reimporta el archivo `.pbiviz` - Power BI actualizará el visual automáticamente.
