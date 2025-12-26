/**
 * SECURE-BUILD.JS - BIMETRYC Code Protection Layer
 * 
 * Este script automatiza la protección de la propiedad intelectual del visual APS Viewer.
 * 1. Descomprime el archivo .pbiviz generado por Power BI Tools.
 * 2. Ofusca el código JavaScript compilado para hacerlo ilegible.
 * 3. Vuelve a empaquetar el visual de forma segura.
 */

const fs = require('fs');
const path = require('path');
const AdmZip = require('adm-zip');
const JavaScriptObfuscator = require('javascript-obfuscator');

const VISUAL_DIR = __dirname;
const DIST_DIR = path.join(VISUAL_DIR, 'dist');

/**
 * Encuentra el archivo .pbiviz más reciente en la carpeta dist
 */
function findLatestPbiviz() {
    const files = fs.readdirSync(DIST_DIR);
    const pbivizFiles = files
        .filter(f => f.endsWith('.pbiviz'))
        .map(f => ({ name: f, time: fs.statSync(path.join(DIST_DIR, f)).mtime.getTime() }))
        .sort((a, b) => b.time - a.time);

    return pbivizFiles.length > 0 ? path.join(DIST_DIR, pbivizFiles[0].name) : null;
}

async function secureBuild() {
    console.log('\n--- BIMETRYC SECURE BUILD ---');

    const pbivizPath = findLatestPbiviz();
    if (!pbivizPath) {
        console.error('Error: No se encontró ningún archivo .pbiviz en la carpeta dist/');
        process.exit(1);
    }

    console.log(`Protegiendo: ${path.basename(pbivizPath)}`);

    try {
        const zip = new AdmZip(pbivizPath);
        const zipEntries = zip.getEntries();

        // Buscamos los archivos JavaScript en la subcarpeta resources/
        let modified = false;

        for (const entry of zipEntries) {
            console.log(`  🔍 Revisando entrada: ${entry.entryName}`);
            // En versiones modernas, el JS está embebido dentro del archivo .pbiviz.json en resources/
            if (entry.entryName.startsWith('resources/') && entry.entryName.endsWith('.pbiviz.json')) {
                console.log(`  -> Procesando recurso: ${entry.entryName}`);

                let jsonContent;
                try {
                    jsonContent = JSON.parse(entry.getData().toString('utf8'));
                    console.log(`  -> Claves en JSON: ${Object.keys(jsonContent).join(', ')}`);
                    if (jsonContent.content) {
                        console.log(`  -> Claves en content: ${Object.keys(jsonContent.content).join(', ')}`);
                    }
                    if (jsonContent.externalJS) {
                        console.log(`  -> externalJS: ${JSON.stringify(jsonContent.externalJS)}`);
                    }
                    if (jsonContent.visualEntryPoint) {
                        console.log(`  -> visualEntryPoint: ${jsonContent.visualEntryPoint}`);
                    }
                } catch (e) {
                    console.error(`Error al parsear ${entry.entryName}:`, e);
                    continue;
                }

                if (jsonContent.content && jsonContent.content.js) {
                    console.log(`  -> ✅ Encontrado código en 'content.js' (${jsonContent.content.js.substring(0, 50)}...)`);
                    console.log(`  -> Ofuscando contenido...`);

                    const originalCode = jsonContent.content.js;

                    const obfuscationResult = JavaScriptObfuscator.obfuscate(originalCode, {
                        compact: true,
                        controlFlowFlattening: true,
                        controlFlowFlatteningThreshold: 1.0,
                        numbersToExpressions: true,
                        simplify: true,
                        stringArray: true,
                        stringArrayEncoding: ['base64'],
                        stringArrayThreshold: 1.0,
                        splitStrings: true,
                        splitStringsChunkLength: 5,
                        unicodeEscapeSequence: false,
                        renameGlobals: false,
                        selfDefending: false,
                        debugProtection: false
                    });

                    jsonContent.content.js = obfuscationResult.getObfuscatedCode();
                    zip.updateFile(entry.entryName, Buffer.from(JSON.stringify(jsonContent), 'utf8'));
                    modified = true;
                }
            }

            // Caso alternativo: archivo .js directo (versiones antiguas o assets)
            if (entry.entryName.startsWith('resources/') && entry.entryName.endsWith('.js')) {
                console.log(`  -> Ofuscando archivo directo: ${entry.entryName}`);
                const originalCode = entry.getData().toString('utf8');
                const obfuscationResult = JavaScriptObfuscator.obfuscate(originalCode, {
                    compact: true,
                    controlFlowFlattening: true,
                    renameGlobals: false
                });
                zip.updateFile(entry.entryName, Buffer.from(obfuscationResult.getObfuscatedCode(), 'utf8'));
                modified = true;
            }
        }

        if (modified) {
            zip.writeZip(pbivizPath);
            console.log('\n--- PROTECCIÓN COMPLETADA CON ÉXITO ---');
            console.log('El archivo .pbiviz ahora está protegido contra ingeniería inversa.');
        } else {
            console.warn('Advertencia: No se encontraron archivos JavaScript para proteger en el paquete.');
        }

    } catch (error) {
        console.error('Error durante el proceso de protección:', error);
        process.exit(1);
    }
}

secureBuild();
