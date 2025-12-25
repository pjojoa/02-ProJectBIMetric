'use strict';

/**
 * Módulo de animación de carga BIMETRYC
 * Convierte el componente React a TypeScript vanilla para uso en Power BI Visual
 */

/**
 * Genera un ID único para evitar conflictos en el SVG
 */
function generateUniqueId(): string {
    // Usar crypto para generar IDs seguros
    const array = new Uint32Array(2);
    if (typeof window !== 'undefined' && window.crypto && window.crypto.getRandomValues) {
        window.crypto.getRandomValues(array);
        return array[0].toString(36) + array[1].toString(36);
    }
    // Fallback para entornos sin crypto (no debería ocurrir en navegadores modernos)
    // eslint-disable-next-line powerbi-visuals/insecure-random
    return Date.now().toString(36) + Math.random().toString(36).substr(2, 5);
}

/**
 * Crea el SVG del logo BIMETRYC
 * @param width - Ancho del logo en píxeles
 * @param showWordmark - Mostrar texto "BIMETRYC" (no usado en la animación, siempre false)
 */
// eslint-disable-next-line max-lines-per-function
function createBimetycLogoSVG(width: number = 768, showWordmark: boolean = false): string {
    const uniqueId = generateUniqueId();
    const svgWidth = showWordmark ? 800 : 256;
    const svgHeight = 256;
    const aspectRatio = svgWidth / svgHeight;
    const height = width / aspectRatio;

    return `
        <svg
            xmlns="http://www.w3.org/2000/svg"
            viewBox="0 0 ${svgWidth} ${svgHeight}"
            width="${width}"
            height="${height}"
            role="img"
            aria-labelledby="title-${uniqueId} desc-${uniqueId}"
            style="display: block;"
        >
            <title id="title-${uniqueId}">BIMETRYC — cubo con letras PJ estilo monoline redondeado</title>
            <desc id="desc-${uniqueId}">Letras PJ con trazo uniforme y esquinas redondeadas, degradado cian-azul y glow suave; integradas en la cara derecha del cubo.</desc>

            <defs>
                <!-- Gradientes del cubo -->
                <linearGradient id="gTop-${uniqueId}" x1="0" y1="1" x2="1" y2="0">
                    <stop offset="0" stop-color="#22D3EE"/>
                    <stop offset="1" stop-color="#7C3AED"/>
                </linearGradient>
                <linearGradient id="gLeft-${uniqueId}" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0" stop-color="#22D3EE" stop-opacity=".9"/>
                    <stop offset="1" stop-color="#22D3EE" stop-opacity=".25"/>
                </linearGradient>
                <linearGradient id="gRight-${uniqueId}" x1="1" y1="0" x2="0" y2="1">
                    <stop offset="0" stop-color="#7C3AED" stop-opacity=".9"/>
                    <stop offset="1" stop-color="#7C3AED" stop-opacity=".25"/>
                </linearGradient>
                <linearGradient id="gEdge-${uniqueId}" x1="0" y1="0" x2="1" y2="1">
                    <stop offset="0" stop-color="#22D3EE"/>
                    <stop offset="1" stop-color="#7C3AED"/>
                </linearGradient>
                <radialGradient id="gGlow-${uniqueId}" cx="128" cy="128" r="96" gradientUnits="userSpaceOnUse">
                    <stop offset="0" stop-color="#22D3EE" stop-opacity=".16"/>
                    <stop offset="1" stop-color="#7C3AED" stop-opacity="0"/>
                </radialGradient>

                <!-- Gradiente de las letras (cian→azul) -->
                <linearGradient id="gLetter-${uniqueId}" x1="0" y1="0" x2="1" y2="1">
                    <stop offset="0" stop-color="#22D3EE"/>
                    <stop offset="1" stop-color="#1D4ED8"/>
                </linearGradient>

                <!-- Glow suave para letras -->
                <filter id="fLetterGlow-${uniqueId}" x="-50%" y="-50%" width="200%" height="200%">
                    <feGaussianBlur in="SourceGraphic" stdDeviation="0.9" result="b"/>
                    <feMerge>
                        <feMergeNode in="b"/>
                        <feMergeNode in="SourceGraphic"/>
                    </feMerge>
                </filter>

                <!-- Gradiente para capas de datos -->
                <linearGradient id="gData-${uniqueId}" x1="0" y1="0" x2="1" y2="0">
                    <stop offset="0" stop-color="#10B981"/>
                    <stop offset="1" stop-color="#06B6D4"/>
                </linearGradient>

                <!-- Recortes de caras -->
                <clipPath id="clipLeftFace-${uniqueId}" clipPathUnits="userSpaceOnUse">
                    <polygon points="64,88 128,120 128,192 64,160"/>
                </clipPath>
                <clipPath id="clipRightFace-${uniqueId}" clipPathUnits="userSpaceOnUse">
                    <polygon points="128,120 192,88 192,160 128,192"/>
                </clipPath>
            </defs>

            <!-- ICONO -->
            <g transform="translate(128,128) scale(1.44) translate(-128,-128)">
                <!-- Glow -->
                <circle cx="128" cy="128" r="92" fill="url(#gGlow-${uniqueId})"/>

                <!-- Caras del cubo -->
                <polygon points="128,56 192,88 128,120 64,88" fill="url(#gTop-${uniqueId})" opacity=".98"/>
                <polygon points="64,88 128,120 128,192 64,160" fill="url(#gLeft-${uniqueId})"/>
                <polygon points="192,88 128,120 128,192 192,160" fill="url(#gRight-${uniqueId})"/>

                <!-- Capas de datos (cara izquierda) — con grid + sparkline + puntos -->
                <g clip-path="url(#clipLeftFace-${uniqueId})">
                    <g transform="matrix(0.64,0.32,0,0.72,64,88)">
                        <!-- Estratos -->
                        <rect x="6" y="12" width="88" height="10" rx="4" fill="url(#gData-${uniqueId})" opacity=".95"/>
                        <rect x="6" y="30" width="74" height="10" rx="4" fill="url(#gData-${uniqueId})" opacity=".88"/>
                        <rect x="6" y="48" width="92" height="10" rx="4" fill="url(#gData-${uniqueId})" opacity=".90"/>
                        <rect x="6" y="66" width="80" height="10" rx="4" fill="url(#gData-${uniqueId})" opacity=".84"/>
                        <!-- Grid -->
                        <g stroke="#ffffff" stroke-opacity=".35" stroke-width="1">
                            <line x1="6" y1="24" x2="98" y2="24"/>
                            <line x1="6" y1="42" x2="98" y2="42"/>
                            <line x1="6" y1="60" x2="98" y2="60"/>
                            <line x1="6" y1="78" x2="98" y2="78"/>
                            <line x1="30" y1="8" x2="30" y2="86"/>
                            <line x1="60" y1="8" x2="60" y2="86"/>
                            <line x1="90" y1="8" x2="90" y2="86"/>
                        </g>
                        <!-- Sparkline + puntos -->
                        <path d="M8 22 C 22 14, 36 28, 48 22 S 72 26, 94 16"
                              fill="none" stroke="#0b1" stroke-opacity=".55" stroke-width="2"/>
                        <g fill="#ffffff" opacity=".9" stroke="#0b1" stroke-opacity=".4" stroke-width="1">
                            <circle cx="18" cy="20" r="1.8"/>
                            <circle cx="36" cy="26" r="1.8"/>
                            <circle cx="54" cy="21" r="1.8"/>
                            <circle cx="72" cy="25" r="1.8"/>
                            <circle cx="88" cy="18" r="1.8"/>
                        </g>
                    </g>
                </g>

                <!-- LETRAS PJ (cara derecha) con estilo de la imagen -->
                <g clip-path="url(#clipRightFace-${uniqueId})">
                    <g transform="matrix(0.64,-0.32,0,0.72,128,120)" filter="url(#fLetterGlow-${uniqueId})">
                        <!-- Monoline redondeado (ancho uniforme) -->
                        <g fill="none" stroke="url(#gLetter-${uniqueId})" stroke-width="12" stroke-linecap="round" stroke-linejoin="round" style="vector-effect: non-scaling-stroke;">
                            <!-- P -->
                            <path d="M12 12 L12 90"/>
                            <path d="M12 12 H46 Q78 12 78 34 Q78 52 46 52 H12"/>
                            <!-- J -->
                            <path d="M86 12 L86 66 Q86 90 58 90 H40"/>
                        </g>
                    </g>
                </g>

                <!-- Vértices -->
                <g fill="url(#gEdge-${uniqueId})">
                    <circle cx="128" cy="56" r="3"/>
                    <circle cx="64" cy="88" r="3"/>
                    <circle cx="192" cy="88" r="3"/>
                    <circle cx="128" cy="120" r="3"/>
                    <circle cx="64" cy="160" r="3"/>
                    <circle cx="192" cy="160" r="3"/>
                    <circle cx="128" cy="192" r="3"/>
                </g>

                <!-- Brillo superior -->
                <path d="M64 88 L128 60 L192 88 Q128 114 64 88 Z" fill="#fff" opacity=".08"/>
            </g>
        </svg>
    `;
}

/**
 * Inyecta los keyframes CSS necesarios para las animaciones
 */
function injectAnimationStyles(): void {
    // Verificar si ya se inyectaron los estilos
    if (document.getElementById('bimetyc-animation-styles')) {
        return;
    }

    const style = document.createElement('style');
    style.id = 'bimetyc-animation-styles';
    style.textContent = `
        @keyframes logoPulseRing {
            0% { transform: scale(0.8); opacity: 0.25; }
            50% { transform: scale(1); opacity: 0.8; }
            100% { transform: scale(1.18); opacity: 0; }
        }
        @keyframes logoFloat {
            0% { transform: translateY(0px); }
            50% { transform: translateY(-10px); }
            100% { transform: translateY(0px); }
        }
        @keyframes logoFadeIn {
            from { opacity: 0; transform: translateY(12px); }
            to { opacity: 1; transform: translateY(0); }
        }
        @keyframes logoSweep {
            0% { transform: translateX(-100%); }
            100% { transform: translateX(100%); }
        }
        @keyframes logoScan {
            0% { transform: translateY(-60px); opacity: 0; }
            30% { opacity: 0.4; }
            70% { opacity: 0.4; }
            100% { transform: translateY(60px); opacity: 0; }
        }
    `;
    document.head.appendChild(style);
}

/**
 * Crea y muestra la animación de carga BIMETRYC
 * @param container - Contenedor donde se mostrará la animación
 * @param duration - Duración de la animación en milisegundos (default: 4000)
 * @param loadingText - Texto de carga (default: "Preparando tu modelo...")
 * @param companyName - Nombre de la empresa (default: "SKYDATABIM S.A.S.")
 * @returns Función para ocultar la animación manualmente
 */
// eslint-disable-next-line max-lines-per-function
export function showLoadingAnimation(
    container: HTMLElement,
    duration: number = 4000,
    loadingText: string = "Preparando tu modelo...",
    companyName: string = "SKYDATABIM S.A.S."
): () => void {
    // Inyectar estilos CSS si no están ya inyectados
    injectAnimationStyles();

    // Crear overlay principal
    const overlay = document.createElement('div');
    overlay.id = 'bimetyc-loading-overlay';
    overlay.style.cssText = `
        position: absolute;
        top: 0;
        left: 0;
        right: 0;
        bottom: 0;
        background: radial-gradient(circle at 30% 30%, rgba(0, 150, 200, 0.22), transparent 55%), radial-gradient(circle at 70% 70%, rgba(0, 255, 170, 0.18), transparent 60%), #02070c;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        z-index: 1200;
        color: #fff;
        pointer-events: none;
        overflow: hidden;
    `;

    // Grid Overlay
    const gridOverlay = document.createElement('div');
    gridOverlay.style.cssText = `
        position: absolute;
        inset: 0;
        background-image: linear-gradient(90deg, rgba(0, 200, 255, 0.07) 1px, transparent 1px), linear-gradient(0deg, rgba(0, 200, 255, 0.05) 1px, transparent 1px);
        background-size: 120px 120px;
        opacity: 0.35;
    `;
    overlay.appendChild(gridOverlay);

    // Sweep Overlay
    const sweepOverlay = document.createElement('div');
    sweepOverlay.style.cssText = `
        position: absolute;
        inset: 0;
        background: linear-gradient(120deg, rgba(0, 0, 0, 0) 40%, rgba(0, 200, 255, 0.05), rgba(0, 0, 0, 0) 60%);
        mix-blend-mode: screen;
        animation: logoSweep 2.2s linear infinite;
    `;
    overlay.appendChild(sweepOverlay);

    // Contenedor del Logo con Anillos y Escaneo
    // Reducido al 40% del tamaño original (reducción del 60%)
    const logoContainer = document.createElement('div');
    logoContainer.style.cssText = `
        position: relative;
        width: 168px;
        height: 168px;
        margin-bottom: 20px;
    `;

    // Anillo Pulsante Exterior
    const outerRing = document.createElement('div');
    outerRing.style.cssText = `
        position: absolute;
        inset: 0;
        border-radius: 50%;
        border: 2px solid rgba(0, 200, 255, 0.45);
        animation: logoPulseRing 1.6s ease-out infinite;
    `;
    logoContainer.appendChild(outerRing);

    // Anillo Pulsante Interior (ajustado proporcionalmente)
    const innerRing = document.createElement('div');
    innerRing.style.cssText = `
        position: absolute;
        inset: 16px;
        border-radius: 50%;
        border: 2px solid rgba(0, 255, 170, 0.35);
        animation: logoPulseRing 1.9s ease-out infinite 0.2s;
    `;
    logoContainer.appendChild(innerRing);

    // Línea de Escaneo (ajustada proporcionalmente)
    const scanLine = document.createElement('div');
    scanLine.style.cssText = `
        position: absolute;
        left: 20px;
        right: 20px;
        height: 12px;
        background: linear-gradient(90deg, transparent, rgba(0, 255, 170, 0.2), transparent);
        filter: blur(2px);
        animation: logoScan 1.2s ease-in-out infinite;
    `;
    logoContainer.appendChild(scanLine);

    // Logo SVG Centrado
    const logoWrapper = document.createElement('div');
    logoWrapper.style.cssText = `
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        z-index: 2;
    `;
    const logoFloat = document.createElement('div');
    logoFloat.style.cssText = `
        animation: logoFloat 2s ease-in-out infinite;
    `;
    // eslint-disable-next-line powerbi-visuals/no-inner-outer-html
    // Logo reducido al 40%: 768 * 0.4 = 307px
    logoFloat.innerHTML = createBimetycLogoSVG(307, false);
    logoWrapper.appendChild(logoFloat);
    logoContainer.appendChild(logoWrapper);

    overlay.appendChild(logoContainer);

    // Texto (reducido proporcionalmente al 40%)
    const textContainer = document.createElement('div');
    textContainer.style.cssText = `
        text-align: center;
        animation: logoFadeIn 0.5s ease forwards;
        padding: 0 10px;
        margin-top: 20px;
    `;

    const titleText = document.createElement('div');
    titleText.style.cssText = `
        font-size: 45px;
        letter-spacing: 6px;
        font-weight: 800;
        color: #00c8ff;
        margin-bottom: 2px;
    `;
    titleText.textContent = 'BIMETRYC';
    textContainer.appendChild(titleText);

    const companyText = document.createElement('div');
    companyText.style.cssText = `
        font-size: 9px;
        letter-spacing: 3px;
        color: rgba(255, 255, 255, 0.65);
        margin-bottom: 4px;
        margin-top: 0px;
    `;
    companyText.textContent = companyName;
    textContainer.appendChild(companyText);

    const loadingTextEl = document.createElement('div');
    loadingTextEl.style.cssText = `
        font-size: 5px;
        letter-spacing: 1px;
        color: rgba(255, 255, 255, 0.75);
    `;
    loadingTextEl.textContent = loadingText;
    textContainer.appendChild(loadingTextEl);

    overlay.appendChild(textContainer);

    // Asegurar que el contenedor tenga position relative
    const containerPosition = window.getComputedStyle(container).position;
    if (containerPosition === 'static') {
        container.style.position = 'relative';
    }

    // Agregar overlay al contenedor
    container.appendChild(overlay);

    // Función para ocultar la animación
    const hideAnimation = (): void => {
        if (overlay && overlay.parentNode) {
            overlay.parentNode.removeChild(overlay);
        }
    };

    // Ocultar automáticamente después de la duración especificada
    const timeoutId = setTimeout(() => {
        hideAnimation();
    }, duration);

    // Retornar función para ocultar manualmente (que también limpia el timeout)
    return (): void => {
        clearTimeout(timeoutId);
        hideAnimation();
    };
}

/**
 * Oculta la animación de carga si está visible
 * @param container - Contenedor donde se muestra la animación
 */
export function hideLoadingAnimation(container: HTMLElement): void {
    const overlay = container.querySelector('#bimetyc-loading-overlay');
    if (overlay && overlay.parentNode) {
        overlay.parentNode.removeChild(overlay);
    }
}
