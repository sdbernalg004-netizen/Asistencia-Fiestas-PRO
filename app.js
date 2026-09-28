/**
 * ============================================================================
 * LÓGICA PRINCIPAL DE LA APLICACIÓN - QR EVENT CHECK-IN PRO (app.js)
 * ============================================================================
 * Estructura del código:
 * 1. CONFIGURACIÓN & FIREBASE: URL de Apps Script y sincronización en tiempo real.
 * 2. CONTROL DE LICENCIAS: Verificación multidispositivo y candado de activación.
 * 3. SINTETIZADOR DE SONIDOS: Feedback auditivo al escanear (Éxito / Duplicado / Error).
 * 4. HERRAMIENTAS PRO & BRANDING: Cambio dinámico de colores, logos y título del evento.
 * 5. LECTOR DE EXCEL & LÍMITES: Carga de lista de invitados con límites dinámicos de plan.
 * 6. TABLA Y BÚSQUEDA: Filtros de asistieron/pendientes y barra de búsqueda en tiempo real.
 * 7. ESCÁNER DE CÁMARA QR: Motor de cámara nativa con aceleración por hardware.
 * 8. GENERADOR MASIVO DE BOLETOS QR: Creación y descarga en archivo comprimido (.ZIP).
 * 9. FUSIÓN MULTIPUERTA OFFLINE: Consolidador de Excels de múltiples puertas.
 */

// ----------------------------------------------------------------------------
// 1. CONFIGURACIÓN Y SERVIDORES (GOOGLE SHEETS Y FIREBASE)
// ----------------------------------------------------------------------------

// URL del Web App de Google Apps Script para validar licencias en la nube
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw7LhIVbvb1H9fc0YgeuWjZWGzb3rBikUa52pE8zu7K8JKA1BQgDSFj9c5kZ3XBiijFog/exec";

// Configuración de la base de datos en tiempo real de Firebase (Realtime Database)
const firebaseConfig = {
  apiKey: "AIzaSyBgUmpivXlg8-6-pYaJOfReTLmWYeZ8GQ8",
  authDomain: "acceso-qr-pro.firebaseapp.com",
  databaseURL: "https://acceso-qr-pro-default-rtdb.firebaseio.com",
  projectId: "acceso-qr-pro",
  storageBucket: "acceso-qr-pro.firebasestorage.app",
  messagingSenderId: "912113755513",
  appId: "1:912113755513:web:8dcf14636a5438bf28fba8",
  measurementId: "G-WQXM1XX8NW"
};

/**
 * Sanitizar claves para Firebase (reemplaza caracteres no permitidos en JSON)
 */
function sanitizeFirebaseKey(val) {
    if (!val) return "default";
    return String(val).replace(/[\.\#\$\[\]]/g, "_");
}

// Inicializar Firebase si las credenciales son válidas
const isFirebaseActive = typeof firebase !== 'undefined' && firebaseConfig.databaseURL && !firebaseConfig.databaseURL.includes("TU_PROYECTO");
let dbRef = null;

if (isFirebaseActive) {
    firebase.initializeApp(firebaseConfig);
}

/**
 * Enviar un registro de asistencia en tiempo real a Firebase Realtime Database
 */
function syncCheckInToFirebase(guestId, attended, time, door = "Local") {
    if (!isFirebaseActive) return;
    const licenseKey = secureGetKey("pro-active-license-key") || "TEST-123-KEY";
    const cleanKey = sanitizeFirebaseKey(licenseKey);
    const cleanEvent = sanitizeFirebaseKey(activeSheetName);
    const cleanGuestId = sanitizeFirebaseKey(guestId);
    
    firebase.database().ref(`licencias/${cleanKey}/eventos/${cleanEvent}/asistencias/${cleanGuestId}`).set({
        attended: attended,
        attendTime: time,
        door: door
    });
}

// Global state variables
let guestData = []; // Array of guest objects
let originalWorkbook = null; 
let activeSheetName = ""; 
let html5QrScanner = null; 
let cameraList = []; 
let isScannerActive = false; 
let currentFilter = "all"; 
let originalColumnsOrder = []; 

// Merge variables
let mergeFilesData = []; // Array of parsed excel workbooks/rows for merging

// DOM Elements
const dropZone = document.getElementById("drop-zone");
const excelFileInput = document.getElementById("excel-file-input");
const fileInfo = document.getElementById("file-info");
const fileNameDisplay = document.getElementById("file-name-display");
const resetFileBtn = document.getElementById("reset-file-btn");
const fileStatusIndicator = document.getElementById("file-status-indicator");

const statTotal = document.getElementById("stat-total");
const statAttended = document.getElementById("stat-attended");
const statPending = document.getElementById("stat-pending");
const progressFill = document.getElementById("progress-fill");
const progressPercentage = document.getElementById("progress-percentage");

const cameraSelect = document.getElementById("camera-select");
const toggleCameraBtn = document.getElementById("toggle-camera-btn");
const scannerOverlay = document.getElementById("scanner-overlay");

const resultCard = document.getElementById("result-card");
const resultPlaceholder = document.getElementById("result-placeholder");
const resultDetails = document.getElementById("result-details");
const resultIcon = document.getElementById("result-icon");
const resultTitle = document.getElementById("result-title");
const resultName = document.getElementById("result-name");
const resultId = document.getElementById("result-id");
const resultQty = document.getElementById("result-qty");
const resultTime = document.getElementById("result-time");
const resultStatusIconWrapper = document.getElementById("result-status-icon-wrapper");

const searchInput = document.getElementById("search-input");
const filterButtons = document.querySelectorAll(".btn-filter");
const guestTableBody = document.getElementById("guest-table-body");
const downloadExcelBtn = document.getElementById("download-excel-btn");

// PRO Tab & Branding elements
const tabButtons = document.querySelectorAll(".tab-btn");
const tabContents = document.querySelectorAll(".tab-content");

const themeColorPicker = document.getElementById("theme-color-picker");
const themeColorCode = document.getElementById("theme-color-code");
const logoFileInput = document.getElementById("logo-file-input");
const logoFileName = document.getElementById("logo-file-name");
const appTitleInput = document.getElementById("app-title-input");
const saveBrandingBtn = document.getElementById("save-branding-btn");
const resetBrandingBtn = document.getElementById("reset-branding-btn");

const customLogoImg = document.getElementById("custom-logo");
const defaultLogoIcon = document.getElementById("default-logo-icon");
const appTitleDisplay = document.getElementById("app-title-display");
const appDescDisplay = document.getElementById("app-desc-display");
const themeColorMeta = document.getElementById("theme-color-meta");

// PRO QR Generator elements
const generateQrsBtn = document.getElementById("generate-qrs-btn");
const qrGenStatusBox = document.getElementById("qr-gen-status-box");
const qrGenMsg = document.getElementById("qr-gen-msg");
const qrProgressContainer = document.getElementById("qr-progress-container");
const qrProgressFill = document.getElementById("qr-progress-fill");
const qrProgressLbl = document.getElementById("qr-progress-lbl");

// PRO Merge elements
const mergeDropZone = document.getElementById("merge-drop-zone");
const mergeFilesInput = document.getElementById("merge-files-input");
const mergedFilesListWrapper = document.getElementById("merged-files-list-wrapper");
const mergedFilesList = document.getElementById("merged-files-list");
const processMergeBtn = document.getElementById("process-merge-btn");
const resetMergeBtn = document.getElementById("reset-merge-btn");

// License Lock DOM Elements
const licenseLockScreen = document.getElementById("license-lock-screen");
const activationKeyInput = document.getElementById("activation-key-input");
const activateAppBtn = document.getElementById("activate-app-btn");
const activationErrorMsg = document.getElementById("activation-error-msg");
const errorText = document.getElementById("error-text");
const syncStatusIndicator = document.getElementById("sync-status-indicator");

// New Feature State & DOM Elements
let isVoiceEnabled = localStorage.getItem("pro-voice-enabled") !== "false";
let currentLang = localStorage.getItem("pro-lang") || "es";
let isLightTheme = localStorage.getItem("pro-theme") === "light";
let isTorchOn = false;
let activeVideoTrack = null;
let checkinTimestamps = []; // { time, qty, hour }

const themeToggleBtn = document.getElementById("theme-toggle-btn");
const themeIcon = document.getElementById("theme-icon");
const themeBtnText = document.getElementById("theme-btn-text");

const voiceToggleBtn = document.getElementById("voice-toggle-btn");
const voiceIcon = document.getElementById("voice-icon");
const voiceBtnText = document.getElementById("voice-btn-text");

const langToggleBtn = document.getElementById("lang-toggle-btn");
const langBtnText = document.getElementById("lang-btn-text");

const torchBtn = document.getElementById("torch-btn");

const flowRateDisplay = document.getElementById("flow-rate-display");
const flowRateText = document.getElementById("flow-rate-text");
const openReportBtn = document.getElementById("open-report-btn");

const resultVipBanner = document.getElementById("result-vip-banner");
const resultTableBox = document.getElementById("result-table-box");
const resultTable = document.getElementById("result-table");
const resultNotesRow = document.getElementById("result-notes-row");
const resultNotes = document.getElementById("result-notes");
const partialCheckinBox = document.getElementById("partial-checkin-box");
const partialButtonsGroup = document.getElementById("partial-buttons-group");

const individualQrModal = document.getElementById("individual-qr-modal");
const closeIndivQrModalBtn = document.getElementById("close-indiv-qr-modal-btn");
const indivQrCodeBox = document.getElementById("indiv-qr-code-box");
const indivGuestName = document.getElementById("indiv-guest-name");
const indivGuestId = document.getElementById("indiv-guest-id");
const indivGuestQty = document.getElementById("indiv-guest-qty");
const indivGuestTable = document.getElementById("indiv-guest-table");
const indivWaBtn = document.getElementById("indiv-wa-btn");
const indivDownloadBtn = document.getElementById("indiv-download-btn");

const executiveReportModal = document.getElementById("executive-report-modal");
const closeRepModalBtn = document.getElementById("close-rep-modal-btn");
const repEventTitle = document.getElementById("rep-event-title");
const repTotalGuests = document.getElementById("rep-total-guests");
const repTotalAttended = document.getElementById("rep-total-attended");
const repTotalPending = document.getElementById("rep-total-pending");
const repPercentage = document.getElementById("rep-percentage");
const repPeakHour = document.getElementById("rep-peak-hour");
const repPeakRate = document.getElementById("rep-peak-rate");
const repTablesBreakdown = document.getElementById("rep-tables-breakdown");
const printRepBtn = document.getElementById("print-rep-btn");

// Dictionary for Internationalization (ES / EN)
const i18n = {
    es: {
        scannerTab: "Escáner",
        toolsTab: "Herramientas Pro",
        guideBtn: "Guía de Uso",
        loadDbTitle: "Cargar Base de Datos",
        loadDbDesc: "Carga tu Excel original de invitados para iniciar el control de asistencia.",
        dropLabel: "Arrastra tu Excel aquí o haz clic para seleccionar",
        downloadTplBtn: "Descargar Plantilla Excel de Ejemplo",
        statsTitle: "Estadísticas en Tiempo Real",
        statGuests: "Invitados",
        statAttended: "Asistieron",
        statPending: "Pendientes",
        progressLbl: "Progreso del evento",
        reportBtn: "Reporte Ejecutivo",
        downloadExcelBtn: "Descargar Excel Actualizado",
        downloadHint: "Descarga para actualizar los registros de esta puerta de acceso.",
        scannerTitle: "Escáner QR",
        startScannerBtn: "Iniciar Escáner",
        stopScannerBtn: "Detener Escáner",
        waitingScan: "Esperando escaneo de código QR...",
        vipBadge: "INVITADO VIP / ESPECIAL",
        metaPasses: "Pases:",
        metaTime: "Hora:",
        partialLabel: "¿Entrada parcial?",
        guestListTitle: "Lista Completa de Invitados",
        filterAll: "Todos",
        filterAttended: "Asistieron",
        filterPending: "Pendientes",
        thGuest: "Invitado",
        thPasses: "Pases",
        thQr: "Código QR",
        thAttendance: "Asistencia",
        thActions: "Acciones",
        emptyTable: "Carga un archivo Excel para ver la lista de invitados."
    },
    en: {
        scannerTab: "Scanner",
        toolsTab: "Pro Tools",
        guideBtn: "User Guide",
        loadDbTitle: "Load Database",
        loadDbDesc: "Load your original Excel guest list to start attendance check-in.",
        dropLabel: "Drag and drop your Excel here or click to select",
        downloadTplBtn: "Download Sample Excel Template",
        statsTitle: "Real-Time Statistics",
        statGuests: "Guests",
        statAttended: "Attended",
        statPending: "Pending",
        progressLbl: "Event progress",
        reportBtn: "Executive Report",
        downloadExcelBtn: "Download Updated Excel",
        downloadHint: "Download to update records for this entrance gate.",
        scannerTitle: "QR Scanner",
        startScannerBtn: "Start Scanner",
        stopScannerBtn: "Stop Scanner",
        waitingScan: "Waiting for QR code scan...",
        vipBadge: "VIP / SPECIAL GUEST",
        metaPasses: "Passes:",
        metaTime: "Time:",
        partialLabel: "Partial check-in?",
        guestListTitle: "Complete Guest List",
        filterAll: "All",
        filterAttended: "Attended",
        filterPending: "Pending",
        thGuest: "Guest",
        thPasses: "Passes",
        thQr: "QR Code",
        thAttendance: "Attendance",
        thActions: "Actions",
        emptyTable: "Load an Excel file to view the guest list."
    }
};

// Initialization
document.addEventListener("DOMContentLoaded", () => {
    checkAppActivation();
    setupFirebaseSyncIndicator();
    loadSavedBranding();
    setupTabs();
    setupFileLoaders();
    setupSearchAndFilters();
    setupCameraOptions();
    setupBrandingHandlers();
    setupGuideModal();
    
    downloadExcelBtn.addEventListener("click", exportUpdatedExcel);
    generateQrsBtn.addEventListener("click", generateBulkQrsZip);
    
    const downloadTemplateBtn = document.getElementById("download-template-btn");
    if (downloadTemplateBtn) {
        downloadTemplateBtn.addEventListener("click", downloadSampleTemplate);
    }
    
    setupNewFeatureControls();
});

function setupGuideModal() {
    const openGuideBtn = document.getElementById("open-guide-btn");
    const guideModal = document.getElementById("guide-modal");
    const closeGuideModalBtn = document.getElementById("close-guide-modal-btn");
    const guideModalOkBtn = document.getElementById("guide-modal-ok-btn");

    if (openGuideBtn && guideModal) {
        openGuideBtn.addEventListener("click", () => guideModal.classList.remove("hidden"));
        if (closeGuideModalBtn) closeGuideModalBtn.addEventListener("click", () => guideModal.classList.add("hidden"));
        if (guideModalOkBtn) guideModalOkBtn.addEventListener("click", () => guideModal.classList.add("hidden"));
    }
}

/**
 * GENERADOR DE PLANTILLA EXCEL DE EJEMPLO
 * Crea un archivo .xlsx limpio con las columnas correctas e invitados ficticios.
 */
function downloadSampleTemplate() {
    const sampleData = [
        { "ID": "101", "Invitado": "Juan Pérez", "Cantidad": 2, "Mesa": "Mesa 1", "VIP": "VIP", "Teléfono": "5512345678", "Notas": "Vegetariano", "QR": "INV-101", "Asistencia": "" },
        { "ID": "102", "Invitado": "María Rodríguez", "Cantidad": 1, "Mesa": "Mesa 1", "VIP": "VIP", "Teléfono": "5587654321", "Notas": "", "QR": "INV-102", "Asistencia": "" },
        { "ID": "103", "Invitado": "Carlos López", "Cantidad": 3, "Mesa": "Mesa 2", "VIP": "No", "Teléfono": "5533221100", "Notas": "", "QR": "INV-103", "Asistencia": "" },
        { "ID": "104", "Invitado": "Ana Martínez", "Cantidad": 2, "Mesa": "Mesa 2", "VIP": "No", "Teléfono": "5544556677", "Notas": "Alergia mariscos", "QR": "INV-104", "Asistencia": "" },
        { "ID": "105", "Invitado": "Luis García", "Cantidad": 1, "Mesa": "Mesa 3", "VIP": "No", "Teléfono": "5599887766", "Notas": "", "QR": "INV-105", "Asistencia": "" }
    ];

    const worksheet = XLSX.utils.json_to_sheet(sampleData, { header: ["ID", "Invitado", "Cantidad", "Mesa", "VIP", "Teléfono", "Notas", "QR", "Asistencia"] });
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Invitados");
    XLSX.writeFile(workbook, "Plantilla_Invitados_AccesoQR.xlsx");
}

function setupFirebaseSyncIndicator() {
    if (!syncStatusIndicator) return;
    if (!isFirebaseActive) {
        syncStatusIndicator.className = "sync-status-indicator offline";
        syncStatusIndicator.innerHTML = `<span class="dot"></span> Sincro: Desactivada`;
        return;
    }
    
    const connectedRef = firebase.database().ref(".info/connected");
    connectedRef.on("value", (snap) => {
        if (snap.val() === true) {
            syncStatusIndicator.className = "sync-status-indicator online";
            syncStatusIndicator.innerHTML = `<span class="dot"></span> Sincro: En Línea`;
        } else {
            syncStatusIndicator.className = "sync-status-indicator offline";
            syncStatusIndicator.innerHTML = `<span class="dot"></span> Sincro: Offline (Esperando red)`;
        }
    });
}

// ----------------------------------------------------------------------------
// 2. MÓDULO DE SEGURIDAD, CIFRADO Y CONTROL DE ACTIVACIÓN
// ----------------------------------------------------------------------------

const SEC_SALT = "AQR_SEC_2026_PRO_v2!";
let expectedCaptchaAnswer = 0;
let pageLoadTimestamp = Date.now();
let failedActivationAttempts = 0;
let lockoutTimerInterval = null;

/**
 * Guardar datos cifrados/ofuscados en localStorage
 */
function secureSaveKey(key, value) {
    try {
        const payload = JSON.stringify({ v: value, s: SEC_SALT, t: Date.now() });
        const encoded = btoa(encodeURIComponent(payload));
        localStorage.setItem(key, encoded);
    } catch(e) {
        localStorage.setItem(key, String(value));
    }
}

/**
 * Leer datos cifrados/ofuscados de localStorage
 */
function secureGetKey(key) {
    try {
        const raw = localStorage.getItem(key);
        if (!raw) return null;
        const decoded = decodeURIComponent(atob(raw));
        const parsed = JSON.parse(decoded);
        return parsed.v;
    } catch(e) {
        return localStorage.getItem(key);
    }
}

/**
 * Generar desafío matemático Anti-Bot
 */
function generateMathCaptcha() {
    const num1 = Math.floor(Math.random() * 8) + 2;
    const num2 = Math.floor(Math.random() * 8) + 1;
    expectedCaptchaAnswer = num1 + num2;
    const captchaElem = document.getElementById("captcha-question");
    if (captchaElem) {
        captchaElem.textContent = `¿Cuánto es ${num1} + ${num2}?`;
    }
    const answerInput = document.getElementById("captcha-answer-input");
    if (answerInput) answerInput.value = "";
}

/**
 * Control del temporizador de bloqueo por fuerza bruta
 */
function startLockoutTimer(seconds) {
    const lockoutUntil = Date.now() + (seconds * 1000);
    localStorage.setItem("pro-lockout-until", String(lockoutUntil));
    
    activateAppBtn.setAttribute("disabled", "true");
    activationKeyInput.setAttribute("disabled", "true");
    
    if (lockoutTimerInterval) clearInterval(lockoutTimerInterval);
    
    let remaining = seconds;
    showActivationError(`🔒 Bloqueo de Seguridad: Demasiados intentos fallidos. Reintente en ${remaining} segundos.`);
    
    lockoutTimerInterval = setInterval(() => {
        remaining--;
        if (remaining <= 0) {
            clearInterval(lockoutTimerInterval);
            localStorage.removeItem("pro-lockout-until");
            activateAppBtn.removeAttribute("disabled");
            activationKeyInput.removeAttribute("disabled");
            activateAppBtn.innerHTML = 'Activar Aplicación <i class="ti ti-key"></i>';
            activationErrorMsg.classList.add("hidden");
            failedActivationAttempts = 0;
            generateMathCaptcha();
        } else {
            showActivationError(`🔒 Bloqueo de Seguridad: Demasiados intentos fallidos. Reintente en ${remaining} segundos.`);
        }
    }, 1000);
}

// Lógica de Activación y Licencias
function checkAppActivation() {
    generateMathCaptcha();
    
    // Verificar si hay un bloqueo temporal por fuerza bruta activo
    const lockoutUntil = parseInt(localStorage.getItem("pro-lockout-until") || "0", 10);
    if (Date.now() < lockoutUntil) {
        const remainingSec = Math.ceil((lockoutUntil - Date.now()) / 1000);
        startLockoutTimer(remainingSec);
    }

    const isActivated = secureGetKey("pro-license-validated");
    if (isActivated === "true") {
        licenseLockScreen.classList.add("hidden");
    } else {
        licenseLockScreen.classList.remove("hidden");
        activateAppBtn.addEventListener("click", handleActivationSubmit);
    }
}

function handleActivationSubmit() {
    // 1. Verificación Honeypot (Trampa de Bots)
    const hpValue = document.getElementById("sec-hp-field") ? document.getElementById("sec-hp-field").value : "";
    if (hpValue !== "") {
        console.warn("Seguridad: Intento de bot detectado vía Honeypot.");
        showActivationError("Acceso rechazado por filtros de seguridad.");
        return;
    }

    // 2. Verificación de tiempo de interacción humana (< 800ms indica script automatizado)
    if (Date.now() - pageLoadTimestamp < 800) {
        showActivationError("Acceso automatizado detectado. Por favor interactúa manualmente.");
        return;
    }

    // 3. Verificación de Desafío Anti-Bot (Captcha Matemático)
    const userAnswer = parseInt(document.getElementById("captcha-answer-input") ? document.getElementById("captcha-answer-input").value : "0", 10);
    if (isNaN(userAnswer) || userAnswer !== expectedCaptchaAnswer) {
        showActivationError("La respuesta a la verificación humana es incorrecta. Inténtalo de nuevo.");
        generateMathCaptcha();
        return;
    }

    const key = activationKeyInput.value.trim();
    if (!key) {
        showActivationError("Por favor ingresa una clave de licencia.");
        return;
    }

    // Generar o recuperar ID de dispositivo único y persistente
    let deviceId = secureGetKey("pro-device-id");
    if (!deviceId) {
        deviceId = "dev_" + Math.random().toString(36).substr(2, 9) + "_" + Date.now();
        secureSaveKey("pro-device-id", deviceId);
    }

    // Bypass para desarrollo / Pruebas iniciales locales del usuario
    if (key === "TEST-123-KEY") {
        secureSaveKey("pro-license-validated", "true");
        secureSaveKey("pro-active-license-key", key);
        licenseLockScreen.classList.add("hidden");
        playSound('success');
        alert("¡Aplicación activada con éxito (Clave de prueba local)! Conexión Cifrada.");
        return;
    }

    // Si no ha configurado la URL aún
    if (GOOGLE_SCRIPT_URL === "INSERTA_AQUI_TU_URL_DE_GOOGLE_APPS_SCRIPT") {
        showActivationError("El servidor de licencias no está configurado. Configura GOOGLE_SCRIPT_URL en app.js o usa la clave 'TEST-123-KEY' para probar localmente.");
        return;
    }

    activateAppBtn.setAttribute("disabled", "true");
    activateAppBtn.textContent = "Cifrando y Validando...";
    activationErrorMsg.classList.add("hidden");

    // Construcción de Payload Cifrado en Base64
    const secureObj = {
        key: key,
        device: deviceId,
        ts: Date.now(),
        nonce: Math.random().toString(36).substring(2, 10)
    };
    const encodedPayload = btoa(encodeURIComponent(JSON.stringify(secureObj)));

    const fetchUrl = `${GOOGLE_SCRIPT_URL}?payload=${encodeURIComponent(encodedPayload)}&key=${encodeURIComponent(key)}&device=${encodeURIComponent(deviceId)}`;

    fetch(fetchUrl)
        .then(response => response.json())
        .then(data => {
            activateAppBtn.removeAttribute("disabled");
            activateAppBtn.innerHTML = 'Activar Aplicación <i class="ti ti-key"></i>';

            if (data.success) {
                failedActivationAttempts = 0;
                secureSaveKey("pro-license-validated", "true");
                secureSaveKey("pro-active-license-key", key);
                if (data.client) {
                    secureSaveKey("pro-license-client", data.client);
                }
                if (data.guestLimit) {
                    secureSaveKey("pro-license-guest-limit", String(data.guestLimit));
                } else {
                    secureSaveKey("pro-license-guest-limit", "9999");
                }
                if (typeof data.allowQrGen !== "undefined") {
                    secureSaveKey("pro-license-allow-qr-gen", String(data.allowQrGen));
                } else {
                    secureSaveKey("pro-license-allow-qr-gen", "true");
                }
                licenseLockScreen.classList.add("hidden");
                playSound('success');
                alert("🔒 Conexión Cifrada Exitosa. ¡Aplicación autenticada y activada!");
            } else {
                failedActivationAttempts++;
                generateMathCaptcha();
                playSound('error');

                if (failedActivationAttempts >= 3) {
                    startLockoutTimer(30);
                } else {
                    showActivationError(data.message || `Error al validar la licencia. Intentos restantes: ${3 - failedActivationAttempts}`);
                }
            }
        })
        .catch(err => {
            console.error("Security Error:", err);
            activateAppBtn.removeAttribute("disabled");
            activateAppBtn.innerHTML = 'Activar Aplicación <i class="ti ti-key"></i>';
            showActivationError("Error de conexión segura con el servidor. Verifica tu internet e inténtalo de nuevo.");
            playSound('error');
            generateMathCaptcha();
        });
}

function showActivationError(msg) {
    errorText.textContent = msg;
    activationErrorMsg.classList.remove("hidden");
}

// Sound feedback synthesizer
function playSound(type) {
    try {
        const AudioContext = window.AudioContext || window.webkitAudioContext;
        if (!AudioContext) return;
        const ctx = new AudioContext();
        
        const osc = ctx.createOscillator();
        const gain = ctx.createGain();
        osc.connect(gain);
        gain.connect(ctx.destination);
        
        if (type === 'success') {
            osc.frequency.setValueAtTime(880, ctx.currentTime);
            gain.gain.setValueAtTime(0.1, ctx.currentTime);
            osc.start();
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.15);
            osc.stop(ctx.currentTime + 0.15);
        } else if (type === 'warning') {
            osc.frequency.setValueAtTime(587.33, ctx.currentTime);
            gain.gain.setValueAtTime(0.1, ctx.currentTime);
            osc.start();
            gain.gain.setValueAtTime(0.1, ctx.currentTime);
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.12);
            
            setTimeout(() => {
                const ctx2 = new AudioContext();
                const osc2 = ctx2.createOscillator();
                const gain2 = ctx2.createGain();
                osc2.connect(gain2);
                gain2.connect(ctx2.destination);
                osc2.frequency.setValueAtTime(587.33, ctx2.currentTime);
                gain2.gain.setValueAtTime(0.1, ctx2.currentTime);
                osc2.start();
                gain2.gain.exponentialRampToValueAtTime(0.001, ctx2.currentTime + 0.12);
                osc2.stop(ctx2.currentTime + 0.12);
            }, 180);

            osc.stop(ctx.currentTime + 0.12);
        } else if (type === 'error') {
            osc.type = 'sawtooth';
            osc.frequency.setValueAtTime(150, ctx.currentTime);
            gain.gain.setValueAtTime(0.15, ctx.currentTime);
            osc.start();
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.4);
            osc.stop(ctx.currentTime + 0.4);
        } else if (type === 'vip') {
            // Fanfarria triunfal VIP (Arpegio ascendente C5 -> E5 -> G5 -> C6)
            const notes = [523.25, 659.25, 783.99, 1046.50];
            notes.forEach((freq, idx) => {
                setTimeout(() => {
                    try {
                        const vCtx = new (window.AudioContext || window.webkitAudioContext)();
                        const vOsc = vCtx.createOscillator();
                        const vGain = vCtx.createGain();
                        vOsc.type = 'triangle';
                        vOsc.connect(vGain);
                        vGain.connect(vCtx.destination);
                        vOsc.frequency.setValueAtTime(freq, vCtx.currentTime);
                        vGain.gain.setValueAtTime(0.15, vCtx.currentTime);
                        vOsc.start();
                        vGain.gain.exponentialRampToValueAtTime(0.001, vCtx.currentTime + 0.3);
                        vOsc.stop(vCtx.currentTime + 0.3);
                    } catch(err) {}
                }, idx * 110);
            });
        }
    } catch (e) {
        console.warn("Audio feedback error:", e);
    }
}

// Tabs Switcher
function setupTabs() {
    tabButtons.forEach(btn => {
        btn.addEventListener("click", () => {
            tabButtons.forEach(b => b.classList.remove("active"));
            tabContents.forEach(c => c.classList.remove("active"));
            
            btn.classList.add("active");
            document.getElementById(btn.dataset.tab).classList.add("active");
        });
    });
}

// Branding & Customization Lógica
function setupBrandingHandlers() {
    themeColorPicker.addEventListener("input", (e) => {
        const color = e.target.value;
        themeColorCode.textContent = color.toUpperCase();
    });

    logoFileInput.addEventListener("change", (e) => {
        if (e.target.files.length) {
            const file = e.target.files[0];
            logoFileName.textContent = file.name;
        }
    });

    saveBrandingBtn.addEventListener("click", () => {
        const color = themeColorPicker.value;
        const title = appTitleInput.value.trim();
        const logoFile = logoFileInput.files[0];

        // Save Color & Title
        localStorage.setItem("pro-accent-color", color);
        if (title) {
            localStorage.setItem("pro-app-title", title);
        }

        // Apply theme color
        applyAccentColor(color);
        if (title) {
            appTitleDisplay.textContent = title;
        }

        // Handle logo conversion to Base64
        if (logoFile) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const base64Logo = e.target.result;
                localStorage.setItem("pro-app-logo", base64Logo);
                applyLogo(base64Logo);
            };
            reader.readAsDataURL(logoFile);
        }

        alert("¡Personalización de marca guardada con éxito!");
    });

    resetBrandingBtn.addEventListener("click", () => {
        localStorage.removeItem("pro-accent-color");
        localStorage.removeItem("pro-app-title");
        localStorage.removeItem("pro-app-logo");
        
        // Reset defaults
        applyAccentColor("#6366F1");
        themeColorPicker.value = "#6366F1";
        themeColorCode.textContent = "#6366F1";
        
        appTitleDisplay.textContent = "QR Event Check-In PRO";
        appTitleInput.value = "";
        
        logoFileInput.value = "";
        logoFileName.textContent = "Por defecto (Icono QR)";
        
        customLogoImg.classList.add("hidden");
        defaultLogoIcon.classList.remove("hidden");

        alert("Diseño restablecido a valores predeterminados.");
    });
}

function loadSavedBranding() {
    const savedColor = localStorage.getItem("pro-accent-color");
    const savedTitle = localStorage.getItem("pro-app-title");
    const savedLogo = localStorage.getItem("pro-app-logo");

    if (savedColor) {
        themeColorPicker.value = savedColor;
        themeColorCode.textContent = savedColor.toUpperCase();
        applyAccentColor(savedColor);
    }
    
    if (savedTitle) {
        appTitleDisplay.textContent = savedTitle;
        appTitleInput.value = savedTitle;
    }

    if (savedLogo) {
        applyLogo(savedLogo);
    }
}

function applyAccentColor(color) {
    document.documentElement.style.setProperty('--accent', color);
    themeColorMeta.setAttribute("content", color);
}

function applyLogo(base64Logo) {
    customLogoImg.src = base64Logo;
    customLogoImg.classList.remove("hidden");
    defaultLogoIcon.classList.add("hidden");
}

// File Drag & Drop Setup
function setupFileLoaders() {
    dropZone.addEventListener("click", () => excelFileInput.click());
    
    dropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        dropZone.classList.add("dragover");
    });
    
    dropZone.addEventListener("dragleave", () => {
        dropZone.classList.remove("dragover");
    });
    
    dropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        dropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length) {
            handleExcelFile(e.dataTransfer.files[0]);
        }
    });
    
    excelFileInput.addEventListener("change", (e) => {
        if (e.target.files.length) {
            handleExcelFile(e.target.files[0]);
        }
    });

    resetFileBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        resetAppState();
    });
}

function handleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            originalWorkbook = workbook;
            activeSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[activeSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            if (jsonData.length === 0) {
                alert("El archivo Excel está vacío.");
                return;
            }
            
            // Validar Límite Dinámico de Invitados según la licencia contratada
            const limitRaw = secureGetKey("pro-license-guest-limit");
            const guestLimit = limitRaw ? parseInt(limitRaw, 10) : 9999;
            if (jsonData.length > guestLimit) {
                alert(`⚠️ LÍMITE DE LICENCIA ALCANZADO\n\nTu licencia actual permite hasta ${guestLimit} invitados por evento. El archivo que estás intentando cargar contiene ${jsonData.length} invitados.\n\nPara ampliar tu capacidad a 600 o más invitados, por favor contacta a soporte o actualiza tu plan en nuestro sitio web.`);
                return;
            }
            
            const firstRowKeys = Object.keys(jsonData[0]);
            originalColumnsOrder = firstRowKeys;

            guestData = jsonData.map((row, index) => {
                const rawName = String(row["Invitado"] || row["invitado"] || row["Nombre"] || row["nombre"] || "Sin Nombre").trim();
                const tableVal = String(row["Mesa"] || row["mesa"] || row["Table"] || row["table"] || row["Zona"] || row["zona"] || "-").trim();
                const vipVal = String(row["VIP"] || row["vip"] || row["Categoría"] || row["Categoria"] || row["Tipo"] || "").trim().toUpperCase();
                const isVip = (vipVal === "SI" || vipVal === "SÍ" || vipVal === "TRUE" || vipVal === "VIP" || rawName.toUpperCase().includes("(VIP)"));
                const phoneVal = String(row["Telefono"] || row["telefono"] || row["Teléfono"] || row["Celular"] || row["celular"] || row["Phone"] || "").trim();
                const notesVal = String(row["Notas"] || row["notas"] || row["Observaciones"] || row["Alergias"] || "").trim();

                return {
                    id: String(row["ID"] || row["id"] || index + 1).trim(),
                    name: rawName,
                    quantity: parseInt(row["Cantidad"] || row["cantidad"] || row["Pases"] || row["pases"] || 1),
                    table: tableVal,
                    isVip: isVip,
                    phone: phoneVal,
                    notes: notesVal,
                    qrValue: String(row["QR"] || row["qr"] || row["Codigo"] || "").trim(),
                    attendance: String(row["Asistencia"] || row["asistencia"] || "").trim(),
                    rawRow: row
                };
            });

            fileNameDisplay.textContent = file.name;
            dropZone.classList.add("hidden");
            fileInfo.classList.remove("hidden");
            fileStatusIndicator.classList.add("loaded");
            fileStatusIndicator.innerHTML = `<span class="dot"></span> Archivo cargado: ${file.name}`;
            
            searchInput.removeAttribute("disabled");
            toggleCameraBtn.removeAttribute("disabled");
            downloadExcelBtn.removeAttribute("disabled");
            
            // Enable or Disable QR Generator PRO based on Plan
            const allowQrGen = secureGetKey("pro-license-allow-qr-gen") !== "false";
            if (allowQrGen) {
                generateQrsBtn.removeAttribute("disabled");
                qrGenMsg.innerHTML = `<i class="ti ti-circle-check text-success"></i> Lista cargada con <strong>${guestData.length}</strong> invitados listos para generar QRs.`;
            } else {
                generateQrsBtn.setAttribute("disabled", "true");
                qrGenMsg.innerHTML = `<i class="ti ti-lock text-warning"></i> Tu plan contratado es <strong>Solo Escáner</strong>. Para generar códigos QR en masa, actualiza a la Licencia Completa o PRO.`;
            }
            
            // Sincronización en tiempo real Firebase
            if (isFirebaseActive && activeSheetName) {
                const licenseKey = secureGetKey("pro-active-license-key") || "TEST-123-KEY";
                const cleanKey = sanitizeFirebaseKey(licenseKey);
                const cleanEvent = sanitizeFirebaseKey(activeSheetName);
                
                if (dbRef) {
                    dbRef.off();
                }
                
                dbRef = firebase.database().ref(`licencias/${cleanKey}/eventos/${cleanEvent}/asistencias`);
                dbRef.on('value', (snapshot) => {
                    const val = snapshot.val();
                    if (val) {
                        let updated = false;
                        guestData.forEach(guest => {
                            const cleanGuestId = sanitizeFirebaseKey(guest.id);
                            if (val[cleanGuestId] && val[cleanGuestId].attended) {
                                const expectedAttendance = val[cleanGuestId].attendTime ? `Sí (${val[cleanGuestId].attendTime})` : "Sí";
                                if (guest.attendance !== expectedAttendance) {
                                    guest.attendance = expectedAttendance;
                                    guest.rawRow["Asistencia"] = expectedAttendance;
                                    updated = true;
                                }
                            } else {
                                if (guest.attendance) {
                                    guest.attendance = "";
                                    guest.rawRow["Asistencia"] = "";
                                    updated = true;
                                }
                            }
                        });
                        if (updated) {
                            updateDashboard();
                            renderGuestTable();
                        }
                    }
                });
            }

            updateDashboard();
            renderGuestTable();
            playSound('success');
            
        } catch (err) {
            console.error(err);
            alert("Error al leer el archivo Excel. Asegúrate de que sea un archivo de Excel válido.");
        }
    };
    reader.readAsArrayBuffer(file);
}

function resetAppState() {
    if (dbRef) {
        dbRef.off();
        dbRef = null;
    }
    stopScanner();
    guestData = [];
    originalWorkbook = null;
    activeSheetName = "";
    originalColumnsOrder = [];
    
    dropZone.classList.remove("hidden");
    fileInfo.classList.add("hidden");
    fileStatusIndicator.classList.remove("loaded");
    fileStatusIndicator.innerHTML = `<span class="dot"></span> Sin archivo cargado`;
    
    searchInput.setAttribute("disabled", "true");
    searchInput.value = "";
    toggleCameraBtn.setAttribute("disabled", "true");
    downloadExcelBtn.setAttribute("disabled", "true");
    
    // Disable QR Gen PRO
    generateQrsBtn.setAttribute("disabled", "true");
    qrGenMsg.textContent = "Debes cargar la base de datos Excel primero para generar los códigos QR.";
    qrProgressContainer.classList.add("hidden");
    qrProgressLbl.classList.add("hidden");
    
    updateDashboard();
    
    guestTableBody.innerHTML = `
        <tr>
            <td colspan="6" class="table-empty">
                <i class="ti ti-file-excel-off"></i>
                Carga un archivo Excel para ver la lista de invitados.
            </td>
        </tr>
    `;
    
    toggleCameraBtn.innerHTML = `<i class="ti ti-camera"></i> Iniciar Escáner`;
    toggleCameraBtn.className = "btn-secondary";
    scannerOverlay.classList.add("hidden");
    
    resetResultDisplay();
}

function updateDashboard() {
    const total = guestData.length;
    const attended = guestData.filter(g => g.attendance && g.attendance.toLowerCase().trim() !== "").length;
    const pending = total - attended;
    const percentage = total > 0 ? Math.round((attended / total) * 100) : 0;
    
    statTotal.textContent = total;
    statAttended.textContent = attended;
    statPending.textContent = pending;
    
    progressFill.style.width = `${percentage}%`;
    progressPercentage.textContent = `${percentage}%`;
}

function renderGuestTable() {
    if (guestData.length === 0) return;
    
    const searchTerm = searchInput.value.toLowerCase().trim();
    
    let filteredList = guestData.filter(guest => {
        const matchesSearch = 
            guest.name.toLowerCase().includes(searchTerm) || 
            guest.id.toLowerCase().includes(searchTerm) || 
            (guest.table && guest.table.toLowerCase().includes(searchTerm)) ||
            (guest.phone && guest.phone.includes(searchTerm)) ||
            guest.qrValue.toLowerCase().includes(searchTerm);
            
        const isAttended = guest.attendance && guest.attendance.toLowerCase().trim() !== "";
        if (currentFilter === "attended") return matchesSearch && isAttended;
        if (currentFilter === "pending") return matchesSearch && !isAttended;
        return matchesSearch;
    });

    if (filteredList.length === 0) {
        guestTableBody.innerHTML = `
            <tr>
                <td colspan="7" class="table-empty">
                    <i class="ti ti-search-off"></i>
                    ${currentLang === 'en' ? 'No guests found.' : 'No se encontraron invitados.'}
                </td>
            </tr>
        `;
        return;
    }

    guestTableBody.innerHTML = filteredList.map(guest => {
        const hasAttended = guest.attendance && guest.attendance.toLowerCase().trim() !== "";
        const statusBadge = hasAttended 
            ? `<span class="badge success"><i class="ti ti-check"></i> ${guest.attendance}</span>` 
            : `<span class="badge pending">${currentLang === 'en' ? 'Pending' : 'Pendiente'}</span>`;
            
        const actionButton = hasAttended
            ? `<button class="btn-action btn-secondary" onclick="toggleAttendance('${guest.id}', true)"><i class="ti ti-rotate-clockwise"></i> ${currentLang === 'en' ? 'Revert' : 'Revertir'}</button>`
            : `<button class="btn-action btn-checkin" onclick="toggleAttendance('${guest.id}', false)"><i class="ti ti-circle-check"></i> ${currentLang === 'en' ? 'Check-in' : 'Registrar'}</button>`;

        const vipBadgeHtml = guest.isVip ? `<span class="vip-badge-tag"><i class="ti ti-crown"></i> VIP</span> ` : "";
        const tableBadgeHtml = (guest.table && guest.table !== "-") 
            ? `<span class="badge-vip-table"><i class="ti ti-armchair"></i> ${guest.table}</span>` 
            : `<span style="color:var(--text-secondary)">-</span>`;

        return `
            <tr>
                <td><strong>${guest.id}</strong></td>
                <td>${vipBadgeHtml}<strong>${guest.name}</strong></td>
                <td>${guest.quantity}</td>
                <td>${tableBadgeHtml}</td>
                <td><code>${guest.qrValue || guest.id}</code></td>
                <td>${statusBadge}</td>
                <td>
                    <div style="display: flex; gap: 6px; align-items: center; flex-wrap: wrap;">
                        ${actionButton}
                        <button class="btn-whatsapp" title="Enviar pase por WhatsApp" onclick="sharePassWhatsApp('${guest.id}')"><i class="ti ti-brand-whatsapp"></i></button>
                        <button class="btn-qr-view" title="Ver / Descargar Código QR" onclick="openIndividualQrModal('${guest.id}')"><i class="ti ti-qrcode"></i></button>
                    </div>
                </td>
            </tr>
        `;
    }).join('');
}

function toggleAttendance(id, revert = false) {
    const guest = guestData.find(g => g.id === id);
    if (!guest) return;

    if (revert) {
        const guestIndex = guestData.findIndex(g => g.id === id);
        guestData[guestIndex].attendance = "";
        guestData[guestIndex].rawRow["Asistencia"] = "";
        syncCheckInToFirebase(id, false, "", "Revertido Manual");
        updateDashboard();
        renderGuestTable();
        resetResultDisplay();
    } else {
        processCheckIn(guest, guest.quantity, "Manual");
    }
}

function processCheckInById(id, qty) {
    const guest = guestData.find(g => g.id === id);
    if (guest) {
        processCheckIn(guest, qty, "Parcial");
    }
}

function processCheckIn(guest, qtyToCheckIn = null, source = "Escáner") {
    const guestIndex = guestData.findIndex(g => g.id === guest.id);
    if (guestIndex === -1) return;

    const isFullCheckIn = (qtyToCheckIn === null || qtyToCheckIn >= guest.quantity);
    const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });

    let attendanceStr = "";
    if (isFullCheckIn) {
        attendanceStr = `Sí (${time})`;
    } else {
        attendanceStr = `Parcial (${qtyToCheckIn}/${guest.quantity} a las ${time})`;
    }

    guestData[guestIndex].attendance = attendanceStr;
    guestData[guestIndex].rawRow["Asistencia"] = attendanceStr;

    // Registrar métrica de tiempo para velocidad
    checkinTimestamps.push({
        time: Date.now(),
        qty: isFullCheckIn ? guest.quantity : qtyToCheckIn,
        hour: time.split(":")[0] + ":00"
    });
    updateFlowRate();

    // Sincronizar en tiempo real a Firebase
    syncCheckInToFirebase(guest.id, true, time, `${source} ${!isFullCheckIn ? '(' + qtyToCheckIn + '/' + guest.quantity + ')' : ''}`);

    updateDashboard();
    renderGuestTable();

    // Mostrar tarjeta de resultado completa
    displayResult(
        guest.isVip ? 'vip' : 'success',
        guest.isVip ? (currentLang === 'en' ? '👑 Welcome VIP Guest!' : '👑 ¡Bienvenido Invitado VIP!') : (currentLang === 'en' ? 'Successful Check-in!' : '¡Registro Exitoso!'),
        guest.name,
        guest.id,
        isFullCheckIn ? guest.quantity : `${qtyToCheckIn}/${guest.quantity}`,
        attendanceStr,
        guest.table,
        guest.notes,
        guest
    );

    // Audio & Voz Concierge
    playSound(guest.isVip ? 'vip' : 'success');
    speakWelcome(guest, isFullCheckIn ? guest.quantity : qtyToCheckIn);
}

function setupSearchAndFilters() {
    searchInput.addEventListener("input", renderGuestTable);
    
    filterButtons.forEach(btn => {
        btn.addEventListener("click", () => {
            filterButtons.forEach(b => b.classList.remove("active"));
            btn.classList.add("active");
            currentFilter = btn.dataset.filter;
            renderGuestTable();
        });
    });
}

// Camera controls
function setupCameraOptions() {
    Html5Qrcode.getCameras().then(devices => {
        if (devices && devices.length) {
            cameraList = devices;
            cameraSelect.innerHTML = devices.map((device, idx) => 
                `<option value="${device.id}">${device.label || `Cámara ${idx + 1}`}</option>`
            ).join('');
            
            toggleCameraBtn.addEventListener("click", toggleScanner);
        } else {
            cameraSelect.innerHTML = `<option value="">Cámaras no disponibles</option>`;
        }
    }).catch(err => {
        console.error("Camera list error:", err);
        cameraSelect.innerHTML = `<option value="">Error de permisos</option>`;
    });
}

function toggleScanner() {
    if (isScannerActive) {
        stopScanner();
    } else {
        startScanner();
    }
}

function startScanner() {
    const cameraId = cameraSelect.value;
    if (!cameraId) {
        alert("Por favor selecciona una cámara.");
        return;
    }

    resetResultDisplay();
    isScannerActive = true;
    
    toggleCameraBtn.innerHTML = `<i class="ti ti-camera-off"></i> ${currentLang === 'en' ? 'Stop Scanner' : 'Detener Escáner'}`;
    toggleCameraBtn.className = "btn-secondary btn-danger-hover";
    scannerOverlay.classList.remove("hidden");

    html5QrScanner = new Html5Qrcode("qr-reader");
    html5QrScanner.start(
        cameraId, 
        {
            fps: 10,
            qrbox: (width, height) => {
                const size = Math.min(width, height) * 0.7;
                return { width: size, height: size };
            }
        },
        onQrCodeSuccess,
        onQrCodeError
    ).then(() => {
        // Detectar si el dispositivo y cámara soportan Linterna / Flash
        setTimeout(() => {
            const videoElem = document.querySelector("#qr-reader video");
            if (videoElem && videoElem.srcObject) {
                const tracks = videoElem.srcObject.getVideoTracks();
                if (tracks && tracks.length > 0) {
                    activeVideoTrack = tracks[0];
                    const caps = activeVideoTrack.getCapabilities ? activeVideoTrack.getCapabilities() : {};
                    if (caps.torch && torchBtn) {
                        torchBtn.classList.remove("hidden");
                    }
                }
            }
        }, 800);
    }).catch(err => {
        console.error(err);
        stopScanner();
        alert("No se pudo acceder a la cámara.");
    });
}

function stopScanner() {
    isScannerActive = false;
    toggleCameraBtn.innerHTML = `<i class="ti ti-camera"></i> ${currentLang === 'en' ? 'Start Scanner' : 'Iniciar Escáner'}`;
    toggleCameraBtn.className = "btn-secondary";
    scannerOverlay.classList.add("hidden");

    if (torchBtn) {
        torchBtn.classList.add("hidden");
        torchBtn.classList.remove("active");
        isTorchOn = false;
    }
    activeVideoTrack = null;

    if (html5QrScanner) {
        html5QrScanner.stop().then(() => {
            html5QrScanner = null;
        }).catch(err => console.error(err));
    }
}

function onQrCodeSuccess(decodedText) {
    let qrVal = decodedText.trim();
    
    let guest = guestData.find(g => 
        (g.qrValue && g.qrValue.toLowerCase() === qrVal.toLowerCase()) || 
        g.id.toLowerCase() === qrVal.toLowerCase()
    );

    if (!guest) {
        try {
            const url = new URL(decodedText);
            const idParam = url.searchParams.get("id");
            if (idParam) {
                guest = guestData.find(g => g.id.toLowerCase() === idParam.toLowerCase().trim());
            }
        } catch(e) {}
    }

    if (!guest) {
        displayResult('danger', currentLang === 'en' ? 'Guest Not Found' : 'Invitado No Encontrado', `${currentLang === 'en' ? 'The QR code' : 'El código QR'} "${qrVal}" ${currentLang === 'en' ? 'is not registered.' : 'no está registrado.'}`, qrVal);
        playSound('error');
        return;
    }

    const hasAttended = guest.attendance && guest.attendance.toLowerCase().trim() !== "";
    if (hasAttended && !guest.attendance.includes("Parcial")) {
        displayResult('warning', currentLang === 'en' ? 'Already Checked-In' : 'Asistencia Ya Registrada', guest.name, guest.id, guest.quantity, guest.attendance, guest.table, guest.notes, guest);
        playSound('warning');
    } else {
        processCheckIn(guest, guest.quantity, "Escáner");
    }
}

function onQrCodeError(errorMessage) {}

function resetResultDisplay() {
    resultCard.className = "card result-card";
    resultPlaceholder.classList.remove("hidden");
    resultDetails.classList.add("hidden");
}

function displayResult(type, title, name, id = "-", qty = "-", time = "-", table = "-", notes = "", guestObj = null) {
    resultPlaceholder.classList.add("hidden");
    resultDetails.classList.remove("hidden");
    resultCard.className = "card result-card " + type;
    
    resultTitle.textContent = title;
    resultName.textContent = name;
    resultId.textContent = id;
    resultQty.textContent = qty;
    resultTime.textContent = time;

    if (resultTable) resultTable.textContent = table || "-";
    
    if (resultNotesRow) {
        if (notes && notes.trim() !== "") {
            resultNotes.textContent = notes;
            resultNotesRow.classList.remove("hidden");
        } else {
            resultNotesRow.classList.add("hidden");
        }
    }

    if (resultVipBanner) {
        if (type === 'vip' || (guestObj && guestObj.isVip)) {
            resultVipBanner.classList.remove("hidden");
        } else {
            resultVipBanner.classList.add("hidden");
        }
    }

    if (partialCheckinBox && partialButtonsGroup) {
        if (guestObj && guestObj.quantity > 1 && (!guestObj.attendance || guestObj.attendance.includes("Parcial"))) {
            partialCheckinBox.classList.remove("hidden");
            let btnsHtml = "";
            for (let i = 1; i < guestObj.quantity; i++) {
                btnsHtml += `<button class="partial-btn" onclick="processCheckInById('${guestObj.id}', ${i})">${i} pase${i > 1 ? 's' : ''}</button>`;
            }
            partialButtonsGroup.innerHTML = btnsHtml;
        } else {
            partialCheckinBox.classList.add("hidden");
        }
    }

    resultStatusIconWrapper.className = "result-status-icon-wrapper";
    if (type === 'vip') {
        resultIcon.className = "ti ti-crown";
    } else if (type === 'success') {
        resultIcon.className = "ti ti-circle-check";
    } else if (type === 'warning') {
        resultIcon.className = "ti ti-alert-triangle";
    } else {
        resultIcon.className = "ti ti-circle-x";
    }
}

function exportUpdatedExcel() {
    if (guestData.length === 0) return;

    const outputRows = guestData.map(guest => {
        const row = { ...guest.rawRow };
        const idKey = Object.keys(row).find(k => k.toLowerCase() === "id") || "ID";
        const nameKey = Object.keys(row).find(k => k.toLowerCase() === "invitado") || "Invitado";
        const qtyKey = Object.keys(row).find(k => k.toLowerCase() === "cantidad") || "Cantidad";
        const qrKey = Object.keys(row).find(k => k.toLowerCase() === "qr") || "QR";
        const attendanceKey = Object.keys(row).find(k => k.toLowerCase() === "asistencia") || "Asistencia";

        row[idKey] = guest.id;
        row[nameKey] = guest.name;
        row[qtyKey] = guest.quantity;
        row[qrKey] = guest.qrValue;
        row[attendanceKey] = guest.attendance;
        return row;
    });

    try {
        const newWorksheet = XLSX.utils.json_to_sheet(outputRows, { header: originalColumnsOrder });
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, activeSheetName);
        XLSX.writeFile(newWorkbook, "Asistencia_PRO.xlsx");
        playSound('success');
    } catch(err) {
        console.error(err);
        alert("Error al exportar.");
    }
}

// PRO FEATURE: QR Generator & Bulk ZIP Downloader
function generateBulkQrsZip() {
    if (guestData.length === 0) return;

    generateQrsBtn.setAttribute("disabled", "true");
    qrProgressContainer.classList.remove("hidden");
    qrProgressLbl.classList.remove("hidden");

    const zip = new JSZip();
    let currentIdx = 0;
    const totalGuests = guestData.length;

    // Use intervals to prevent thread blocking so browser doesn't freeze
    const interval = setInterval(() => {
        if (currentIdx >= totalGuests) {
            clearInterval(interval);
            
            // Finalize and download ZIP
            qrGenMsg.textContent = "Empaquetando archivos QRs en un ZIP...";
            zip.generateAsync({ type: "blob" }).then(function(content) {
                // Custom trigger download anchor to support offline environment
                const link = document.createElement("a");
                link.href = URL.createObjectURL(content);
                link.download = "boletos_qr_evento.zip";
                link.click();
                
                // Reset UI
                generateQrsBtn.removeAttribute("disabled");
                qrGenMsg.innerHTML = `<i class="ti ti-circle-check text-success"></i> ¡Descarga completa! Se guardó <strong>${totalGuests}</strong> códigos QR en un ZIP.`;
                qrProgressContainer.classList.add("hidden");
                qrProgressLbl.classList.add("hidden");
                playSound('success');
            }).catch(err => {
                console.error("ZIP Generation error:", err);
                alert("Error al generar el archivo ZIP.");
                generateQrsBtn.removeAttribute("disabled");
            });
            return;
        }

        // Process guest QR code
        const guest = guestData[currentIdx];
        const qrContent = guest.qrValue || guest.id;
        
        try {
            // Generate QR code using qrcode-generator library
            const typeNumber = 0; // Auto detect size
            const errorCorrectionLevel = 'H'; // High correction
            const qr = qrcode(typeNumber, errorCorrectionLevel);
            qr.addData(qrContent);
            qr.make();
            
            // Retrieve base64 png payload from the library
            // Cell size = 8, margin = 2
            const imgDataUri = qr.createDataURL(8, 2);
            const base64Data = imgDataUri.split(',')[1];
            
            // Clean guest name for filename
            const cleanName = guest.name.replace(/[^a-z0-9]/gi, '_');
            const filename = `QR_${guest.id}_${cleanName}.png`;
            
            zip.file(filename, base64Data, { base64: true });
        } catch(e) {
            console.error("Individual QR creation error:", e);
        }

        currentIdx++;
        
        // Update progress bar
        const progressPercent = Math.round((currentIdx / totalGuests) * 100);
        qrProgressFill.style.width = `${progressPercent}%`;
        qrProgressLbl.textContent = `Generando: ${currentIdx} / ${totalGuests}`;
    }, 20); // Small delay to allow UI to render progress
}

// PRO FEATURE: Multi-door Excel Merging
function setupMergeHandlers() {
    mergeDropZone.addEventListener("click", () => mergeFilesInput.click());
    
    mergeDropZone.addEventListener("dragover", (e) => {
        e.preventDefault();
        mergeDropZone.classList.add("dragover");
    });
    
    mergeDropZone.addEventListener("dragleave", () => {
        mergeDropZone.classList.remove("dragover");
    });
    
    mergeDropZone.addEventListener("drop", (e) => {
        e.preventDefault();
        mergeDropZone.classList.remove("dragover");
        if (e.dataTransfer.files.length) {
            handleMergeFilesSelection(e.dataTransfer.files);
        }
    });
    
    mergeFilesInput.addEventListener("change", (e) => {
        if (e.target.files.length) {
            handleMergeFilesSelection(e.target.files);
        }
    });

    processMergeBtn.addEventListener("click", processExcelMerging);
    
    resetMergeBtn.addEventListener("click", () => {
        mergeFilesData = [];
        mergedFilesListWrapper.classList.add("hidden");
        mergeDropZone.classList.remove("hidden");
        mergeFilesInput.value = "";
    });
}

function handleMergeFilesSelection(filesList) {
    let filesLoaded = 0;
    const totalFiles = filesList.length;
    mergeFilesData = [];
    mergedFilesList.innerHTML = "";

    Array.from(filesList).forEach(file => {
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                const rows = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
                
                mergeFilesData.push({
                    fileName: file.name,
                    rows: rows,
                    sheetName: sheetName,
                    columnsOrder: Object.keys(rows[0] || {})
                });

                // Display file badge in UI list
                const sizeKB = Math.round(file.size / 1024);
                mergedFilesList.innerHTML += `
                    <li>
                        <span><i class="ti ti-file-check text-success"></i> <strong>${file.name}</strong> (${rows.length} filas)</span>
                        <span class="logo-file-name">${sizeKB} KB</span>
                    </li>
                `;

                filesLoaded++;
                if (filesLoaded === totalFiles) {
                    mergeDropZone.classList.add("hidden");
                    mergedFilesListWrapper.classList.remove("hidden");
                    playSound('success');
                }
            } catch(err) {
                console.error(err);
                alert(`Error al cargar el archivo: ${file.name}`);
            }
        };
        reader.readAsArrayBuffer(file);
    });
}

function processExcelMerging() {
    if (mergeFilesData.length < 2) {
        alert("Por favor carga al menos 2 archivos de Excel para realizar la fusión.");
        return;
    }

    try {
        // Use the first file loaded as the template base structure
        const baseFile = mergeFilesData[0];
        const mergedRows = JSON.parse(JSON.stringify(baseFile.rows)); // Deep copy
        
        // Find the keys/columns for ID and Assistance
        const firstRow = mergedRows[0] || {};
        const idKey = Object.keys(firstRow).find(k => k.toLowerCase() === "id") || "ID";
        const attendanceKey = Object.keys(firstRow).find(k => k.toLowerCase() === "asistencia") || "Asistencia";

        // Mapeamos todas las filas de los archivos secundarios indexándolas por su ID
        const logsMap = {}; // id -> list of assistance strings

        mergeFilesData.forEach(fileObj => {
            fileObj.rows.forEach(row => {
                const idVal = String(row[idKey] || "").trim();
                const attVal = String(row[attendanceKey] || "").trim();
                
                if (idVal && attVal && attVal.toLowerCase() !== "") {
                    if (!logsMap[idVal]) {
                        logsMap[idVal] = [];
                    }
                    logsMap[idVal].push(attVal);
                }
            });
        });

        // Fusionamos los datos de asistencia en nuestra lista base
        let updatedCount = 0;
        mergedRows.forEach(row => {
            const idVal = String(row[idKey] || "").trim();
            const originalAtt = String(row[attendanceKey] || "").trim();
            
            if (idVal && logsMap[idVal] && logsMap[idVal].length > 0) {
                // If template did not have attendance, or we found new logs
                // We pick the first attendance log we found (or earliest one)
                // Let's filter unique values
                const uniqueLogs = [...new Set(logsMap[idVal])];
                row[attendanceKey] = uniqueLogs.join(" | "); // Combine logs in case of double scan or pick first
                updatedCount++;
            } else {
                row[attendanceKey] = originalAtt; // Keep original (usually empty)
            }
        });

        // Write workbook and download
        const newWorksheet = XLSX.utils.json_to_sheet(mergedRows, { header: baseFile.columnsOrder });
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, baseFile.sheetName);
        
        XLSX.writeFile(newWorkbook, "Reporte_FUSIONADO.xlsx");
        
        playSound('success');
        alert(`¡Fusión completa! Se consolidó la asistencia de ${updatedCount} invitados.`);
        
    } catch(err) {
        console.error("Merge execution error:", err);
        alert("Ocurrió un error inesperado al fusionar las bases de datos.");
    }
}

// ============================================================================
// NUEVAS MEJORAS DE ALTO IMPACTO (CONTROLES, IDIOMA, VOZ, WHATSAPP Y REPORTES)
// ============================================================================

function setupNewFeatureControls() {
    // 1. Inicializar Tema (☀️ Modo Día / 🌙 Modo Noche)
    if (isLightTheme) {
        document.body.classList.add("light-theme");
        if (themeIcon) themeIcon.className = "ti ti-moon";
        if (themeBtnText) themeBtnText.textContent = "Modo Noche";
    }

    if (themeToggleBtn) {
        themeToggleBtn.addEventListener("click", () => {
            isLightTheme = !isLightTheme;
            if (isLightTheme) {
                document.body.classList.add("light-theme");
                if (themeIcon) themeIcon.className = "ti ti-moon";
                if (themeBtnText) themeBtnText.textContent = currentLang === 'en' ? "Dark Mode" : "Modo Noche";
                localStorage.setItem("pro-theme", "light");
            } else {
                document.body.classList.remove("light-theme");
                if (themeIcon) themeIcon.className = "ti ti-sun";
                if (themeBtnText) themeBtnText.textContent = currentLang === 'en' ? "Sunlight Mode" : "Modo Día";
                localStorage.setItem("pro-theme", "dark");
            }
        });
    }

    // 2. Inicializar Voz de Bienvenida Concierge
    if (voiceBtnText) {
        voiceBtnText.textContent = isVoiceEnabled 
            ? (currentLang === 'en' ? "Voice: ON" : "Voz: Sí") 
            : (currentLang === 'en' ? "Voice: OFF" : "Voz: No");
    }
    if (voiceIcon) {
        voiceIcon.className = isVoiceEnabled ? "ti ti-volume" : "ti ti-volume-off";
    }

    if (voiceToggleBtn) {
        voiceToggleBtn.addEventListener("click", () => {
            isVoiceEnabled = !isVoiceEnabled;
            localStorage.setItem("pro-voice-enabled", isVoiceEnabled);
            if (voiceBtnText) {
                voiceBtnText.textContent = isVoiceEnabled 
                    ? (currentLang === 'en' ? "Voice: ON" : "Voz: Sí") 
                    : (currentLang === 'en' ? "Voice: OFF" : "Voz: No");
            }
            if (voiceIcon) {
                voiceIcon.className = isVoiceEnabled ? "ti ti-volume" : "ti ti-volume-off";
            }
        });
    }

    // 3. Inicializar Selector de Idioma (ES / EN)
    if (langBtnText) {
        langBtnText.textContent = currentLang === 'en' ? "ES" : "EN";
    }
    if (langToggleBtn) {
        langToggleBtn.addEventListener("click", () => {
            currentLang = currentLang === 'es' ? 'en' : 'es';
            localStorage.setItem("pro-lang", currentLang);
            if (langBtnText) {
                langBtnText.textContent = currentLang === 'en' ? "ES" : "EN";
            }
            applyLanguage();
        });
    }
    applyLanguage();

    // 4. Linterna / Flash
    if (torchBtn) {
        torchBtn.addEventListener("click", () => {
            if (!activeVideoTrack) return;
            try {
                isTorchOn = !isTorchOn;
                activeVideoTrack.applyConstraints({
                    advanced: [{ torch: isTorchOn }]
                }).then(() => {
                    if (isTorchOn) {
                        torchBtn.classList.add("active");
                        torchBtn.innerHTML = '<i class="ti ti-bolt-off"></i>';
                    } else {
                        torchBtn.classList.remove("active");
                        torchBtn.innerHTML = '<i class="ti ti-bolt"></i>';
                    }
                }).catch(err => {
                    console.warn("Torch constraint error:", err);
                });
            } catch(e) {
                console.warn("Torch failed:", e);
            }
        });
    }

    // 5. Modal Reporte Ejecutivo
    if (openReportBtn) {
        openReportBtn.addEventListener("click", openExecutiveReport);
    }
    if (closeRepModalBtn && executiveReportModal) {
        closeRepModalBtn.addEventListener("click", () => executiveReportModal.classList.add("hidden"));
    }
    if (printRepBtn) {
        printRepBtn.addEventListener("click", () => window.print());
    }

    // 6. Modal QR Individual
    if (closeIndivQrModalBtn && individualQrModal) {
        closeIndivQrModalBtn.addEventListener("click", () => individualQrModal.classList.add("hidden"));
    }
}

function applyLanguage() {
    const langData = i18n[currentLang] || i18n.es;
    
    document.querySelectorAll("[data-i18n]").forEach(el => {
        const key = el.getAttribute("data-i18n");
        if (langData[key]) {
            el.textContent = langData[key];
        }
    });

    if (searchInput) {
        searchInput.placeholder = currentLang === 'en' 
            ? "Search by name, ID, table..." 
            : "Buscar por nombre, ID, mesa...";
    }

    if (themeBtnText) {
        themeBtnText.textContent = isLightTheme 
            ? (currentLang === 'en' ? "Dark Mode" : "Modo Noche") 
            : (currentLang === 'en' ? "Sunlight Mode" : "Modo Día");
    }

    if (voiceBtnText) {
        voiceBtnText.textContent = isVoiceEnabled 
            ? (currentLang === 'en' ? "Voice: ON" : "Voz: Sí") 
            : (currentLang === 'en' ? "Voice: OFF" : "Voz: No");
    }

    renderGuestTable();
}

function speakWelcome(guest, qty) {
    if (!isVoiceEnabled || !('speechSynthesis' in window)) return;
    try {
        window.speechSynthesis.cancel();
        
        let text = "";
        const tableText = (guest.table && guest.table !== "-") 
            ? (currentLang === 'en' ? `, Table ${guest.table}` : `, Mesa ${guest.table}`) 
            : "";
        const passesText = (qty && qty > 1) 
            ? (currentLang === 'en' ? `, ${qty} guests` : `, ${qty} personas`) 
            : "";

        if (currentLang === 'en') {
            if (guest.isVip) {
                text = `Welcome ${guest.name}, VIP guest!${tableText}`;
            } else {
                text = `Welcome ${guest.name}!${tableText}${passesText}`;
            }
        } else {
            if (guest.isVip) {
                text = `¡Bienvenido ${guest.name}, invitado de honor!${tableText}`;
            } else {
                text = `¡Bienvenido ${guest.name}!${tableText}${passesText}`;
            }
        }

        const utterance = new SpeechSynthesisUtterance(text);
        utterance.rate = 1.0;
        utterance.pitch = 1.0;
        utterance.lang = currentLang === 'en' ? 'en-US' : 'es-MX';
        window.speechSynthesis.speak(utterance);
    } catch(e) {
        console.warn("TTS error:", e);
    }
}

function sharePassWhatsApp(guestId) {
    const guest = guestData.find(g => g.id === guestId);
    if (!guest) return;
    
    const eventName = appTitleDisplay ? appTitleDisplay.textContent : "AccesoQR PRO";
    const tableInfo = (guest.table && guest.table !== "-") ? `\n📍 Mesa: ${guest.table}` : "";
    const passesInfo = `\n🎟️ Pases autorizados: ${guest.quantity}`;
    const codeInfo = `\n🔑 Código de acceso: ${guest.qrValue || guest.id}`;
    
    let text = `¡Hola *${guest.name}*! Te compartimos tu pase de acceso personal para *${eventName}*:${passesInfo}${tableInfo}${codeInfo}\n\nPresenta este mensaje o tu código QR en la puerta de entrada para registrar tu asistencia. ¡Te esperamos!`;
    
    let cleanPhone = (guest.phone || "").replace(/\D/g, "");
    let waUrl = "";
    if (cleanPhone.length >= 10) {
        if (cleanPhone.length === 10) cleanPhone = "52" + cleanPhone; // México por defecto
        waUrl = `https://wa.me/${cleanPhone}?text=${encodeURIComponent(text)}`;
    } else {
        waUrl = `https://api.whatsapp.com/send?text=${encodeURIComponent(text)}`;
    }
    
    window.open(waUrl, "_blank");
}

function openIndividualQrModal(guestId) {
    const guest = guestData.find(g => g.id === guestId);
    if (!guest) return;

    indivGuestName.textContent = guest.name;
    indivGuestId.textContent = guest.id;
    indivGuestQty.textContent = guest.quantity;
    indivGuestTable.textContent = guest.table || "-";

    try {
        const qr = qrcode(0, 'M');
        qr.addData(guest.qrValue || guest.id);
        qr.make();
        indivQrCodeBox.innerHTML = qr.createImgTag(5, 8);
    } catch(e) {
        indivQrCodeBox.innerHTML = `<span style="color:red">Error generando QR</span>`;
    }

    indivWaBtn.onclick = () => sharePassWhatsApp(guest.id);
    indivDownloadBtn.onclick = () => downloadIndividualQrImage(guest);

    individualQrModal.classList.remove("hidden");
}

function downloadIndividualQrImage(guest) {
    const img = indivQrCodeBox.querySelector("img");
    if (!img) return;
    const a = document.createElement("a");
    a.href = img.src;
    a.download = `Pase_QR_${guest.id}_${guest.name.replace(/\s+/g, '_')}.png`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

function updateFlowRate() {
    const fiveMinutesAgo = Date.now() - (5 * 60 * 1000);
    const recentCheckins = checkinTimestamps.filter(c => c.time >= fiveMinutesAgo);
    const totalRecentGuests = recentCheckins.reduce((acc, curr) => acc + (curr.qty || 1), 0);
    const ratePerMin = Math.round((totalRecentGuests / 5) * 10) / 10;
    
    if (flowRateText) {
        flowRateText.textContent = `Flujo: ${ratePerMin} pers/min`;
    }
}

function openExecutiveReport() {
    const total = guestData.length;
    const attended = guestData.filter(g => g.attendance && g.attendance.toLowerCase().trim() !== "").length;
    const pending = total - attended;
    const percentage = total > 0 ? Math.round((attended / total) * 100) : 0;

    repEventTitle.textContent = "Evento: " + (appTitleDisplay ? appTitleDisplay.textContent : "AccesoQR PRO");
    repTotalGuests.textContent = total;
    repTotalAttended.textContent = attended;
    repTotalPending.textContent = pending;
    repPercentage.textContent = `${percentage}%`;

    // Calcular hora pico
    const hourCounts = {};
    checkinTimestamps.forEach(item => {
        if (item.hour) {
            hourCounts[item.hour] = (hourCounts[item.hour] || 0) + (item.qty || 1);
        }
    });

    let peakHour = "--:--";
    let peakCount = 0;
    Object.keys(hourCounts).forEach(h => {
        if (hourCounts[h] > peakCount) {
            peakCount = hourCounts[h];
            peakHour = h;
        }
    });

    repPeakHour.textContent = peakHour !== "--:--" ? peakHour : (attended > 0 ? "Registro continuo" : "--:--");
    repPeakRate.textContent = `${peakCount} asistentes`;

    // Desglose por mesas
    const tableStats = {};
    guestData.forEach(g => {
        const t = (g.table && g.table !== "-") ? g.table : "Sin Mesa";
        if (!tableStats[t]) {
            tableStats[t] = { total: 0, attended: 0 };
        }
        tableStats[t].total += g.quantity;
        if (g.attendance && g.attendance.toLowerCase().trim() !== "") {
            tableStats[t].attended += g.quantity;
        }
    });

    const tablesHtml = Object.keys(tableStats).map(t => {
        const item = tableStats[t];
        const pct = item.total > 0 ? Math.round((item.attended / item.total) * 100) : 0;
        return `
            <div style="display: flex; justify-content: space-between; align-items: center; padding: 6px 0; border-bottom: 1px solid rgba(255,255,255,0.08);">
                <span><strong>${t}</strong></span>
                <span>${item.attended} / ${item.total} pases (${pct}%)</span>
            </div>
        `;
    }).join("");

    repTablesBreakdown.innerHTML = tablesHtml || `<span style="color:var(--text-secondary)">No hay datos de mesas cargados.</span>`;

    executiveReportModal.classList.remove("hidden");
}

