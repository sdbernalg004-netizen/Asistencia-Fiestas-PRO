// Core Application Logic for QR Event Check-In PRO

// CONFIGURACIÓN DE LICENCIA (Pega aquí la URL de tu Google Apps Script una vez implementado)
const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbze_E6zOwQyUUaj-gdc2WcOV0NeDBassrXUQFuA1ERAl8mIlolRi1mlybKZ9i67x1lt-w/exec";

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

// Initialization
document.addEventListener("DOMContentLoaded", () => {
    checkAppActivation();
    loadSavedBranding();
    setupTabs();
    setupFileLoaders();
    setupSearchAndFilters();
    setupCameraOptions();
    setupBrandingHandlers();
    setupMergeHandlers();
    
    downloadExcelBtn.addEventListener("click", exportUpdatedExcel);
    generateQrsBtn.addEventListener("click", generateBulkQrsZip);
});

// Lógica de Activación y Licencias
function checkAppActivation() {
    const isActivated = localStorage.getItem("pro-license-validated");
    if (isActivated === "true") {
        licenseLockScreen.classList.add("hidden");
    } else {
        licenseLockScreen.classList.remove("hidden");
        activateAppBtn.addEventListener("click", handleActivationSubmit);
    }
}

function handleActivationSubmit() {
    const key = activationKeyInput.value.trim();
    if (!key) {
        showActivationError("Por favor ingresa una clave de licencia.");
        return;
    }

    // Generar o recuperar ID de dispositivo único y persistente
    let deviceId = localStorage.getItem("pro-device-id");
    if (!deviceId) {
        deviceId = "dev_" + Math.random().toString(36).substr(2, 9) + "_" + Date.now();
        localStorage.setItem("pro-device-id", deviceId);
    }

    // Bypass para desarrollo / Pruebas iniciales locales del usuario
    if (key === "TEST-123-KEY") {
        localStorage.setItem("pro-license-validated", "true");
        localStorage.setItem("pro-active-license-key", key);
        licenseLockScreen.classList.add("hidden");
        playSound('success');
        alert("¡Aplicación activada con éxito (Clave de prueba local)!");
        return;
    }

    // Si no ha configurado la URL aún
    if (GOOGLE_SCRIPT_URL === "INSERTA_AQUI_TU_URL_DE_GOOGLE_APPS_SCRIPT") {
        showActivationError("El servidor de licencias no está configurado. Configura GOOGLE_SCRIPT_URL en app.js o usa la clave 'TEST-123-KEY' para probar localmente.");
        return;
    }

    activateAppBtn.setAttribute("disabled", "true");
    activateAppBtn.textContent = "Validando clave...";
    activationErrorMsg.classList.add("hidden");

    const fetchUrl = `${GOOGLE_SCRIPT_URL}?key=${encodeURIComponent(key)}&device=${encodeURIComponent(deviceId)}`;

    fetch(fetchUrl)
        .then(response => response.json())
        .then(data => {
            activateAppBtn.removeAttribute("disabled");
            activateAppBtn.innerHTML = 'Activar Aplicación <i class="ti ti-key"></i>';

            if (data.success) {
                localStorage.setItem("pro-license-validated", "true");
                localStorage.setItem("pro-active-license-key", key);
                if (data.client) {
                    localStorage.setItem("pro-license-client", data.client);
                }
                licenseLockScreen.classList.add("hidden");
                playSound('success');
                alert("¡Aplicación activada y validada con éxito!");
            } else {
                showActivationError(data.message || "Error al validar la licencia.");
                playSound('error');
            }
        })
        .catch(err => {
            console.error(err);
            activateAppBtn.removeAttribute("disabled");
            activateAppBtn.innerHTML = 'Activar Aplicación <i class="ti ti-key"></i>';
            showActivationError("Error de conexión con el servidor. Verifica tu internet e inténtalo de nuevo.");
            playSound('error');
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
            
            const firstRowKeys = Object.keys(jsonData[0]);
            originalColumnsOrder = firstRowKeys;

            guestData = jsonData.map((row, index) => {
                return {
                    id: String(row["ID"] || row["id"] || index + 1).trim(),
                    name: String(row["Invitado"] || row["invitado"] || row["Nombre"] || row["nombre"] || "Sin Nombre").trim(),
                    quantity: parseInt(row["Cantidad"] || row["cantidad"] || row["Pases"] || row["pases"] || 1),
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
            
            // Enable QR Generator PRO
            generateQrsBtn.removeAttribute("disabled");
            qrGenMsg.innerHTML = `<i class="ti ti-circle-check text-success"></i> Lista cargada con <strong>${guestData.length}</strong> invitados listos para generar QRs.`;
            
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
            guest.qrValue.toLowerCase().includes(searchTerm);
            
        const isAttended = guest.attendance && guest.attendance.toLowerCase().trim() !== "";
        if (currentFilter === "attended") return matchesSearch && isAttended;
        if (currentFilter === "pending") return matchesSearch && !isAttended;
        return matchesSearch;
    });

    if (filteredList.length === 0) {
        guestTableBody.innerHTML = `
            <tr>
                <td colspan="6" class="table-empty">
                    <i class="ti ti-search-off"></i>
                    No se encontraron invitados.
                </td>
            </tr>
        `;
        return;
    }

    guestTableBody.innerHTML = filteredList.map(guest => {
        const hasAttended = guest.attendance && guest.attendance.toLowerCase().trim() !== "";
        const statusBadge = hasAttended 
            ? `<span class="badge success"><i class="ti ti-check"></i> ${guest.attendance}</span>` 
            : `<span class="badge pending">Pendiente</span>`;
            
        const actionButton = hasAttended
            ? `<button class="btn-action btn-secondary" onclick="toggleAttendance('${guest.id}', true)"><i class="ti ti-rotate-clockwise"></i> Revertir</button>`
            : `<button class="btn-action btn-checkin" onclick="toggleAttendance('${guest.id}', false)"><i class="ti ti-circle-check"></i> Registrar</button>`;

        return `
            <tr>
                <td><strong>${guest.id}</strong></td>
                <td>${guest.name}</td>
                <td>${guest.quantity}</td>
                <td><code>${guest.qrValue || guest.id}</code></td>
                <td>${statusBadge}</td>
                <td>${actionButton}</td>
            </tr>
        `;
    }).join('');
}

function toggleAttendance(id, revert = false) {
    const guestIndex = guestData.findIndex(g => g.id === id);
    if (guestIndex === -1) return;

    if (revert) {
        guestData[guestIndex].attendance = "";
        guestData[guestIndex].rawRow["Asistencia"] = "";
    } else {
        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        guestData[guestIndex].attendance = `Sí (${time})`;
        guestData[guestIndex].rawRow["Asistencia"] = `Sí (${time})`;
    }

    updateDashboard();
    renderGuestTable();
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
    
    toggleCameraBtn.innerHTML = `<i class="ti ti-camera-off"></i> Detener Escáner`;
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
    ).catch(err => {
        console.error(err);
        stopScanner();
        alert("No se pudo acceder a la cámara.");
    });
}

function stopScanner() {
    isScannerActive = false;
    toggleCameraBtn.innerHTML = `<i class="ti ti-camera"></i> Iniciar Escáner`;
    toggleCameraBtn.className = "btn-secondary";
    scannerOverlay.classList.add("hidden");

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
        displayResult('danger', 'Invitado No Encontrado', `El código QR "${qrVal}" no está registrado.`, qrVal);
        playSound('error');
        return;
    }

    const hasAttended = guest.attendance && guest.attendance.toLowerCase().trim() !== "";
    if (hasAttended) {
        displayResult('warning', 'Asistencia Ya Registrada', guest.name, guest.id, guest.quantity, guest.attendance);
        playSound('warning');
    } else {
        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const guestIndex = guestData.findIndex(g => g.id === guest.id);
        
        guestData[guestIndex].attendance = `Sí (${time})`;
        guestData[guestIndex].rawRow["Asistencia"] = `Sí (${time})`;
        
        updateDashboard();
        renderGuestTable();
        displayResult('success', '¡Registro Exitoso!', guest.name, guest.id, guest.quantity, `Sí (${time})`);
        playSound('success');
    }
}

function onQrCodeError(errorMessage) {}

function resetResultDisplay() {
    resultCard.className = "card result-card";
    resultPlaceholder.classList.remove("hidden");
    resultDetails.classList.add("hidden");
}

function displayResult(type, title, name, id = "-", qty = "-", time = "-") {
    resultPlaceholder.classList.add("hidden");
    resultDetails.classList.remove("hidden");
    resultCard.className = "card result-card " + type;
    
    resultTitle.textContent = title;
    resultName.textContent = name;
    resultId.textContent = id;
    resultQty.textContent = qty;
    resultTime.textContent = time;

    resultStatusIconWrapper.className = "result-status-icon-wrapper";
    if (type === 'success') {
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
