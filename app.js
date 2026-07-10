// Core Application Logic: QR Check-In with Excel Database

// Global state variables
let guestData = []; // Array of guest objects
let originalWorkbook = null; // Stored reference to the parsed workbook
let activeSheetName = ""; // Name of the active excel sheet
let html5QrScanner = null; // Scanner instance
let cameraList = []; // Array of available camera devices
let isScannerActive = false; // Scanner status tracker
let currentFilter = "all"; // 'all', 'attended', 'pending'
let originalColumnsOrder = []; // Track original order of columns in sheet

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

// Initialization
document.addEventListener("DOMContentLoaded", () => {
    setupFileLoaders();
    setupSearchAndFilters();
    setupCameraOptions();
    
    downloadExcelBtn.addEventListener("click", exportUpdatedExcel);
});

// Sound feedback synthesizer using Web Audio API
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
            // High clear beep
            osc.frequency.setValueAtTime(880, ctx.currentTime);
            gain.gain.setValueAtTime(0.1, ctx.currentTime);
            osc.start();
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.15);
            osc.stop(ctx.currentTime + 0.15);
        } else if (type === 'warning') {
            // Two medium alert beeps
            osc.frequency.setValueAtTime(587.33, ctx.currentTime); // D5
            gain.gain.setValueAtTime(0.1, ctx.currentTime);
            osc.start();
            
            // Beep 1
            gain.gain.setValueAtTime(0.1, ctx.currentTime);
            gain.gain.exponentialRampToValueAtTime(0.001, ctx.currentTime + 0.12);
            
            // Beep 2 (Simulated with frequency modification or standard delay)
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
            // Low error buzz
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

// File Drag & Drop Handlers
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

// Processing the Excel File
function handleExcelFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            originalWorkbook = workbook;
            activeSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[activeSheetName];
            
            // Parse Sheet to JSON
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
            
            if (jsonData.length === 0) {
                alert("El archivo Excel está vacío.");
                return;
            }
            
            // Map headers exactly to find appropriate fields
            const firstRowKeys = Object.keys(jsonData[0]);
            originalColumnsOrder = firstRowKeys;

            // Normalize fields into guest objects
            guestData = jsonData.map((row, index) => {
                return {
                    id: String(row["ID"] || row["id"] || index + 1).trim(),
                    name: String(row["Invitado"] || row["invitado"] || row["Nombre"] || row["nombre"] || "Sin Nombre").trim(),
                    quantity: parseInt(row["Cantidad"] || row["cantidad"] || row["Pases"] || row["pases"] || 1),
                    qrValue: String(row["QR"] || row["qr"] || row["Codigo"] || "").trim(),
                    attendance: String(row["Asistencia"] || row["asistencia"] || "").trim(),
                    rawRow: row // Save original data structure to preserve other columns
                };
            });

            // UI updates
            fileNameDisplay.textContent = file.name;
            dropZone.classList.add("hidden");
            fileInfo.classList.remove("hidden");
            fileStatusIndicator.classList.add("loaded");
            fileStatusIndicator.innerHTML = `<span class="dot"></span> Archivo cargado: ${file.name}`;
            
            searchInput.removeAttribute("disabled");
            toggleCameraBtn.removeAttribute("disabled");
            downloadExcelBtn.removeAttribute("disabled");
            
            updateDashboard();
            renderGuestTable();
            playSound('success');
            
        } catch (err) {
            console.error(err);
            alert("Error al leer el archivo Excel. Asegúrate de que sea un archivo de Excel válido (.xlsx o .xls).");
        }
    };
    reader.readAsArrayBuffer(file);
}

// Reset application state to initial
function resetAppState() {
    stopScanner();
    guestData = [];
    originalWorkbook = null;
    activeSheetName = "";
    originalColumnsOrder = [];
    
    // UI resets
    dropZone.classList.remove("hidden");
    fileInfo.classList.add("hidden");
    fileStatusIndicator.classList.remove("loaded");
    fileStatusIndicator.innerHTML = `<span class="dot"></span> Sin archivo cargado`;
    
    searchInput.setAttribute("disabled", "true");
    searchInput.value = "";
    toggleCameraBtn.setAttribute("disabled", "true");
    downloadExcelBtn.setAttribute("disabled", "true");
    
    updateDashboard();
    
    guestTableBody.innerHTML = `
        <tr>
            <td colspan="6" class="table-empty">
                <i class="ti ti-file-excel-off"></i>
                Carga un archivo Excel para ver la lista de invitados.
            </td>
        </tr>
    `;
    
    // Reset scanner viewport layout
    toggleCameraBtn.innerHTML = `<i class="ti ti-camera"></i> Iniciar Escáner`;
    toggleCameraBtn.className = "btn-secondary";
    scannerOverlay.classList.add("hidden");
    
    // Reset result panel
    resetResultDisplay();
}

// Update the Statistics Dashboard
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

// Render the Guest Table with current filters
function renderGuestTable() {
    if (guestData.length === 0) return;
    
    const searchTerm = searchInput.value.toLowerCase().trim();
    
    // Filter logic
    let filteredList = guestData.filter(guest => {
        // Search matches Name, ID or QR value
        const matchesSearch = 
            guest.name.toLowerCase().includes(searchTerm) || 
            guest.id.toLowerCase().includes(searchTerm) || 
            guest.qrValue.toLowerCase().includes(searchTerm);
            
        // Filter tabs matches
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
            ? `<button class="btn-action btn-secondary" onclick="toggleAttendance('${guest.id}', true)" title="Desmarcar asistencia"><i class="ti ti-rotate-clockwise"></i> Revertir</button>`
            : `<button class="btn-action btn-checkin" onclick="toggleAttendance('${guest.id}', false)"><i class="ti ti-circle-check"></i> Registrar</button>`;

        return `
            <tr>
                <td><strong>${guest.id}</strong></td>
                <td>${guest.name}</td>
                <td>${guest.quantity}</td>
                <td><code class="qr-code-text">${guest.qrValue || guest.id}</code></td>
                <td>${statusBadge}</td>
                <td>${actionButton}</td>
            </tr>
        `;
    }).join('');
}

// Toggle attendance status manually or via scan
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

// Setup search box and status filter tabs
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

// Setup camera devices listing
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
            console.warn("No se detectaron cámaras en el dispositivo.");
        }
    }).catch(err => {
        console.error("Error obteniendo cámaras:", err);
        cameraSelect.innerHTML = `<option value="">Error de permisos de cámara</option>`;
    });
}

// Scanner Toggle Switch
function toggleScanner() {
    if (isScannerActive) {
        stopScanner();
    } else {
        startScanner();
    }
}

// Start QR scanner sequence
function startScanner() {
    const cameraId = cameraSelect.value;
    if (!cameraId) {
        alert("Por favor selecciona una cámara.");
        return;
    }

    resetResultDisplay();
    isScannerActive = true;
    
    // UI status updating
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
        console.error("Error al iniciar lector QR:", err);
        stopScanner();
        alert("No se pudo acceder a la cámara seleccionada.");
    });
}

// Stop scanner sequence
function stopScanner() {
    isScannerActive = false;
    toggleCameraBtn.innerHTML = `<i class="ti ti-camera"></i> Iniciar Escáner`;
    toggleCameraBtn.className = "btn-secondary";
    scannerOverlay.classList.add("hidden");

    if (html5QrScanner) {
        html5QrScanner.stop().then(() => {
            html5QrScanner = null;
        }).catch(err => {
            console.error("Error al apagar escáner:", err);
        });
    }
}

// Scanned Success Handler
function onQrCodeSuccess(decodedText) {
    // Process text. QR can contain a full ID, or a URL ending with the ID, or a custom string.
    let qrVal = decodedText.trim();
    
    // Look for matching guest by either QR string, or ID directly
    let guest = guestData.find(g => 
        (g.qrValue && g.qrValue.toLowerCase() === qrVal.toLowerCase()) || 
        g.id.toLowerCase() === qrVal.toLowerCase()
    );

    // If QR contains a URL, try to extract ID parameter as fallback
    if (!guest) {
        try {
            const url = new URL(decodedText);
            const idParam = url.searchParams.get("id");
            if (idParam) {
                guest = guestData.find(g => g.id.toLowerCase() === idParam.toLowerCase().trim());
            }
        } catch(e) {
            // Decoded text is not a URL, which is fine
        }
    }

    if (!guest) {
        // Guest not found
        displayResult('danger', 'Invitado No Encontrado', `El código QR "${qrVal}" no está registrado.`, qrVal);
        playSound('error');
        return;
    }

    const hasAttended = guest.attendance && guest.attendance.toLowerCase().trim() !== "";
    if (hasAttended) {
        // Guest already checked-in
        displayResult('warning', 'Asistencia Ya Registrada', guest.name, guest.id, guest.quantity, guest.attendance);
        playSound('warning');
    } else {
        // Registering guest check-in
        const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        
        // Update data
        const guestIndex = guestData.findIndex(g => g.id === guest.id);
        guestData[guestIndex].attendance = `Sí (${time})`;
        guestData[guestIndex].rawRow["Asistencia"] = `Sí (${time})`;
        
        // Render updates
        updateDashboard();
        renderGuestTable();
        
        // Display result
        displayResult('success', '¡Registro Exitoso!', guest.name, guest.id, guest.quantity, `Sí (${time})`);
        playSound('success');
    }
}

// Scanned Error Handler (Ignored to avoid spam console log on each frame)
function onQrCodeError(errorMessage) {
    // Console spam prevention
}

// Reset the scan result details visual card
function resetResultDisplay() {
    resultCard.className = "card result-card";
    resultPlaceholder.classList.remove("hidden");
    resultDetails.classList.add("hidden");
}

// Display scan result details
function displayResult(type, title, name, id = "-", qty = "-", time = "-") {
    resultPlaceholder.classList.add("hidden");
    resultDetails.classList.remove("hidden");
    
    // Clear styles
    resultCard.className = "card result-card " + type;
    
    resultTitle.textContent = title;
    resultName.textContent = name;
    resultId.textContent = id;
    resultQty.textContent = qty;
    resultTime.textContent = time;

    // Update status icon
    resultStatusIconWrapper.className = "result-status-icon-wrapper";
    if (type === 'success') {
        resultIcon.className = "ti ti-circle-check";
    } else if (type === 'warning') {
        resultIcon.className = "ti ti-alert-triangle";
    } else {
        resultIcon.className = "ti ti-circle-x";
    }
}

// Export the updated guest records back into the Excel spreadsheet
function exportUpdatedExcel() {
    if (guestData.length === 0) return;

    // Reconstruct the spreadsheet row structure exactly as it was originally loaded
    // preserving any extra/custom columns from the user's template
    const outputRows = guestData.map(guest => {
        const row = { ...guest.rawRow };
        // Sync values to the row
        // Locate matching keys to avoid creating duplicate columns (like lowercase "asistencia" if it exists)
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
        
        // Create a new workbook or modify the existing one
        const newWorkbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, activeSheetName);
        
        // Write file and trigger browser download dialog
        XLSX.writeFile(newWorkbook, "RS_Registrado.xlsx");
        
        playSound('success');
    } catch(err) {
        console.error("Error al exportar Excel:", err);
        alert("Hubo un error al exportar el archivo Excel. Por favor reintenta.");
    }
}
