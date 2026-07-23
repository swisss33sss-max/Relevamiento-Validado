const JSON_DB_FILE = 'equipos.json';

let jsonDb = []; // Primary JSON Database

// DOM Elements
const sheetSelect = document.getElementById('sheetSelect');
const searchType = document.getElementById('searchType');
const searchInput = document.getElementById('searchInput');
const equipmentList = document.getElementById('equipmentList');
const resultsArea = document.getElementById('resultsArea');
const detailsGrid = document.getElementById('detailsGrid');
const equipmentImage = document.getElementById('equipmentImage');
const noImageText = document.getElementById('noImageText');
const localFileWarning = document.getElementById('localFileWarning');
const manualFileInput = document.getElementById('manualFileInput');
const manualCertInput = document.getElementById('manualCertInput');
const manualInventarioInput = document.getElementById('manualInventarioInput');
const certButton = document.getElementById('certButton');
const certInventarioButton = document.getElementById('certInventarioButton');
const inventarioDataSection = document.getElementById('inventarioDataSection');
const inventarioDetailsGrid = document.getElementById('inventarioDetailsGrid');
const addCertSection = document.getElementById('addCertSection');
const qrCodeContainer = document.getElementById('qrCodeContainer');
const qrcodeElement = document.getElementById('qrcode');
let qrCodeInstance = null;
const manualLinkInput = document.getElementById('manualLinkInput');
const saveManualLinkBtn = document.getElementById('saveManualLinkBtn');
const extraDataView = document.getElementById('extraDataView');
const extraDataEdit = document.getElementById('extraDataEdit');
const toggleEditDataBtn = document.getElementById('toggleEditDataBtn');
const saveExtraDataBtn = document.getElementById('saveExtraDataBtn');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const exportJsonBtn = document.getElementById('exportJsonBtn');
const viewFabricante = document.getElementById('viewFabricante');
const viewModelo = document.getElementById('viewModelo');
const viewDescripcion = document.getElementById('viewDescripcion');
const inputFabricante = document.getElementById('inputFabricante');
const inputModelo = document.getElementById('inputModelo');
const inputDescripcion = document.getElementById('inputDescripcion');

// Modal & Image Elements
const addEquipmentModalBtn = document.getElementById('addEquipmentModalBtn');
const addModal = document.getElementById('addModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const cancelModalBtn = document.getElementById('cancelModalBtn');
const addEquipmentForm = document.getElementById('addEquipmentForm');
const newImagenFile = document.getElementById('newImagenFile');
const newImagen = document.getElementById('newImagen');
let uploadedImageBase64 = "";

if (newImagenFile) {
    newImagenFile.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            newImagen.value = file.name;
            const reader = new FileReader();
            reader.onload = (evt) => {
                uploadedImageBase64 = evt.target.result;
            };
            reader.readAsDataURL(file);
        }
    });
}

// Mobile Navigation Drawer
const mobileMenuToggle = document.getElementById('mobileMenuToggle');
const closeSidebarBtn = document.getElementById('closeSidebarBtn');
const sidebarOverlay = document.getElementById('sidebarOverlay');
const sidebarPanel = document.getElementById('sidebarPanel');
const resultCount = document.getElementById('resultCount');

function openSidebar() {
    if (sidebarPanel) sidebarPanel.classList.add('open');
    if (sidebarOverlay) sidebarOverlay.classList.add('active');
}

function closeSidebar() {
    if (sidebarPanel) sidebarPanel.classList.remove('open');
    if (sidebarOverlay) sidebarOverlay.classList.remove('active');
}

if (mobileMenuToggle) mobileMenuToggle.addEventListener('click', openSidebar);
if (closeSidebarBtn) closeSidebarBtn.addEventListener('click', closeSidebar);
if (sidebarOverlay) sidebarOverlay.addEventListener('click', closeSidebar);

// Initialize Application (CORS-Safe)
async function init() {
    // 1. Try preloaded data from window.INITIAL_EQUIPOS_DB (Bypasses file:// CORS restriction)
    if (window.INITIAL_EQUIPOS_DB && Array.isArray(window.INITIAL_EQUIPOS_DB) && window.INITIAL_EQUIPOS_DB.length > 0) {
        jsonDb = [...window.INITIAL_EQUIPOS_DB];
        console.log(`Cargados ${jsonDb.length} equipos desde window.INITIAL_EQUIPOS_DB`);
    } else {
        // 2. Fallback to fetch equipos.json if running on HTTP server
        try {
            const response = await fetch(JSON_DB_FILE);
            if (response.ok) {
                jsonDb = await response.json();
            }
        } catch (e) {
            console.warn("Navegación local file:// detectada.");
        }
    }

    // 3. Check localStorage for custom added/updated items
    const localDbStr = localStorage.getItem('equipos_custom_db');
    let localDb = [];
    if (localDbStr) {
        try { localDb = JSON.parse(localDbStr); } catch(e) {}
    }

    if (localDb && localDb.length > 0) {
        localDb.forEach(item => {
            const idx = jsonDb.findIndex(x => 
                (x.serie && item.serie && x.serie.toLowerCase() === item.serie.toLowerCase()) ||
                (x.equipo && item.equipo && x.equipo.toLowerCase() === item.equipo.toLowerCase())
            );
            if (idx !== -1) {
                jsonDb[idx] = { ...jsonDb[idx], ...item };
            } else {
                jsonDb.unshift(item);
            }
        });
    }

    if (jsonDb.length > 0) {
        populateSheetDropdownFromDb();
        filterAndRenderDb();
    }
}

function populateSheetDropdownFromDb() {
    if (!sheetSelect) return;
    sheetSelect.innerHTML = '<option value="TODAS">Todas las Hojas</option>';
    const sheets = setOfSheetsFromDb();
    sheets.forEach(s => {
        const opt = document.createElement('option');
        opt.value = s;
        opt.textContent = s;
        sheetSelect.appendChild(opt);
    });
}

function setOfSheetsFromDb() {
    const set = new Set();
    jsonDb.forEach(item => {
        if (item.hoja) set.add(item.hoja);
    });
    return Array.from(set);
}

function filterAndRenderDb() {
    const selectedSheet = sheetSelect.value;
    const sTerm = (searchInput.value || "").trim().toLowerCase();
    const type = searchType.value;

    let filtered = jsonDb;

    if (selectedSheet && selectedSheet !== 'TODAS') {
        filtered = filtered.filter(x => x.hoja === selectedSheet);
    }

    if (sTerm) {
        filtered = filtered.filter(x => {
            const eqVal = (x.equipo || "").toLowerCase();
            const serVal = (x.serie || "").toLowerCase();

            if (type === 'equipo') return eqVal.includes(sTerm);
            if (type === 'serie') return serVal.includes(sTerm);
            return eqVal.includes(sTerm) || serVal.includes(sTerm);
        });
    }

    renderDbList(filtered, sTerm);
}

function renderDbList(items, highlightTerm = "") {
    equipmentList.innerHTML = '';
    const type = searchType.value;
    const isSerie = type === 'serie';

    if (resultCount) {
        resultCount.textContent = items.length;
        resultCount.classList.remove('hidden');
    }

    items.forEach(item => {
        let mainVal = isSerie ? (item.serie || item.equipo) : (item.equipo || item.serie);
        let subVal = isSerie ? item.equipo : item.serie;

        if (!mainVal) mainVal = "(Sin datos)";

        const li = document.createElement('li');
        let displayHTML = mainVal;
        let subHTMLText = subVal || "";

        if (highlightTerm) {
            const regex = new RegExp(`(${highlightTerm})`, "gi");
            displayHTML = displayHTML.replace(regex, "<strong>$1</strong>");
            if (subHTMLText) subHTMLText = subHTMLText.replace(regex, "<strong>$1</strong>");
        }

        let subHTML = '';
        if (subHTMLText) {
            subHTML = `<span class="sub-info">${isSerie ? 'Equipo: ' : 'Serie: '} ${subHTMLText}</span>`;
        }

        li.innerHTML = displayHTML + subHTML;

        li.addEventListener('click', () => {
            document.querySelectorAll('.equipment-list li').forEach(el => el.classList.remove('active'));
            li.classList.add('active');
            showDbResults(item);
            closeSidebar();
        });

        equipmentList.appendChild(li);
    });
}

function showDbResults(item) {
    resultsArea.classList.remove('hidden');
    detailsGrid.innerHTML = '';

    const mainDetails = item.detalles || { "Equipo": item.equipo, "Número de Serie": item.serie, "Hoja": item.hoja };
    
    Object.keys(mainDetails).forEach(key => {
        const val = mainDetails[key];
        if (val !== undefined && val !== "") {
            let valueHTML = `<span class="detail-value">${val}</span>`;
            if (String(val).trim().startsWith('http')) {
                valueHTML = `<a href="${String(val).trim()}" target="_blank" rel="noopener" style="background: rgba(59, 130, 246, 0.2); color: #60a5fa; padding: 0.2rem 0.6rem; border-radius: 6px; text-decoration: none; font-size: 0.85rem; font-weight: bold; border: 1px solid #3b82f6; display: inline-block;">🔗 Abrir Enlace</a>`;
            }
            const div = document.createElement('div');
            div.className = 'detail-item';
            div.innerHTML = `<span class="detail-label">${key}</span>${valueHTML}`;
            detailsGrid.appendChild(div);
        }
    });

    let fabricante = item.fabricante || "";
    let modelo = item.modelo || "";
    let descripcion = item.descripcion || "";

    viewFabricante.textContent = fabricante || "-";
    viewModelo.textContent = modelo || "-";
    viewDescripcion.textContent = descripcion || "-";

    extraDataView.classList.remove('hidden');
    extraDataEdit.classList.add('hidden');
    toggleEditDataBtn.textContent = '✏️ Editar';

    toggleEditDataBtn.onclick = () => {
        const isHidden = extraDataEdit.classList.contains('hidden');
        if (isHidden) {
            extraDataView.classList.add('hidden');
            extraDataEdit.classList.remove('hidden');
            toggleEditDataBtn.textContent = '❌ Cancelar';
            inputFabricante.value = fabricante;
            inputModelo.value = modelo;
            inputDescripcion.value = descripcion;
        } else {
            extraDataView.classList.remove('hidden');
            extraDataEdit.classList.add('hidden');
            toggleEditDataBtn.textContent = '✏️ Editar';
        }
    };

    saveExtraDataBtn.onclick = () => {
        item.fabricante = inputFabricante.value.trim();
        item.modelo = inputModelo.value.trim();
        item.descripcion = inputDescripcion.value.trim();

        saveDbToLocal();

        viewFabricante.textContent = item.fabricante || "-";
        viewModelo.textContent = item.modelo || "-";
        viewDescripcion.textContent = item.descripcion || "-";

        extraDataView.classList.remove('hidden');
        extraDataEdit.classList.add('hidden');
        toggleEditDataBtn.textContent = '✏️ Editar';
    };

    let certLink = item.linkCertificado || "";
    if (certLink) {
        certButton.href = certLink;
        certButton.classList.remove('hidden');
        addCertSection.classList.add('hidden');
    } else {
        certButton.classList.add('hidden');
        addCertSection.classList.remove('hidden');
        manualLinkInput.value = '';

        saveManualLinkBtn.onclick = () => {
            const newLink = manualLinkInput.value.trim();
            if (newLink) {
                item.linkCertificado = newLink;
                saveDbToLocal();
                certButton.href = newLink;
                certButton.classList.remove('hidden');
                addCertSection.classList.add('hidden');
            }
        };
    }

    // QR Code
    qrCodeContainer.classList.add('hidden');
    qrcodeElement.innerHTML = '';
    if (certLink && typeof QRCode !== 'undefined') {
        qrCodeContainer.classList.remove('hidden');
        if (qrCodeInstance) {
            qrCodeInstance.clear();
            qrCodeInstance.makeCode(certLink);
        } else {
            qrCodeInstance = new QRCode(qrcodeElement, {
                text: certLink,
                width: 150,
                height: 150,
                colorDark: "#000000",
                colorLight: "#ffffff",
                correctLevel: QRCode.CorrectLevel.M
            });
        }
    }

    // Equipment Image Processing
    const imageName = item.imagen;
    equipmentImage.onclick = null;
    if (imageName && String(imageName).trim() !== "") {
        let imgStr = String(imageName).trim();
        let imgPath = imgStr;
        
        if (!imgStr.startsWith('data:image') && !imgStr.toLowerCase().startsWith('http')) {
            imgPath = `Imagenes/${imgStr}`;
            if (!imgPath.toLowerCase().match(/\.(jpg|jpeg|png|gif|webp)$/)) {
                imgPath += '.jpg';
            }
        }
        equipmentImage.src = imgPath;
        equipmentImage.classList.remove('hidden');
        noImageText.classList.add('hidden');
    } else {
        equipmentImage.classList.add('hidden');
        noImageText.classList.remove('hidden');
    }
}

function saveDbToLocal() {
    try {
        localStorage.setItem('equipos_custom_db', JSON.stringify(jsonDb));
    } catch(e) {}
}

// Modal Handlers for Adding New Equipment
if (addEquipmentModalBtn) {
    addEquipmentModalBtn.addEventListener('click', () => {
        uploadedImageBase64 = "";
        if (addModal) addModal.classList.remove('hidden');
    });
}

if (closeModalBtn) {
    closeModalBtn.addEventListener('click', () => {
        if (addModal) addModal.classList.add('hidden');
    });
}

if (cancelModalBtn) {
    cancelModalBtn.addEventListener('click', () => {
        if (addModal) addModal.classList.add('hidden');
    });
}

if (addEquipmentForm) {
    addEquipmentForm.addEventListener('submit', (e) => {
        e.preventDefault();

        const equipo = document.getElementById('newEquipo').value.trim();
        const serie = document.getElementById('newSerie').value.trim();
        const fabricante = document.getElementById('newFabricante').value.trim();
        const modelo = document.getElementById('newModelo').value.trim();
        const hoja = document.getElementById('newHoja').value.trim().toUpperCase() || 'GENERAL';
        const imagenText = document.getElementById('newImagen').value.trim();
        const imagen = uploadedImageBase64 || imagenText;
        const linkCertificado = document.getElementById('newCertificado').value.trim();
        const descripcion = document.getElementById('newDescripcion').value.trim();

        const newItem = {
            hoja,
            equipo,
            serie,
            imagen,
            fabricante,
            modelo,
            descripcion,
            linkCertificado,
            detalles: {
                "Equipo": equipo,
                "Número de Serie": serie,
                "Hoja": hoja,
                "Fabricante": fabricante,
                "Modelo": modelo
            }
        };

        const existingIdx = jsonDb.findIndex(x => x.serie && x.serie.toLowerCase() === serie.toLowerCase());
        if (existingIdx !== -1) {
            jsonDb[existingIdx] = { ...jsonDb[existingIdx], ...newItem };
        } else {
            jsonDb.unshift(newItem);
        }

        saveDbToLocal();
        populateSheetDropdownFromDb();
        filterAndRenderDb();

        addEquipmentForm.reset();
        uploadedImageBase64 = "";
        if (addModal) addModal.classList.add('hidden');

        alert(`¡Equipo '${equipo}' (${serie}) guardado correctamente en la base de datos!`);
    });
}

// Export equipos.json File for Git Commit
if (exportJsonBtn) {
    exportJsonBtn.addEventListener('click', () => {
        const jsonStr = JSON.stringify(jsonDb, null, 2);
        const blob = new Blob([jsonStr], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'equipos.json';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    });
}

// Manual Excel File Inputs Listener (For updating DB from Excel)
if (manualFileInput) {
    manualFileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            processWorkbook(wb);
        };
        reader.readAsArrayBuffer(file);
    });
}

if (manualCertInput) {
    manualCertInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            processCertWorkbook(wb);
            const msg = document.getElementById('certLoadedMsg');
            if (msg) msg.classList.remove('hidden');
        };
        reader.readAsArrayBuffer(file);
    });
}

if (manualInventarioInput) {
    manualInventarioInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function(evt) {
            const data = new Uint8Array(evt.target.result);
            const wb = XLSX.read(data, { type: 'array' });
            processInventarioWorkbook(wb);
            const msg = document.getElementById('inventarioLoadedMsg');
            if (msg) msg.classList.remove('hidden');
        };
        reader.readAsArrayBuffer(file);
    });
}

function processWorkbook(wb) {
    wb.SheetNames.forEach(sheetName => {
        const worksheet = wb.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        if (data.length > 0) {
            const h = data[0] || [];
            const eqIdx = h.findIndex(x => String(x).toLowerCase().includes('equipo'));
            const serIdx = h.findIndex(x => String(x).toLowerCase().includes('serie') || String(x).toLowerCase().includes('sn'));
            const imgIdx = h.findIndex(x => String(x).toLowerCase().includes('imagen'));

            for (let i = 1; i < data.length; i++) {
                const r = data[i];
                if (!r) continue;
                const eq = String(r[eqIdx >= 0 ? eqIdx : 0] || "").trim();
                const ser = String(r[serIdx >= 0 ? serIdx : 2] || "").trim();
                const img = String(r[imgIdx >= 0 ? imgIdx : 6] || "").trim();

                if (eq || ser) {
                    const detalles = {};
                    h.forEach((headerName, colIdx) => {
                        if (headerName && r[colIdx]) detalles[headerName] = r[colIdx];
                    });
                    
                    const newItem = {
                        hoja: sheetName,
                        equipo: eq,
                        serie: ser,
                        imagen: img,
                        fabricante: "",
                        modelo: "",
                        descripcion: "",
                        linkCertificado: "",
                        detalles
                    };

                    const idx = jsonDb.findIndex(x => x.serie && ser && x.serie.toLowerCase() === ser.toLowerCase());
                    if (idx !== -1) {
                        jsonDb[idx] = { ...jsonDb[idx], ...newItem };
                    } else {
                        jsonDb.push(newItem);
                    }
                }
            }
        }
    });

    saveDbToLocal();
    populateSheetDropdownFromDb();
    filterAndRenderDb();
    alert("¡Lista de equipos actualizada en la base de datos!");
}

function processCertWorkbook(wb) {
    const sheetName = wb.SheetNames[0];
    const worksheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;

        const c_eq = String(row[1] || "").trim().toLowerCase();
        const c_ser = String(row[3] || "").trim().toLowerCase();
        const c_tipo = String(row[2] || "").trim();
        const c_fab = String(row[4] || "").trim();
        const c_mod = String(row[5] || "").trim();
        const c_link = String(row[6] || "").trim();

        jsonDb.forEach(eq => {
            const m_eq = (eq.equipo || "").trim().toLowerCase();
            const m_ser = (eq.serie || "").trim().toLowerCase();

            if ((c_ser && m_ser && c_ser === m_ser) || (c_eq && m_eq && c_eq === m_eq)) {
                if (c_fab) eq.fabricante = c_fab;
                if (c_mod) eq.modelo = c_mod;
                if (c_link) eq.linkCertificado = c_link;
            }
        });
    }

    saveDbToLocal();
    filterAndRenderDb();
}

function processInventarioWorkbook(wb) {
    const sheetName = wb.SheetNames[0];
    const worksheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    const headers = data[0] || [];

    let invEqIndex = headers.findIndex(h => h && String(h).toLowerCase().includes('equipo'));
    if (invEqIndex === -1) invEqIndex = 1;
    let invSerIndex = headers.findIndex(h => h && String(h).toLowerCase().includes('serie'));
    if (invSerIndex === -1) invSerIndex = 3;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        const invEq = String(row[invEqIndex] || "").trim().toLowerCase();
        const invSer = String(row[invSerIndex] || "").trim().toLowerCase();

        jsonDb.forEach(eq => {
            const m_eq = (eq.equipo || "").trim().toLowerCase();
            const m_ser = (eq.serie || "").trim().toLowerCase();

            if ((invSer && m_ser && invSer === m_ser) || (invEq && m_eq && invEq === m_eq)) {
                headers.forEach((hName, colIdx) => {
                    if (hName && row[colIdx]) {
                        if (!eq.detalles) eq.detalles = {};
                        eq.detalles[`[Inv] ${hName}`] = row[colIdx];
                    }
                });
            }
        });
    }

    saveDbToLocal();
    filterAndRenderDb();
}

// Search and Filter Listeners
if (sheetSelect) sheetSelect.addEventListener('change', filterAndRenderDb);
if (searchType) searchType.addEventListener('change', filterAndRenderDb);
if (searchInput) searchInput.addEventListener('input', filterAndRenderDb);

// Export Consolidate Excel
if (exportExcelBtn) {
    exportExcelBtn.addEventListener('click', () => {
        const exportData = [];
        if (jsonDb.length > 0) {
            exportData.push(["Hoja", "Equipo", "Serie", "Fabricante", "Modelo", "Imagen", "Link Certificado", "Descripción"]);
            jsonDb.forEach(x => {
                exportData.push([x.hoja || "", x.equipo || "", x.serie || "", x.fabricante || "", x.modelo || "", x.imagen || "", x.linkCertificado || "", x.descripcion || ""]);
            });
        }
        const ws = XLSX.utils.aoa_to_sheet(exportData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Base_Datos_Consolidada");
        const dateString = new Date().toISOString().split('T')[0];
        XLSX.writeFile(wb, `Base_Datos_Equipos_${dateString}.xlsx`);
    });
}

// Start Application
window.addEventListener('DOMContentLoaded', init);
