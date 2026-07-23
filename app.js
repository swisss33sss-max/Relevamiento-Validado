const EXCEL_FILE = 'IC-SM-CAT-LAB LISTA CATAMARCA LAB  --  2.xlsx';
const CERT_FILE = 'Etiquetas Laboratorio.xls';
const INVENTARIO_FILE = 'Inventario Catamarca.xlsx';

let workbook = null;
let sheetsData = {}; 
let currentSheetData = [];
let headers = [];
let certData = [];
let certHeaders = [];
let certRawData = [];

let inventarioData = [];
let inventarioHeaders = [];
let inventarioRawData = [];

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
const cleanDuplicatesBtn = document.getElementById('cleanDuplicatesBtn');
const extraDataView = document.getElementById('extraDataView');
const extraDataEdit = document.getElementById('extraDataEdit');
const toggleEditDataBtn = document.getElementById('toggleEditDataBtn');
const saveExtraDataBtn = document.getElementById('saveExtraDataBtn');
const exportExcelBtn = document.getElementById('exportExcelBtn');
const viewFabricante = document.getElementById('viewFabricante');
const viewModelo = document.getElementById('viewModelo');
const viewDescripcion = document.getElementById('viewDescripcion');
const inputFabricante = document.getElementById('inputFabricante');
const inputModelo = document.getElementById('inputModelo');
const inputDescripcion = document.getElementById('inputDescripcion');

// Responsive UI elements
const mobileMenuToggle = document.getElementById('mobileMenuToggle');
const closeSidebarBtn = document.getElementById('closeSidebarBtn');
const sidebarOverlay = document.getElementById('sidebarOverlay');
const sidebarPanel = document.getElementById('sidebarPanel');
const resultCount = document.getElementById('resultCount');

// Drawer helpers for mobile devices
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

// Initialization
async function init() {
    try {
        const response = await fetch(EXCEL_FILE);
        const arrayBuffer = await response.arrayBuffer();
        
        const wb = XLSX.read(arrayBuffer, { type: 'array' });
        processWorkbook(wb);

        // Try to load certificates file
        try {
            const certRes = await fetch(CERT_FILE);
            const certBuf = await certRes.arrayBuffer();
            const certWb = XLSX.read(certBuf, { type: 'array' });
            processCertWorkbook(certWb);
            const msg = document.getElementById('certLoadedMsg');
            if (msg) msg.classList.remove('hidden');
        } catch (e) {
            console.warn("No se pudo cargar 'Etiquetas Laboratorio.xls' automáticamente.", e);
        }

        // Try to load Inventario file
        try {
            const invRes = await fetch(INVENTARIO_FILE);
            const invBuf = await invRes.arrayBuffer();
            const invWb = XLSX.read(invBuf, { type: 'array' });
            processInventarioWorkbook(invWb);
            const msg = document.getElementById('inventarioLoadedMsg');
            if (msg) msg.classList.remove('hidden');
        } catch (e) {
            console.warn("No se pudo cargar 'Inventario Catamarca.xlsx' automáticamente.", e);
        }

    } catch (error) {
        console.error("Error loading Excel file:", error);
        if (window.location.protocol === 'file:') {
            // Show fallback for local file:// execution
            if (localFileWarning) localFileWarning.classList.remove('hidden');
            if (sheetSelect) sheetSelect.innerHTML = '<option value="">Requiere archivo manual...</option>';
        } else {
            if (sheetSelect) sheetSelect.innerHTML = '<option value="">Error cargando archivo</option>';
        }
    }
}

function processWorkbook(wb) {
    workbook = wb;
    sheetSelect.innerHTML = '';
    sheetsData = {};
    
    workbook.SheetNames.forEach(sheetName => {
        const option = document.createElement('option');
        option.value = sheetName;
        option.textContent = sheetName;
        sheetSelect.appendChild(option);
        
        const worksheet = workbook.Sheets[sheetName];
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        sheetsData[sheetName] = data;
    });

    sheetSelect.disabled = false;
    searchInput.disabled = false;
    if (localFileWarning) localFileWarning.classList.add('hidden');
    loadSheet(sheetSelect.value);
}

function processCertWorkbook(wb) {
    certData = [];
    certRawData = [];
    const sheetName = wb.SheetNames[0];
    const worksheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    
    if (data.length > 0) {
        certHeaders = data[0];
    }
    
    // Start from row 1 to skip headers
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        certRawData.push(row);
        
        const equipo = String(row[1] || "").trim().toLowerCase(); // Col B
        const serie = String(row[3] || "").trim().toLowerCase();  // Col D
        const link = String(row[6] || "").trim();                 // Col G
        const tipo = String(row[2] || "").trim(); // Col C
        const fabricante = String(row[4] || "").trim(); // Col E
        const modelo = String(row[5] || "").trim(); // Col F
        
        if (equipo || serie) {
            certData.push({ equipo, serie, link, tipo, fabricante, modelo, rowIndex: i - 1 });
        }
    }
}

function processInventarioWorkbook(wb) {
    inventarioData = [];
    inventarioRawData = [];
    const sheetName = wb.SheetNames[0];
    const worksheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    
    if (data.length > 0) {
        inventarioHeaders = data[0];
    }
    
    // Determine indices for "Equipo" and "Serie" dynamically
    let invEqIndex = inventarioHeaders.findIndex(h => h && String(h).toLowerCase().includes('equipo'));
    if (invEqIndex === -1) invEqIndex = 1; // Default to 'Campo/equipo'
    let invSerIndex = inventarioHeaders.findIndex(h => h && String(h).toLowerCase().includes('serie'));
    if (invSerIndex === -1) invSerIndex = 3; // Default to 'N Serie'

    let invLinkIndex = inventarioHeaders.findIndex(h => {
        const lowerH = String(h).toLowerCase();
        return lowerH.includes('link') || lowerH.includes('url') || lowerH.includes('certificado') || lowerH.includes('documento');
    });

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        inventarioRawData.push(row);
        
        const equipo = String(row[invEqIndex] || "").trim().toLowerCase();
        const serie = String(row[invSerIndex] || "").trim().toLowerCase();
        
        let link = "";
        if (invLinkIndex !== -1) {
            link = String(row[invLinkIndex] || "").trim();
        } else {
            // Check all columns for something starting with http
            for (let c = 0; c < row.length; c++) {
                const val = String(row[c] || "").trim();
                if (val.startsWith('http')) {
                    link = val;
                    break;
                }
            }
        }
        
        if (equipo || serie) {
            inventarioData.push({ equipo, serie, link, rowData: row, rowIndex: i - 1 });
        }
    }
}

// Event Listeners for File Inputs
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
        
        const activeLi = document.querySelector('.equipment-list li.active');
        if (activeLi) activeLi.click(); // re-click to refresh details
    };
    reader.readAsArrayBuffer(file);
});

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
        
        const activeLi = document.querySelector('.equipment-list li.active');
        if (activeLi) activeLi.click(); // re-click to refresh details
    };
    reader.readAsArrayBuffer(file);
});

// Dynamic column indices for the main list
let eqIndex = 0;
let serIndex = 2;
let imgIndex = 6;

function loadSheet(sheetName) {
    const data = sheetsData[sheetName];
    if (!data || data.length === 0) return;

    headers = data[0] || [];
    currentSheetData = data.slice(1); // skip headers
    
    // Dynamically find indices based on headers
    eqIndex = headers.findIndex(h => h && String(h).toLowerCase().includes('equipo'));
    if (eqIndex === -1) eqIndex = 0;
    
    serIndex = headers.findIndex(h => h && String(h).toLowerCase().includes('serie'));
    if (serIndex === -1) serIndex = headers.findIndex(h => h && String(h).toLowerCase().includes('sn'));
    if (serIndex === -1) serIndex = 2;
    
    imgIndex = headers.findIndex(h => h && String(h).toLowerCase() === 'imagen');
    if (imgIndex === -1) imgIndex = headers.findIndex(h => h && String(h).toLowerCase().includes('imagen'));
    if (imgIndex === -1) imgIndex = 6;
    
    // Clear search and results
    searchInput.value = '';
    resultsArea.classList.add('hidden');
    renderList(currentSheetData);
}

sheetSelect.addEventListener('change', (e) => {
    loadSheet(e.target.value);
});

searchType.addEventListener('change', () => {
    const val = searchInput.value;
    filterList(val);
});

searchInput.addEventListener('input', function() {
    const val = this.value;
    filterList(val);
});

function filterList(searchTerm) {
    if (!searchTerm) {
        renderList(currentSheetData);
        return;
    }
    
    const type = searchType.value;
    
    const matches = currentSheetData.filter(row => {
        const valA = row[eqIndex] ? String(row[eqIndex]).toLowerCase() : "";
        const valC = row[serIndex] ? String(row[serIndex]).toLowerCase() : "";
        const s = searchTerm.toLowerCase();
        
        if (type === 'equipo') {
            return valA.includes(s);
        } else if (type === 'serie') {
            return valC.includes(s);
        } else {
            return valA.includes(s) || valC.includes(s);
        }
    });
    
    renderList(matches, searchTerm);
}

function renderList(rows, highlightTerm = "") {
    equipmentList.innerHTML = '';
    const type = searchType.value;
    const isSerie = type === 'serie';
    const mainColIndex = isSerie ? serIndex : eqIndex;
    const subColIndex = isSerie ? eqIndex : serIndex;

    // Update Result Count Badge
    if (resultCount) {
        resultCount.textContent = rows.length;
        resultCount.classList.remove('hidden');
    }

    rows.forEach(row => {
        let mainVal = String(row[mainColIndex] || "");
        let subVal = String(row[subColIndex] || "");
        
        if (!mainVal && !subVal) return;
        if (!mainVal) mainVal = "(Sin datos)";

        const li = document.createElement('li');
        
        let displayHTML = mainVal;
        let subHTMLText = subVal;

        if (highlightTerm) {
            const regex = new RegExp(`(${highlightTerm})`, "gi");
            if (type === 'ambos') {
                displayHTML = displayHTML.replace(regex, "<strong>$1</strong>");
                subHTMLText = subHTMLText.replace(regex, "<strong>$1</strong>");
            } else {
                displayHTML = displayHTML.replace(regex, "<strong>$1</strong>");
            }
        }
        
        let subHTML = '';
        if (subVal && subVal !== 'undefined') {
            subHTML = `<span class="sub-info">${isSerie ? 'Equipo: ' : 'Serie: '} ${subHTMLText}</span>`;
        }

        li.innerHTML = displayHTML + subHTML;
        
        li.addEventListener('click', () => {
            document.querySelectorAll('.equipment-list li').forEach(el => el.classList.remove('active'));
            li.classList.add('active');
            showResults(row);
            closeSidebar(); // Auto-close drawer on mobile when an item is selected
        });

        equipmentList.appendChild(li);
    });
}

function showResults(row) {
    resultsArea.classList.remove('hidden');
    
    // Populate details grid
    detailsGrid.innerHTML = '';
    headers.forEach((header, index) => {
        if (!header) return;
        const value = row[index];
        if (value !== undefined && value !== "") {
            let valueHTML = `<span class="detail-value">${value}</span>`;
            
            if (String(value).trim().startsWith('http')) {
                valueHTML = `<a href="${String(value).trim()}" target="_blank" rel="noopener" style="background: rgba(59, 130, 246, 0.2); color: #60a5fa; padding: 0.2rem 0.6rem; border-radius: 6px; text-decoration: none; font-size: 0.85rem; font-weight: bold; border: 1px solid #3b82f6; display: inline-block;">🔗 Abrir Enlace</a>`;
            }
            
            const div = document.createElement('div');
            div.className = 'detail-item';
            div.innerHTML = `
                <span class="detail-label">${header}</span>
                ${valueHTML}
            `;
            detailsGrid.appendChild(div);
        }
    });

    // Check for inventario
    certInventarioButton.classList.add('hidden');
    certInventarioButton.href = "#";
    inventarioDataSection.classList.add('hidden');
    inventarioDetailsGrid.innerHTML = '';
    
    const mainEquipo = String(row[eqIndex] || "").trim().toLowerCase();
    const mainSerie = String(row[serIndex] || "").trim().toLowerCase();

    let invMatch = undefined;

    if (inventarioData.length > 0) {
        invMatch = inventarioData.find(c => 
            (c.serie !== "" && mainSerie !== "" && c.serie === mainSerie) ||
            (c.equipo !== "" && mainEquipo !== "" && c.equipo === mainEquipo)
        );
        
        if (invMatch) {
            inventarioDataSection.classList.remove('hidden');
            
            inventarioHeaders.forEach((header, index) => {
                if (!header) return;
                const value = invMatch.rowData[index];
                if (value !== undefined && value !== "") {
                    let valueHTML = `<span class="detail-value">${value}</span>`;
                    if (String(value).trim().startsWith('http')) {
                        valueHTML = `<a href="${String(value).trim()}" target="_blank" rel="noopener" style="background: rgba(168, 85, 247, 0.2); color: #c084fc; padding: 0.2rem 0.6rem; border-radius: 6px; text-decoration: none; font-size: 0.85rem; font-weight: bold; border: 1px solid #a855f7; display: inline-block;">🔗 Abrir Enlace</a>`;
                    }
                    const div = document.createElement('div');
                    div.className = 'detail-item';
                    div.innerHTML = `<span class="detail-label">${header}</span>${valueHTML}`;
                    inventarioDetailsGrid.appendChild(div);
                }
            });
            
            if (invMatch.link) {
                certInventarioButton.href = invMatch.link;
                certInventarioButton.classList.remove('hidden');
            }
        }
    }
    
    // First check local storage for a manually added link
    const localKey = `cert_${mainEquipo}_${mainSerie}`;
    let certLink = localStorage.getItem(localKey) || "";
    let searchData = null;
    
    if (certData.length > 0) {
        const match = certData.find(c => 
            c.equipo === mainEquipo && c.serie === mainSerie
        );
        if (match) {
            searchData = match;
            if (!certLink && match.link) {
                certLink = match.link;
            }
        }
    }

    // QR Code generation
    qrCodeContainer.classList.add('hidden');
    qrcodeElement.innerHTML = '';
    
    let finalLinkForQr = "";
    if (certLink) finalLinkForQr = certLink;
    else if (typeof invMatch !== 'undefined' && invMatch && invMatch.link) finalLinkForQr = invMatch.link;
    else {
        for (let i = 0; i < row.length; i++) {
            if (String(row[i]).trim().startsWith('http')) {
                finalLinkForQr = String(row[i]).trim();
                break;
            }
        }
    }

    if (finalLinkForQr && typeof QRCode !== 'undefined') {
        qrCodeContainer.classList.remove('hidden');
        if (qrCodeInstance) {
            qrCodeInstance.clear();
            qrCodeInstance.makeCode(finalLinkForQr);
        } else {
            qrCodeInstance = new QRCode(qrcodeElement, {
                text: finalLinkForQr,
                width: 150,
                height: 150,
                colorDark : "#000000",
                colorLight : "#ffffff",
                correctLevel : QRCode.CorrectLevel.M
            });
        }
    }

    // Load extra data (Fabricante, Modelo, Descripcion)
    const localExtraKey = `extra_${mainEquipo}_${mainSerie}`;
    let savedExtra = null;
    try {
        const savedStr = localStorage.getItem(localExtraKey);
        if (savedStr) savedExtra = JSON.parse(savedStr);
    } catch(e) {}

    let fabricante = savedExtra?.fabricante || (searchData ? searchData.fabricante : "");
    let modelo = savedExtra?.modelo || (searchData ? searchData.modelo : "");
    let descripcion = savedExtra?.descripcion || (searchData ? searchData.descripcion : "");

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
            
            const handlePaste = (e) => {
                setTimeout(() => {
                    const val = e.target.value;
                    let text = val;
                    let parsedFabricante = inputFabricante.value;
                    let parsedModelo = inputModelo.value;
                    let isParsed = false;

                    text = text.replace(/•/g, '').replace(/[\r\n]+/g, ' ');

                    const lookahead = '(?=\\s*(?:modelo|marca|fabricante|(?:breve\\s+)?descripci[oó]n)\\s*[:\\-\\/]|$)';
                    
                    const fabRegex = new RegExp(`(?:fabricante|marca)[^:\\-]*[:\\-]\\s*(.*?)${lookahead}`, 'i');
                    const modRegex = new RegExp(`modelo[^:\\-]*[:\\-]\\s*(.*?)${lookahead}`, 'i');
                    const descRegex = new RegExp(`(?:breve\\s+)?descripci[oó]n[^:\\-]*[:\\-]\\s*(.*?)${lookahead}`, 'i');

                    const fabMatch = text.match(fabRegex);
                    if (fabMatch && fabMatch[1]) {
                        parsedFabricante = fabMatch[1].trim();
                        isParsed = true;
                    }

                    const modMatch = text.match(modRegex);
                    if (modMatch && modMatch[1]) {
                        parsedModelo = modMatch[1].trim();
                        isParsed = true;
                    }

                    let descText = text;
                    const descMatch = text.match(descRegex);
                    if (descMatch && descMatch[1]) {
                        descText = descMatch[1].trim();
                        isParsed = true;
                    } else if (isParsed) {
                        descText = text.replace(fabRegex, '').replace(modRegex, '').trim();
                    }

                    if (isParsed) {
                        if (parsedFabricante) inputFabricante.value = parsedFabricante;
                        if (parsedModelo) inputModelo.value = parsedModelo;
                        
                        let words = descText.split(/\s+/);
                        if (words.length > 100) {
                            descText = words.slice(0, 100).join(' ') + '...';
                        }
                        inputDescripcion.value = descText;

                        if (e.target.id !== 'inputDescripcion') {
                            e.target.value = parsedFabricante || parsedModelo;
                            setTimeout(() => {
                                inputFabricante.value = parsedFabricante;
                                inputModelo.value = parsedModelo;
                            }, 50);
                        }
                    }
                }, 10);
            };
            inputFabricante.onpaste = handlePaste;
            inputModelo.onpaste = handlePaste;
            inputDescripcion.onpaste = handlePaste;

        } else {
            extraDataView.classList.remove('hidden');
            extraDataEdit.classList.add('hidden');
            toggleEditDataBtn.textContent = '✏️ Editar';
        }
    };

    saveExtraDataBtn.onclick = () => {
        fabricante = inputFabricante.value.trim();
        modelo = inputModelo.value.trim();
        descripcion = inputDescripcion.value.trim();
        
        localStorage.setItem(localExtraKey, JSON.stringify({ fabricante, modelo, descripcion }));
        
        viewFabricante.textContent = fabricante || "-";
        viewModelo.textContent = modelo || "-";
        viewDescripcion.textContent = descripcion || "-";
        
        extraDataView.classList.remove('hidden');
        extraDataEdit.classList.add('hidden');
        toggleEditDataBtn.textContent = '✏️ Editar';
    };

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
                localStorage.setItem(localKey, newLink);
                certButton.href = newLink;
                certButton.classList.remove('hidden');
                addCertSection.classList.add('hidden');
            }
        };
    }

    // Handle Image
    const imageName = row[imgIndex];
    
    equipmentImage.onclick = null;
    equipmentImage.style.cursor = 'default';
    equipmentImage.title = "";

    if (imageName && String(imageName).trim() !== "") {
        let imgStr = String(imageName).trim();
        let imgPath = imgStr;
        
        if (!imgStr.toLowerCase().startsWith('http')) {
            imgPath = `Imagenes/${imgStr}`;
            if (!imgPath.toLowerCase().match(/\.(jpg|jpeg|png|gif)$/)) {
                imgPath += '.jpg';
            }
        }

        equipmentImage.src = imgPath;
        equipmentImage.classList.remove('hidden');
        noImageText.classList.add('hidden');
        
        equipmentImage.style.cursor = 'pointer';
        
        equipmentImage.onmouseover = () => {
            equipmentImage.style.transform = 'scale(1.02)';
        };
        equipmentImage.onmouseout = () => {
            equipmentImage.style.transform = 'scale(1)';
        };

        const hasExtraData = (searchData && searchData.tipo) || fabricante || modelo || descripcion;

        if (hasExtraData) {
            equipmentImage.title = "Haz clic para buscar información de este equipo en internet";
        } else {
            equipmentImage.title = "Haz clic derecho y selecciona 'Buscar imagen con Google' para identificarlo";
        }

        equipmentImage.onclick = () => {
            if (hasExtraData) {
                const tipoStr = (searchData && searchData.tipo) ? searchData.tipo : "";
                const query = `${tipoStr} ${fabricante} ${modelo}`.trim();
                window.open(`https://www.google.com/search?q=${encodeURIComponent(query)}`, '_blank');
            } else {
                alert("Como la imagen está en tu computadora local, no podemos enviarla automáticamente.\n\nPara buscarla por imagen, haz CLIC DERECHO sobre esta foto y selecciona 'Buscar imagen con Google' o 'Buscar en la web'.");
            }
        };
    } else {
        equipmentImage.classList.add('hidden');
        noImageText.classList.remove('hidden');
    }
}

// Export to Excel
exportExcelBtn.addEventListener('click', () => {
    if (!inventarioRawData || inventarioRawData.length === 0) {
        alert("No hay datos de Inventario cargados para exportar. Por favor, carga el Inventario primero (Botón 3).");
        return;
    }

    const validIndices = [];
    const cleanHeaders = [];
    if (headers && headers.length > 0) {
        headers.forEach((h, i) => {
            let headerName = h && String(h).trim() !== "" ? String(h).trim() : "";
            
            if (i === imgIndex && headerName === "") {
                headerName = "Imagen";
            }

            if (headerName !== "") {
                validIndices.push(i);
                cleanHeaders.push(headerName);
            } else if (i === imgIndex) {
                validIndices.push(i);
                cleanHeaders.push("Imagen");
            }
        });
    }

    const exportHeaders = [
        ...inventarioHeaders, 
        ...cleanHeaders, 
        "Tipo Extra", "Fabricante Extra", "Modelo Extra", "Enlace Certificado", "Descripción Adicional"
    ];
    const exportData = [exportHeaders];

    let invEqIndex = inventarioHeaders.findIndex(h => h && String(h).toLowerCase().includes('equipo'));
    if (invEqIndex === -1) invEqIndex = 1;
    let invSerIndex = inventarioHeaders.findIndex(h => h && String(h).toLowerCase().includes('serie'));
    if (invSerIndex === -1) invSerIndex = 3;

    let certEqIndex = certHeaders && certHeaders.length > 0 ? certHeaders.findIndex(h => h && String(h).toLowerCase().includes('equipo')) : -1;
    if (certEqIndex === -1) certEqIndex = 1;
    let certSerIndex = certHeaders && certHeaders.length > 0 ? certHeaders.findIndex(h => h && String(h).toLowerCase().includes('serie')) : -1;
    if (certSerIndex === -1) certSerIndex = 3;

    inventarioRawData.forEach(invRow => {
        const invEq = String(invRow[invEqIndex] || "").trim().toLowerCase();
        const invSer = String(invRow[invSerIndex] || "").trim().toLowerCase();

        let mainMatch = null;

        if (currentSheetData && currentSheetData.length > 0) {
            mainMatch = currentSheetData.find(m => {
                const mEq = String(m[eqIndex] || "").trim().toLowerCase();
                const mSer = String(m[serIndex] || "").trim().toLowerCase();
                
                if (invSer !== "" && mSer !== "" && invSer === mSer) return true;
                if (invEq !== "" && mEq !== "" && invEq === mEq) return true;
                return false;
            });
        }

        let certMatch = null;
        if (certRawData && certRawData.length > 0) {
            certMatch = certRawData.find(c => {
                const cEq = String(c[certEqIndex] || "").trim().toLowerCase();
                const cSer = String(c[certSerIndex] || "").trim().toLowerCase();
                
                if (invSer !== "" && cSer !== "" && invSer === cSer) return true;
                if (invEq !== "" && cEq !== "" && invEq === cEq) return true;
                return false;
            });
        }

        let newRow = [];
        inventarioHeaders.forEach((h, i) => {
            newRow.push(invRow[i] !== undefined ? invRow[i] : "");
        });
        
        validIndices.forEach(idx => {
            newRow.push(mainMatch && mainMatch[idx] !== undefined ? mainMatch[idx] : "");
        });

        let tipo = certMatch ? (certMatch[2] || "") : "";
        let fabricante = certMatch ? (certMatch[4] || "") : "";
        let modelo = certMatch ? (certMatch[5] || "") : "";
        let link = certMatch ? (certMatch[6] || "") : "";
        let descripcion = "";

        const localExtraKey = `extra_${invEq}_${invSer}`;
        try {
            const savedStr = localStorage.getItem(localExtraKey);
            if (savedStr) {
                const parsed = JSON.parse(savedStr);
                descripcion = parsed.descripcion || "";
                if (parsed.fabricante) fabricante = parsed.fabricante;
                if (parsed.modelo) modelo = parsed.modelo;
            }
        } catch(e) {}
        
        const localCertKey = `cert_${invEq}_${invSer}`;
        let manualLink = localStorage.getItem(localCertKey);
        if (manualLink) {
            link = manualLink;
        }

        newRow.push(tipo);
        newRow.push(fabricante);
        newRow.push(modelo);
        newRow.push(link);
        newRow.push(descripcion);

        exportData.push(newRow);
    });

    const ws = XLSX.utils.aoa_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Inventario_Exportado");
    
    const dateString = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Inventario_Exportado_${dateString}.xlsx`);
});

// Start app
window.addEventListener('DOMContentLoaded', init);
