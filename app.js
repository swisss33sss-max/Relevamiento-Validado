const EXCEL_FILE = 'IC-SM-CAT-LAB LISTA CATAMARCA LAB  --  2.xlsx';
const CERT_FILE = 'Etiquetas Laboratorio.xls';

let workbook = null;
let sheetsData = {}; 
let currentSheetData = [];
let headers = [];
let certData = [];

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
const certButton = document.getElementById('certButton');
const addCertSection = document.getElementById('addCertSection');
const manualLinkInput = document.getElementById('manualLinkInput');
const saveManualLinkBtn = document.getElementById('saveManualLinkBtn');
const cleanDuplicatesBtn = document.getElementById('cleanDuplicatesBtn');
const extraDataView = document.getElementById('extraDataView');
const extraDataEdit = document.getElementById('extraDataEdit');
const toggleEditDataBtn = document.getElementById('toggleEditDataBtn');
const saveExtraDataBtn = document.getElementById('saveExtraDataBtn');
const viewFabricante = document.getElementById('viewFabricante');
const viewModelo = document.getElementById('viewModelo');
const viewDescripcion = document.getElementById('viewDescripcion');
const inputFabricante = document.getElementById('inputFabricante');
const inputModelo = document.getElementById('inputModelo');
const inputDescripcion = document.getElementById('inputDescripcion');

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
        } catch (e) {
            console.warn("No se pudo cargar 'Etiquetas Laboratorio.xls' automáticamente.", e);
        }

    } catch (error) {
        console.error("Error loading Excel file:", error);
        if (window.location.protocol === 'file:') {
            // Show fallback for local file:// execution
            localFileWarning.classList.remove('hidden');
            sheetSelect.innerHTML = '<option value="">Requiere archivo manual...</option>';
        } else {
            sheetSelect.innerHTML = '<option value="">Error cargando archivo</option>';
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
    localFileWarning.classList.add('hidden');
    loadSheet(sheetSelect.value);
}

function processCertWorkbook(wb) {
    certData = [];
    const sheetName = wb.SheetNames[0];
    const worksheet = wb.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
    
    // Start from row 1 to skip headers
    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (!row) continue;
        const equipo = String(row[1] || "").trim().toLowerCase(); // Col B
        const serie = String(row[3] || "").trim().toLowerCase();  // Col D
        const link = String(row[6] || "").trim();                 // Col G
        const tipo = String(row[2] || "").trim(); // Col C
        const fabricante = String(row[4] || "").trim(); // Col E
        const modelo = String(row[5] || "").trim(); // Col F
        
        if (equipo || serie) {
            certData.push({ equipo, serie, link, tipo, fabricante, modelo });
        }
    }
}

// Event Listeners
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
        document.getElementById('certLoadedMsg').classList.remove('hidden');
        
        const activeLi = document.querySelector('.equipment-list li.active');
        if (activeLi) activeLi.click(); // re-click to refresh details
    };
    reader.readAsArrayBuffer(file);
});

function loadSheet(sheetName) {
    const data = sheetsData[sheetName];
    if (!data || data.length === 0) return;

    // Assuming first row is headers
    headers = data[0];
    currentSheetData = data.slice(1); // skip headers
    
    // Clear search and results
    searchInput.value = '';
    resultsArea.classList.add('hidden');
    renderList(currentSheetData);
}

sheetSelect.addEventListener('change', (e) => {
    loadSheet(e.target.value);
});

searchType.addEventListener('change', () => {
    // Re-render and re-filter based on the new search type
    const val = searchInput.value;
    filterList(val);
});

searchInput.addEventListener('input', function(e) {
    const val = this.value;
    filterList(val);
});

function filterList(searchTerm) {
    if (!searchTerm) {
        renderList(currentSheetData);
        return;
    }
    
    const isSerie = searchType.value === 'serie';
    const searchColIndex = isSerie ? 2 : 0; // Col C or Col A
    
    const matches = currentSheetData.filter(row => {
        const cellVal = row[searchColIndex] ? String(row[searchColIndex]).toLowerCase() : "";
        return cellVal.includes(searchTerm.toLowerCase());
    });
    
    renderList(matches, searchTerm);
}

function renderList(rows, highlightTerm = "") {
    equipmentList.innerHTML = '';
    const isSerie = searchType.value === 'serie';
    const mainColIndex = isSerie ? 2 : 0;
    const subColIndex = isSerie ? 0 : 2;

    rows.forEach(row => {
        const mainVal = String(row[mainColIndex]);
        if (!mainVal || mainVal === 'undefined') return;

        const li = document.createElement('li');
        
        // Highlight matching part if searching
        let displayHTML = mainVal;
        if (highlightTerm) {
            const regex = new RegExp(`(${highlightTerm})`, "gi");
            displayHTML = displayHTML.replace(regex, "<strong>$1</strong>");
        }
        
        const subVal = String(row[subColIndex]);
        let subHTML = '';
        if (subVal && subVal !== 'undefined') {
            subHTML = `<span class="sub-info">${isSerie ? 'Equipo: ' : 'Serie: '} ${subVal}</span>`;
        }

        li.innerHTML = displayHTML + subHTML;
        
        li.addEventListener('click', () => {
            // Remove active class from all
            document.querySelectorAll('.equipment-list li').forEach(el => el.classList.remove('active'));
            li.classList.add('active');
            showResults(row);
        });

        equipmentList.appendChild(li);
    });
}

function showResults(row) {
    resultsArea.classList.remove('hidden');
    
    // Populate details grid
    detailsGrid.innerHTML = '';
    headers.forEach((header, index) => {
        if (!header) return; // skip empty headers
        const value = row[index];
        if (value !== undefined && value !== "") {
            const div = document.createElement('div');
            div.className = 'detail-item';
            div.innerHTML = `
                <span class="detail-label">${header}</span>
                <span class="detail-value">${value}</span>
            `;
            detailsGrid.appendChild(div);
        }
    });

    // Check for certificate and extra data
    const mainEquipo = String(row[0] || "").trim().toLowerCase(); // Col A
    const mainSerie = String(row[2] || "").trim().toLowerCase();  // Col C
    
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
            
            // Inteligencia de Pegado solicitada por el usuario
            const handlePaste = (e) => {
                setTimeout(() => {
                    const val = e.target.value;
                    let text = val;
                    let parsedFabricante = inputFabricante.value;
                    let parsedModelo = inputModelo.value;
                    let isParsed = false;
                    
                    const lowerText = text.toLowerCase();

                    // Buscar "modelo"
                    const modIdx = lowerText.indexOf('modelo');
                    if (modIdx !== -1) {
                        let start = modIdx + 6;
                        while(start < text.length && (text[start] === ':' || text[start] === ' ')) start++;
                        let end = text.substring(start).search(/[\.\,\n\(;]/);
                        if (end === -1) end = 40;
                        let extracted = text.substring(start, start + end).trim();
                        if(extracted) {
                            parsedModelo = extracted;
                            isParsed = true;
                        }
                    }

                    // Buscar "fabricante" o "marca"
                    let fabIdx = lowerText.indexOf('fabricante');
                    if (fabIdx === -1) fabIdx = lowerText.indexOf('marca');
                    if (fabIdx !== -1) {
                        let offset = lowerText.substring(fabIdx).startsWith('fabricante') ? 10 : 5;
                        let start = fabIdx + offset;
                        // Omitir caracteres como : / o espacios inmediatos
                        while(start < text.length && /[:\/\s]/.test(text[start])) start++;
                        let end = text.substring(start).search(/[\.\,\n\(;]/);
                        if (end === -1) end = 50;
                        let extracted = text.substring(start, start + end).trim();
                        if(extracted) {
                            parsedFabricante = extracted;
                            isParsed = true;
                        }
                    }

                    // Buscar "descripcion"
                    let descText = text;
                    let descIdx = lowerText.indexOf('descripción');
                    if (descIdx === -1) descIdx = lowerText.indexOf('descripcion');
                    if (descIdx !== -1) {
                        let start = descIdx + 11;
                        while(start < text.length && /[:\s]/.test(text[start])) start++;
                        descText = text.substring(start).trim();
                        isParsed = true;
                    }

                    // Aplicar los cambios solo si detectó al menos una palabra clave
                    if (isParsed) {
                        if (parsedFabricante) inputFabricante.value = parsedFabricante;
                        if (parsedModelo) inputModelo.value = parsedModelo;
                        
                        // Limitar a 100 palabras exactas
                        let words = descText.split(/\s+/);
                        if (words.length > 100) {
                            descText = words.slice(0, 100).join(' ') + '...';
                        }
                        inputDescripcion.value = descText;

                        // Si el usuario pegó en otro lado, limpiamos el campo para que quede ordenado
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

    // Handle Image (Column G is index 6)
    const imageName = row[6];
    
    // Clear previous image listeners and styles
    equipmentImage.onclick = null;
    equipmentImage.style.cursor = 'default';
    equipmentImage.title = "";

    if (imageName && String(imageName).trim() !== "") {
        let imgPath = `Imagenes/${String(imageName).trim()}`;
        
        if (!imgPath.toLowerCase().match(/\.(jpg|jpeg|png|gif)$/)) {
            imgPath += '.jpg';
        }

        equipmentImage.src = imgPath;
        equipmentImage.classList.remove('hidden');
        noImageText.classList.add('hidden');
        
        // Setup internet search on click
        equipmentImage.style.cursor = 'pointer';
        equipmentImage.style.transition = 'transform 0.2s, box-shadow 0.2s';
        
        equipmentImage.onmouseover = () => {
            equipmentImage.style.transform = 'scale(1.02)';
            equipmentImage.style.boxShadow = '0 8px 25px rgba(0, 0, 0, 0.3)';
        };
        equipmentImage.onmouseout = () => {
            equipmentImage.style.transform = 'scale(1)';
            equipmentImage.style.boxShadow = '0 4px 15px rgba(0, 0, 0, 0.2)';
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

// Start app
window.addEventListener('DOMContentLoaded', init);
