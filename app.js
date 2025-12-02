// ========================================
// Global State
// ========================================
const AppState = {
    currentTab: 'import',
    importedFiles: [],
    mergedData: null,
    reorganizeFile: null,
    columnMapping: {},
    reorganizedData: null
};

// ========================================
// Utility Functions
// ========================================
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

function showAlert(message, type = 'success') {
    const alertContainer = document.getElementById('alertContainer');
    const alert = document.createElement('div');
    alert.className = `alert alert-${type}`;
    alert.innerHTML = `
        <span>${type === 'success' ? '✅' : type === 'error' ? '❌' : '⚠️'}</span>
        <span>${message}</span>
    `;

    alertContainer.appendChild(alert);

    setTimeout(() => {
        alert.style.opacity = '0';
        setTimeout(() => alert.remove(), 300);
    }, 3000);
}

function createTableHTML(data, limit = 100) {
    if (!data || data.length === 0) return '<div class="text-center p-4">Aucune donnée à afficher</div>';

    const headers = data[0];
    const rows = data.slice(1, limit + 1);

    let html = '<table><thead><tr>';
    headers.forEach(header => {
        html += `<th>${header || ''}</th>`;
    });
    html += '</tr></thead><tbody>';

    rows.forEach(row => {
        html += '<tr>';
        headers.forEach((_, index) => {
            html += `<td>${row[index] || ''}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';

    if (data.length > limit) {
        html += `<div class="text-center p-2 text-muted">... et ${data.length - limit - 1} autres lignes</div>`;
    }

    return html;
}

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {
                    type: 'array',
                    codepage: 65001,
                    cellText: false,
                    cellDates: true,
                    raw: true
                });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                    header: 1,
                    raw: true,
                    defval: ''
                });

                // Convert all cells to strings to preserve SIRET and other large numbers
                const processedData = jsonData.map(row =>
                    row.map(cell => {
                        if (cell === null || cell === undefined) return '';
                        // Convert numbers to strings to preserve precision for large numbers like SIRET
                        if (typeof cell === 'number') {
                            return cell.toString();
                        }
                        return String(cell);
                    })
                );

                resolve(processedData);
            } catch (error) {
                reject(error);
            }
        };

        reader.onerror = () => reject(new Error('Erreur de lecture du fichier'));
        reader.readAsArrayBuffer(file);
    });
}

function similarity(s1, s2) {
    const longer = s1.length > s2.length ? s1 : s2;
    const shorter = s1.length > s2.length ? s2 : s1;

    if (longer.length === 0) return 1.0;

    const editDistance = (s1, s2) => {
        s1 = s1.toLowerCase();
        s2 = s2.toLowerCase();
        const costs = [];

        for (let i = 0; i <= s1.length; i++) {
            let lastValue = i;
            for (let j = 0; j <= s2.length; j++) {
                if (i === 0) {
                    costs[j] = j;
                } else if (j > 0) {
                    let newValue = costs[j - 1];
                    if (s1.charAt(i - 1) !== s2.charAt(j - 1)) {
                        newValue = Math.min(Math.min(newValue, lastValue), costs[j]) + 1;
                    }
                    costs[j - 1] = lastValue;
                    lastValue = newValue;
                }
            }
            if (i > 0) costs[s2.length] = lastValue;
        }
        return costs[s2.length];
    };

    return (longer.length - editDistance(longer, shorter)) / longer.length;
}

function formatPhoneNumber(value) {
    if (!value) return value;

    const strValue = String(value).trim().replace(/[\s\-\.\(\)]/g, '');

    if (/^0\d{9,}$/.test(strValue)) {
        return strValue.substring(1);
    }

    return value;
}

// ========================================
// Tab Navigation
// ========================================
function initTabs() {
    const tabButtons = document.querySelectorAll('.tab-btn');
    const tabContents = document.querySelectorAll('.tab-content');

    tabButtons.forEach(button => {
        button.addEventListener('click', () => {
            const tabId = button.dataset.tab;

            tabButtons.forEach(btn => btn.classList.remove('active'));
            tabContents.forEach(content => content.classList.remove('active'));

            button.classList.add('active');
            document.getElementById(tabId).classList.add('active');

            AppState.currentTab = tabId;
        });
    });
}

// ========================================
// Import & Merge Functionality
// ========================================
function initImportMerge() {
    const uploadZone = document.getElementById('uploadZone');
    const fileInput = document.getElementById('fileInput');
    const fileList = document.getElementById('fileList');
    const fileListContainer = document.getElementById('fileListContainer');
    const mergeCard = document.getElementById('mergeCard');
    const mergeBtn = document.getElementById('mergeBtn');

    uploadZone.addEventListener('click', () => fileInput.click());

    uploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadZone.classList.add('dragover');
    });

    uploadZone.addEventListener('dragleave', () => {
        uploadZone.classList.remove('dragover');
    });

    uploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadZone.classList.remove('dragover');
        handleFiles(e.dataTransfer.files);
    });

    fileInput.addEventListener('change', (e) => {
        handleFiles(e.target.files);
    });

    async function handleFiles(files) {
        const validExtensions = ['.xlsx', '.xls', '.csv'];

        for (let file of files) {
            const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

            if (!validExtensions.includes(ext)) {
                showAlert(`Format non supporté: ${file.name}`, 'error');
                continue;
            }

            try {
                const data = await readExcelFile(file);
                AppState.importedFiles.push({
                    name: file.name,
                    size: file.size,
                    data: data
                });

                showAlert(`Fichier importé: ${file.name}`, 'success');
            } catch (error) {
                showAlert(`Erreur lors de l'import de ${file.name}: ${error.message}`, 'error');
            }
        }

        updateFileList();
    }

    function updateFileList() {
        if (AppState.importedFiles.length === 0) {
            fileListContainer.classList.add('hidden');
            mergeCard.classList.add('hidden');
            return;
        }

        fileListContainer.classList.remove('hidden');
        mergeCard.classList.remove('hidden');

        fileList.innerHTML = '';
        AppState.importedFiles.forEach((file, index) => {
            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <div class="file-info">
                    <span class="file-icon">📄</span>
                    <div>
                        <div class="file-name">${file.name}</div>
                        <div class="file-size">${formatFileSize(file.size)} • ${file.data.length - 1} lignes</div>
                    </div>
                </div>
                <button class="btn-icon" onclick="removeFile(${index})">
                    <span style="font-size: 1.25rem;">✕</span>
                </button>
            `;
            fileList.appendChild(fileItem);
        });
    }

    mergeBtn.addEventListener('click', () => {
        if (AppState.importedFiles.length === 0) {
            showAlert('Aucun fichier à fusionner', 'warning');
            return;
        }

        try {
            const allHeaders = new Set();
            AppState.importedFiles.forEach(file => {
                if (file.data.length > 0) {
                    file.data[0].forEach(header => {
                        if (header) allHeaders.add(header);
                    });
                }
            });

            const headers = Array.from(allHeaders);
            const mergedRows = [headers];

            AppState.importedFiles.forEach(file => {
                if (file.data.length <= 1) return;

                const fileHeaders = file.data[0];
                const fileRows = file.data.slice(1);

                fileRows.forEach(row => {
                    const newRow = new Array(headers.length).fill('');

                    fileHeaders.forEach((header, index) => {
                        const targetIndex = headers.indexOf(header);
                        if (targetIndex !== -1 && row[index] !== undefined) {
                            newRow[targetIndex] = row[index];
                        }
                    });

                    mergedRows.push(newRow);
                });
            });

            AppState.mergedData = mergedRows;

            const mergePreview = document.getElementById('mergePreview');
            const mergeTableContainer = document.getElementById('mergeTableContainer');
            const mergeStats = document.getElementById('mergeStats');

            mergeTableContainer.innerHTML = createTableHTML(mergedRows);
            mergeStats.textContent = `${mergedRows.length - 1} lignes • ${headers.length} colonnes`;
            mergePreview.classList.remove('hidden');

            const actionContainer = document.createElement('div');
            actionContainer.className = 'text-center mt-lg';
            actionContainer.innerHTML = `
                <button id="proceedToReorgBtn" class="btn btn-primary">
                    <span>➡️</span> Réorganiser ce résultat
                </button>
            `;

            const existingBtn = document.getElementById('proceedToReorgBtn');
            if (existingBtn) existingBtn.parentElement.remove();

            mergePreview.appendChild(actionContainer);

            document.getElementById('proceedToReorgBtn').addEventListener('click', () => {
                const reorgTabBtn = document.querySelector('.tab-btn[data-tab="reorganize"]');
                if (reorgTabBtn) {
                    reorgTabBtn.click();
                } else {
                    document.querySelectorAll('.tab-btn')[1].click();
                }

                setTimeout(() => {
                    if (window.loadReorganizeData) {
                        window.loadReorganizeData(AppState.mergedData, 'Fusion_Resultat.xlsx');
                    } else {
                        console.error('loadReorganizeData function not found');
                        showAlert('Erreur interne: fonction de chargement introuvable', 'error');
                    }
                }, 10);
            });

            updateExportTab();

            showAlert(`${AppState.importedFiles.length} fichiers fusionnés avec succès!`, 'success');
        } catch (error) {
            showAlert(`Erreur lors de la fusion: ${error.message}`, 'error');
        }
    });
}

function removeFile(index) {
    AppState.importedFiles.splice(index, 1);
    document.getElementById('fileList').children[index].remove();

    if (AppState.importedFiles.length === 0) {
        document.getElementById('fileListContainer').classList.add('hidden');
        document.getElementById('mergeCard').classList.add('hidden');
    }
}

// ========================================
// Reorganize Functionality
// ========================================
window.loadReorganizeData = function (data, filename) {
    AppState.reorganizeFile = {
        name: filename,
        data: data
    };

    const mappingCard = document.getElementById('mappingCard');
    const importCard = document.getElementById('reorganizeImportCard');

    mappingCard.classList.remove('hidden');
    if (importCard) importCard.classList.add('hidden');

    const event = new CustomEvent('initMapping', { detail: data });
    document.dispatchEvent(event);

    showAlert(`Données chargées pour réorganisation`, 'success');
};

function initReorganize() {
    const reorganizeUploadZone = document.getElementById('reorganizeUploadZone');
    const reorganizeFileInput = document.getElementById('reorganizeFileInput');
    const autoMapBtn = document.getElementById('autoMapBtn');
    const applyMappingBtn = document.getElementById('applyMappingBtn');

    document.addEventListener('initMapping', (e) => {
        initColumnMapping(e.detail);
    });

    reorganizeUploadZone.addEventListener('click', () => reorganizeFileInput.click());

    reorganizeUploadZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        reorganizeUploadZone.classList.add('dragover');
    });

    reorganizeUploadZone.addEventListener('dragleave', () => {
        reorganizeUploadZone.classList.remove('dragover');
    });

    reorganizeUploadZone.addEventListener('drop', (e) => {
        e.preventDefault();
        reorganizeUploadZone.classList.remove('dragover');
        handleReorganizeFile(e.dataTransfer.files[0]);
    });

    reorganizeFileInput.addEventListener('change', (e) => {
        handleReorganizeFile(e.target.files[0]);
    });

    async function handleReorganizeFile(file) {
        if (!file) return;

        const validExtensions = ['.xlsx', '.xls', '.csv'];
        const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

        if (!validExtensions.includes(ext)) {
            showAlert('Format non supporté', 'error');
            return;
        }

        try {
            const data = await readExcelFile(file);
            window.loadReorganizeData(data, file.name);
            showAlert(`Fichier importé: ${file.name}`, 'success');
        } catch (error) {
            showAlert(`Erreur lors de l'import: ${error.message}`, 'error');
        }
    }

    function initColumnMapping(data) {
        if (data.length === 0) return;

        const currentColumns = data[0];
        const currentColumnsDiv = document.getElementById('currentColumns');
        const targetColumnsDiv = document.getElementById('targetColumns');

        AppState.columnMapping = {};
        currentColumnsDiv.innerHTML = '';
        targetColumnsDiv.innerHTML = '';

        currentColumns.forEach((col, index) => {
            const colItem = document.createElement('div');
            colItem.className = 'column-item';
            colItem.textContent = col || `Colonne ${index + 1}`;
            colItem.dataset.index = index;
            colItem.draggable = true;

            colItem.addEventListener('dragstart', (e) => {
                e.dataTransfer.setData('columnIndex', index);
                colItem.classList.add('dragging');
            });

            colItem.addEventListener('dragend', () => {
                colItem.classList.remove('dragging');
            });

            currentColumnsDiv.appendChild(colItem);
        });

        const targetColumns = currentColumns.map((col, i) => col || `Titre ${i + 1}`);

        targetColumns.forEach((col, index) => {
            const colItem = document.createElement('div');
            colItem.className = 'column-item';
            colItem.textContent = col;
            colItem.dataset.targetIndex = index;

            colItem.addEventListener('dragover', (e) => {
                e.preventDefault();
            });

            colItem.addEventListener('drop', (e) => {
                e.preventDefault();
                const sourceIndex = parseInt(e.dataTransfer.getData('columnIndex'));
                AppState.columnMapping[sourceIndex] = index;
                updateMappingDisplay();
                showAlert(`Colonne mappée: ${currentColumns[sourceIndex]} → ${col}`, 'success');
            });

            targetColumnsDiv.appendChild(colItem);
        });
    }

    function updateMappingDisplay() {
        const currentItems = document.querySelectorAll('#currentColumns .column-item');
        const targetItems = document.querySelectorAll('#targetColumns .column-item');

        currentItems.forEach(item => item.classList.remove('mapped'));
        targetItems.forEach(item => item.classList.remove('mapped'));

        Object.entries(AppState.columnMapping).forEach(([source, target]) => {
            currentItems[source]?.classList.add('mapped');
            targetItems[target]?.classList.add('mapped');
        });
    }

    autoMapBtn.addEventListener('click', () => {
        if (!AppState.reorganizeFile) return;

        const currentColumns = AppState.reorganizeFile.data[0];
        const targetColumns = currentColumns;

        AppState.columnMapping = {};

        currentColumns.forEach((currentCol, sourceIndex) => {
            let bestMatch = -1;
            let bestScore = 0;

            targetColumns.forEach((targetCol, targetIndex) => {
                const score = similarity(currentCol || '', targetCol || '');
                if (score > bestScore && score > 0.5) {
                    bestScore = score;
                    bestMatch = targetIndex;
                }
            });

            if (bestMatch !== -1) {
                AppState.columnMapping[sourceIndex] = bestMatch;
            }
        });

        updateMappingDisplay();
        showAlert('Détection automatique effectuée', 'success');
    });

    applyMappingBtn.addEventListener('click', () => {
        if (!AppState.reorganizeFile) return;

        const data = AppState.reorganizeFile.data;
        const headers = data[0];
        const rows = data.slice(1);

        const removeLeadingZero = document.getElementById('removeLeadingZero').checked;

        const newHeaders = new Array(headers.length);
        const newRows = [];

        Object.entries(AppState.columnMapping).forEach(([source, target]) => {
            newHeaders[target] = headers[source];
        });

        newHeaders.forEach((header, index) => {
            if (!header) newHeaders[index] = headers[index] || `Colonne ${index + 1}`;
        });

        rows.forEach(row => {
            const newRow = new Array(headers.length).fill('');
            Object.entries(AppState.columnMapping).forEach(([source, target]) => {
                let value = row[source];

                if (removeLeadingZero && value) {
                    value = formatPhoneNumber(value);
                }

                newRow[target] = value;
            });
            row.forEach((cell, index) => {
                if (AppState.columnMapping[index] === undefined) {
                    let value = cell;

                    if (removeLeadingZero && value) {
                        value = formatPhoneNumber(value);
                    }

                    newRow[index] = value;
                }
            });
            newRows.push(newRow);
        });

        AppState.reorganizedData = [newHeaders, ...newRows];

        const reorganizePreview = document.getElementById('reorganizePreview');
        const reorganizeTableContainer = document.getElementById('reorganizeTableContainer');

        reorganizeTableContainer.innerHTML = createTableHTML(AppState.reorganizedData);
        reorganizePreview.classList.remove('hidden');

        updateExportTab();

        const message = removeLeadingZero
            ? 'Réorganisation appliquée avec formatage des numéros de téléphone !'
            : 'Réorganisation appliquée avec succès!';
        showAlert(message, 'success');
    });
}

// ========================================
// Export Functionality
// ========================================
function initExport() {
    const exportBtn = document.getElementById('exportBtn');
    const formatOptions = document.querySelectorAll('.format-option');

    formatOptions.forEach(option => {
        option.addEventListener('click', () => {
            formatOptions.forEach(opt => opt.classList.remove('selected'));
            option.classList.add('selected');
            option.querySelector('input[type="radio"]').checked = true;
        });
    });

    exportBtn.addEventListener('click', () => {
        const data = AppState.reorganizedData || AppState.mergedData;

        if (!data) {
            showAlert('Aucune donnée à exporter', 'warning');
            return;
        }

        const format = document.querySelector('input[name="exportFormat"]:checked').value;
        const fileName = document.getElementById('exportFileName').value || 'fichier_exporte';

        exportFile(data, fileName, format);
    });
}

function updateExportTab() {
    const data = AppState.reorganizedData || AppState.mergedData;

    if (!data) return;

    const exportMessage = document.getElementById('exportMessage');
    const exportOptions = document.getElementById('exportOptions');
    const exportTableContainer = document.getElementById('exportTableContainer');
    const exportStats = document.getElementById('exportStats');

    exportMessage.classList.add('hidden');
    exportOptions.classList.remove('hidden');

    exportTableContainer.innerHTML = createTableHTML(data);
    exportStats.textContent = `${data.length - 1} lignes • ${data[0].length} colonnes`;
}

function exportFile(data, fileName, format) {
    try {
        // Prepare data with apostrophe prefix for large numbers (SIRET, etc.)
        const formattedData = data.map(row =>
            row.map(cell => {
                const value = String(cell);
                // Check if it's a large number (10+ digits) - prefix with apostrophe for Excel
                if (/^\d{10,}$/.test(value)) {
                    return `'${value}`; // Apostrophe forces Excel to treat as text
                }
                return cell;
            })
        );

        const ws = XLSX.utils.aoa_to_sheet(formattedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

        let bookType = format;
        if (format === 'xlsx') bookType = 'xlsx';
        else if (format === 'xls') bookType = 'xls';
        else if (format === 'csv') bookType = 'csv';
        else if (format === 'ods') bookType = 'ods';

        const writeOptions = {
            bookType: bookType,
            type: 'array',
            codepage: 65001
        };

        if (format === 'csv') {
            writeOptions.FS = ';';
            writeOptions.bookSST = false;
        }

        // Generate file as array buffer
        const wbout = XLSX.write(wb, writeOptions);

        // Create blob and download with proper filename
        const blob = new Blob([wbout], {
            type: format === 'csv'
                ? 'text/csv;charset=utf-8;'
                : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        });

        const link = document.createElement('a');
        const url = URL.createObjectURL(blob);
        link.href = url;
        link.download = `${fileName}.${format}`;

        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);

        // Clean up the URL object
        setTimeout(() => URL.revokeObjectURL(url), 100);

        showAlert(`Fichier exporté: ${fileName}.${format}`, 'success');
    } catch (error) {
        showAlert(`Erreur lors de l'export: ${error.message}`, 'error');
    }
}

// ========================================
// Initialize Application
// ========================================
document.addEventListener('DOMContentLoaded', () => {
    initTabs();
    initImportMerge();
    initReorganize();
    initExport();

    console.log('📊 Excel Manager initialized');
});
