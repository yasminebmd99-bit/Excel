// ========================================
// Global State
// ========================================
const AppState = {
    currentTab: 'import',
    currentStep: 1,
    importedFiles: [],
    mergedData: null,
    reorganizeFile: null,
    columnMapping: {},
    columnOrder: [],
    columnVisibility: {},
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

function showLoader() {
    document.getElementById('loader').classList.remove('hidden');
}

function hideLoader() {
    document.getElementById('loader').classList.add('hidden');
}

function createTableHTML(data, limit = 100) {
    if (!data || data.length === 0) return '<div class="no-data-message"><p>Aucune donnée à afficher</p></div>';

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
            html += `<td>${row[index] !== undefined ? row[index] : ''}</td>`;
        });
        html += '</tr>';
    });

    html += '</tbody></table>';

    if (data.length > limit + 1) {
        html += `<div style="text-align: center; padding: 1rem; color: var(--text-tertiary);">... et ${data.length - limit - 1} autres lignes</div>`;
    }

    return html;
}

function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();

        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const isCSV = file.name.toLowerCase().endsWith('.csv');

                let workbook;

                if (isCSV) {
                    // Pour les fichiers CSV, détecter l'encodage
                    let textContent;

                    // Essayer d'abord UTF-8
                    const decoder = new TextDecoder('utf-8', { fatal: true });
                    try {
                        textContent = decoder.decode(data);
                    } catch (utf8Error) {
                        // Si UTF-8 échoue, utiliser Windows-1252 (Latin-1)
                        const latin1Decoder = new TextDecoder('windows-1252');
                        textContent = latin1Decoder.decode(data);
                    }

                    // Vérifier si le texte contient des caractères de remplacement (signe d'encodage incorrect)
                    if (textContent.includes('\uFFFD')) {
                        // Réessayer avec Windows-1252
                        const latin1Decoder = new TextDecoder('windows-1252');
                        textContent = latin1Decoder.decode(data);
                    }

                    // Supprimer le BOM si présent
                    if (textContent.charCodeAt(0) === 0xFEFF) {
                        textContent = textContent.substring(1);
                    }

                    workbook = XLSX.read(textContent, {
                        type: 'string',
                        cellText: false,
                        cellDates: true,
                        cellNF: false,
                        raw: true
                    });
                } else {
                    // Pour les fichiers Excel, utiliser la lecture binaire standard
                    // cellNF: false pour ignorer les formats de nombres personnalisés (qui peuvent ajouter des 0)
                    workbook = XLSX.read(data, {
                        type: 'array',
                        cellText: false,
                        cellDates: true,
                        cellNF: false,
                        raw: true
                    });
                }

                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];

                // Lire les valeurs brutes sans formatage
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, {
                    header: 1,
                    raw: true,
                    rawNumbers: true,
                    defval: ''
                });

                // Convert all cells to strings to preserve SIRET and other large numbers
                const processedData = jsonData.map(row =>
                    row.map(cell => {
                        if (cell === null || cell === undefined) return '';
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

function formatPhoneNumber(value, removeLeadingZero, removePlus33) {
    if (!value) return value;

    // Si la valeur contient des lettres, on ne touche pas (pour éviter de casser les adresses, emails, etc.)
    if (/[a-zA-Z]/.test(String(value))) return value;

    let strValue = String(value).trim().replace(/[\s\-\.\(\)]/g, '');

    // Remove +33 prefix if option is enabled
    if (removePlus33) {
        if (strValue.startsWith('+33')) {
            strValue = strValue.substring(3);
        } else if (strValue.startsWith('0033')) {
            strValue = strValue.substring(4);
        } else if (strValue.startsWith('33') && strValue.length >= 11) {
            strValue = strValue.substring(2);
        }
    }

    // Remove leading 0 if option is enabled
    if (removeLeadingZero && /^0\d{9,}$/.test(strValue)) {
        strValue = strValue.substring(1);
    }

    return strValue;
}

// ========================================
// Navigation & Progress
// ========================================
function updateProgress(step) {
    AppState.currentStep = step;

    document.querySelectorAll('.progress-steps .step').forEach((el, index) => {
        el.classList.remove('active', 'completed');
        if (index + 1 < step) {
            el.classList.add('completed');
        } else if (index + 1 === step) {
            el.classList.add('active');
        }
    });
}

function switchTab(tabId) {
    AppState.currentTab = tabId;

    // Update tab cards
    document.querySelectorAll('.tab-card').forEach(card => {
        card.classList.remove('active');
        if (card.dataset.tab === tabId) {
            card.classList.add('active');
        }
    });

    // Update tab content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    document.getElementById(tabId).classList.add('active');

    // Update progress
    if (tabId === 'import') updateProgress(1);
    else if (tabId === 'reorganize') updateProgress(2);
    else if (tabId === 'export') updateProgress(3);
}

function initNavigation() {
    // Tab cards click
    document.querySelectorAll('.tab-card').forEach(card => {
        card.addEventListener('click', () => {
            switchTab(card.dataset.tab);
        });
    });

    // Progress steps click
    document.querySelectorAll('.progress-steps .step').forEach(step => {
        step.addEventListener('click', () => {
            const stepNum = parseInt(step.dataset.step);
            if (stepNum === 1) switchTab('import');
            else if (stepNum === 2) switchTab('reorganize');
            else if (stepNum === 3) switchTab('export');
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
    const downloadExampleBtn = document.getElementById('downloadExampleBtn');
    const clearAllFilesBtn = document.getElementById('clearAllFilesBtn');

    if (clearAllFilesBtn) {
        clearAllFilesBtn.addEventListener('click', () => {
            if (AppState.importedFiles.length === 0) return;

            if (confirm('Voulez-vous vraiment supprimer tous les fichiers importés ? Cette action effacera également les données fusionnées.')) {
                AppState.importedFiles = [];
                AppState.mergedData = null;
                updateFileList();

                // Hide merge preview
                document.getElementById('mergePreview').classList.add('hidden');
                document.getElementById('mergeTableContainer').innerHTML = '';
                document.getElementById('mergeStats').innerHTML = '';

                // Clear file input to allow re-selecting the same file
                if (fileInput) fileInput.value = '';

                showAlert('Tous les fichiers ont été supprimés', 'success');
            }
        });
    }

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

    // Download example file
    downloadExampleBtn.addEventListener('click', () => {
        const exampleData = [
            ['Nom', 'Prénom', 'Email', 'Téléphone', 'Ville'],
            ['Dupont', 'Jean', 'jean.dupont@email.com', '0612345678', 'Paris'],
            ['Martin', 'Marie', 'marie.martin@email.com', '0698765432', 'Lyon'],
            ['Bernard', 'Pierre', 'pierre.bernard@email.com', '+33612345678', 'Marseille']
        ];

        const ws = XLSX.utils.aoa_to_sheet(exampleData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Exemple');
        XLSX.writeFile(wb, 'exemple_excel_manager.xlsx');

        showAlert('Fichier d\'exemple téléchargé !', 'success');
    });

    async function handleFiles(files) {
        const validExtensions = ['.xlsx', '.xls', '.csv'];
        showLoader();

        for (let file of files) {
            const ext = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();

            if (!validExtensions.includes(ext)) {
                showAlert(`Format non supporté: ${file.name}`, 'error');
                continue;
            }

            if (file.size > 200 * 1024 * 1024) {
                showAlert(`Fichier très volumineux: ${file.name} (max recommandé 200 Mo)`, 'warning');
            }

            try {
                const data = await readExcelFile(file);
                AppState.importedFiles.push({
                    name: file.name,
                    size: file.size,
                    data: data,
                    rows: data.length - 1,
                    cols: data[0] ? data[0].length : 0
                });

                showAlert(`Fichier importé: ${file.name}`, 'success');
            } catch (error) {
                showAlert(`Erreur lors de l'import de ${file.name}: ${error.message}`, 'error');
            }
        }

        hideLoader();
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

        // Update file count
        document.getElementById('fileCount').textContent = `${AppState.importedFiles.length} fichier(s)`;

        fileList.innerHTML = '';
        let totalRows = 0;
        let totalCols = 0;

        AppState.importedFiles.forEach((file, index) => {
            totalRows += file.rows;
            totalCols = Math.max(totalCols, file.cols);

            const fileItem = document.createElement('div');
            fileItem.className = 'file-item';
            fileItem.innerHTML = `
                <div class="file-info">
                    <span class="file-icon">📄</span>
                    <div>
                        <div class="file-name">${file.name}</div>
                        <div class="file-size">${formatFileSize(file.size)} • ${file.rows} lignes • ${file.cols} colonnes</div>
                    </div>
                </div>
                <button class="btn-icon" onclick="removeFile(${index})" title="Supprimer">
                    <span style="font-size: 1.25rem;">✕</span>
                </button>
            `;
            fileList.appendChild(fileItem);
        });

        // Show import stats
        const importStats = document.getElementById('importStats');
        importStats.classList.remove('hidden');
        importStats.innerHTML = `
            <span>📊 Total: ${AppState.importedFiles.length} fichiers</span>
            <span>📝 ${totalRows} lignes</span>
            <span>📑 ${totalCols} colonnes max</span>
        `;
    }

    mergeBtn.addEventListener('click', () => {
        if (AppState.importedFiles.length === 0) {
            showAlert('Aucun fichier à fusionner', 'warning');
            return;
        }

        showLoader();

        setTimeout(() => {
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

                updateExportTab();
                hideLoader();

                showAlert(`${AppState.importedFiles.length} fichiers fusionnés avec succès!`, 'success');
            } catch (error) {
                hideLoader();
                showAlert(`Erreur lors de la fusion: ${error.message}`, 'error');
            }
        }, 100);
    });

    // Proceed to reorganization button
    document.getElementById('proceedToReorgBtn').addEventListener('click', () => {
        switchTab('reorganize');
        setTimeout(() => {
            window.loadReorganizeData(AppState.mergedData, 'Fusion_Resultat.xlsx');
        }, 50);
    });
}

function removeFile(index) {
    AppState.importedFiles.splice(index, 1);

    const fileList = document.getElementById('fileList');
    const fileListContainer = document.getElementById('fileListContainer');
    const mergeCard = document.getElementById('mergeCard');

    if (AppState.importedFiles.length === 0) {
        fileListContainer.classList.add('hidden');
        mergeCard.classList.add('hidden');
    } else {
        // Re-render file list
        document.getElementById('fileCount').textContent = `${AppState.importedFiles.length} fichier(s)`;
        fileList.children[index].remove();
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

    initColumnControls(data);

    showAlert(`Données chargées pour réorganisation`, 'success');
};

function initColumnControls(data) {
    if (!data || data.length === 0) return;

    const headers = data[0];

    // Initialize column order and visibility
    AppState.columnOrder = headers.map((_, i) => i);
    AppState.columnVisibility = {};
    headers.forEach((_, i) => {
        AppState.columnVisibility[i] = true;
    });

    // Render column checkboxes
    renderColumnCheckboxes(headers);

    // Render sortable columns
    renderSortableColumns(headers);

    // Update preview
    updateColumnPreview();
}

function renderColumnCheckboxes(headers) {
    const container = document.getElementById('columnCheckboxes');
    container.innerHTML = '';

    headers.forEach((header, index) => {
        const label = document.createElement('label');
        label.className = 'column-checkbox';
        label.innerHTML = `
            <input type="checkbox" checked data-col-index="${index}">
            <span>${header || `Colonne ${index + 1}`}</span>
        `;

        const checkbox = label.querySelector('input');
        checkbox.addEventListener('change', (e) => {
            AppState.columnVisibility[index] = e.target.checked;
            label.classList.toggle('hidden-col', !e.target.checked);
            updateColumnPreview();
        });

        container.appendChild(label);
    });
}

function renderSortableColumns(headers) {
    const container = document.getElementById('sortableColumns');
    container.innerHTML = '';

    let draggedIndex = null;

    // Render columns in the current order
    AppState.columnOrder.forEach((originalIndex, displayIndex) => {
        const header = headers[originalIndex];
        const col = document.createElement('div');
        col.className = 'sortable-column';
        col.draggable = true;
        col.dataset.displayIndex = displayIndex;
        col.dataset.originalIndex = originalIndex;
        col.innerHTML = `
            <span class="grip">⠿</span>
            <span class="col-name">${header || `Colonne ${originalIndex + 1}`}</span>
            <span class="col-index">#${displayIndex + 1}</span>
        `;

        // Drag start
        col.addEventListener('dragstart', (e) => {
            draggedIndex = displayIndex;
            col.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
            e.dataTransfer.setData('text/plain', displayIndex.toString());
        });

        // Drag end
        col.addEventListener('dragend', () => {
            col.classList.remove('dragging');
            draggedIndex = null;
            // Remove all drag-over classes
            container.querySelectorAll('.sortable-column').forEach(c => c.classList.remove('drag-over'));
        });

        // Drag over
        col.addEventListener('dragover', (e) => {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            col.classList.add('drag-over');
        });

        // Drag leave
        col.addEventListener('dragleave', () => {
            col.classList.remove('drag-over');
        });

        // Drop
        col.addEventListener('drop', (e) => {
            e.preventDefault();
            col.classList.remove('drag-over');

            const fromIndex = parseInt(e.dataTransfer.getData('text/plain'));
            const toIndex = displayIndex;

            if (fromIndex !== toIndex) {
                // Reorder the columnOrder array
                const newOrder = [...AppState.columnOrder];
                const [movedItem] = newOrder.splice(fromIndex, 1);
                newOrder.splice(toIndex, 0, movedItem);
                AppState.columnOrder = newOrder;

                // Re-render with updated order
                renderSortableColumns(headers);
                updateColumnPreview();

                showAlert(`Colonne déplacée à la position ${toIndex + 1}`, 'success');
            }
        });

        container.appendChild(col);
    });
}

function updateColumnPreview() {
    if (!AppState.reorganizeFile) return;

    const data = AppState.reorganizeFile.data;
    const headers = data[0];
    const rows = data.slice(1, 6); // Show first 5 rows

    // Build preview with current order and visibility
    const visibleColumns = AppState.columnOrder.filter(i => AppState.columnVisibility[i]);

    const previewData = [
        visibleColumns.map(i => headers[i]),
        ...rows.map(row => visibleColumns.map(i => row[i] || ''))
    ];

    const container = document.getElementById('previewTableContainer');
    container.innerHTML = createTableHTML(previewData, 10);

    document.getElementById('columnPreview').classList.remove('hidden');
}

function initReorganize() {
    const reorganizeUploadZone = document.getElementById('reorganizeUploadZone');
    const reorganizeFileInput = document.getElementById('reorganizeFileInput');
    const autoMapBtn = document.getElementById('autoMapBtn');
    const applyMappingBtn = document.getElementById('applyMappingBtn');
    const selectAllBtn = document.getElementById('selectAllColumns');
    const deselectAllBtn = document.getElementById('deselectAllColumns');
    const resetReorganizeBtn = document.getElementById('resetReorganizeBtn');

    if (resetReorganizeBtn) {
        resetReorganizeBtn.addEventListener('click', () => {
            if (confirm('Voulez-vous vraiment réinitialiser la réorganisation actuelle ?')) {
                AppState.reorganizeFile = null;
                AppState.reorganizedData = null;
                AppState.columnOrder = [];
                AppState.columnVisibility = {};

                // Interface reset
                document.getElementById('mappingCard').classList.add('hidden');
                document.getElementById('reorganizeImportCard').classList.remove('hidden');

                // Clear inputs
                if (reorganizeFileInput) reorganizeFileInput.value = '';
                document.getElementById('reorganizePreview').classList.add('hidden');
                document.getElementById('reorganizeTableContainer').innerHTML = '';

                showAlert('Réorganisation réinitialisée', 'success');
            }
        });
    }

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

        showLoader();

        try {
            const data = await readExcelFile(file);
            window.loadReorganizeData(data, file.name);
            hideLoader();
        } catch (error) {
            hideLoader();
            showAlert(`Erreur lors de l'import: ${error.message}`, 'error');
        }
    }

    // Select all columns
    selectAllBtn.addEventListener('click', () => {
        document.querySelectorAll('#columnCheckboxes input[type="checkbox"]').forEach(cb => {
            cb.checked = true;
            cb.dispatchEvent(new Event('change'));
        });
    });

    // Deselect all columns
    deselectAllBtn.addEventListener('click', () => {
        document.querySelectorAll('#columnCheckboxes input[type="checkbox"]').forEach(cb => {
            cb.checked = false;
            cb.dispatchEvent(new Event('change'));
        });
    });

    // Auto map button
    autoMapBtn.addEventListener('click', () => {
        showAlert('Détection automatique effectuée', 'success');
        updateColumnPreview();
    });

    // Apply mapping button
    applyMappingBtn.addEventListener('click', () => {
        if (!AppState.reorganizeFile) return;

        showLoader();

        setTimeout(() => {
            const data = AppState.reorganizeFile.data;
            const headers = data[0];
            const rows = data.slice(1);

            const removeLeadingZero = document.getElementById('removeLeadingZero').checked;
            const removePlus33 = document.getElementById('removePlus33').checked;

            // Get visible columns in order
            const visibleColumns = AppState.columnOrder.filter(i => AppState.columnVisibility[i]);

            // Build new data
            const newHeaders = visibleColumns.map(i => headers[i]);
            const newRows = rows.map(row => {
                return visibleColumns.map(i => {
                    let value = row[i] || '';

                    if ((removeLeadingZero || removePlus33) && value) {
                        value = formatPhoneNumber(value, removeLeadingZero, removePlus33);
                    }

                    return value;
                });
            });

            AppState.reorganizedData = [newHeaders, ...newRows];

            const reorganizePreview = document.getElementById('reorganizePreview');
            const reorganizeTableContainer = document.getElementById('reorganizeTableContainer');

            reorganizeTableContainer.innerHTML = createTableHTML(AppState.reorganizedData);
            reorganizePreview.classList.remove('hidden');

            updateExportTab();
            hideLoader();

            let message = 'Réorganisation appliquée avec succès!';
            if (removeLeadingZero && removePlus33) {
                message = 'Réorganisation appliquée avec suppression du "0" et "+33" !';
            } else if (removeLeadingZero) {
                message = 'Réorganisation appliquée avec suppression du premier "0" !';
            } else if (removePlus33) {
                message = 'Réorganisation appliquée avec suppression du "+33" !';
            }
            showAlert(message, 'success');
        }, 100);
    });

    // Proceed to export button
    const proceedToExportBtn = document.getElementById('proceedToExportBtn');
    if (proceedToExportBtn) {
        proceedToExportBtn.addEventListener('click', () => {
            switchTab('export');
        });
    }
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

        showLoader();

        setTimeout(() => {
            exportFile(data, fileName, format);
            hideLoader();
        }, 100);
    });
}

function updateExportTab() {
    const data = AppState.reorganizedData || AppState.mergedData;

    if (!data) return;

    const noDataMessage = document.getElementById('noDataMessage');
    const exportOptions = document.getElementById('exportOptions');
    const exportTableContainer = document.getElementById('exportTableContainer');
    const exportStats = document.getElementById('exportStats');

    noDataMessage.classList.add('hidden');
    exportOptions.classList.remove('hidden');

    exportTableContainer.innerHTML = createTableHTML(data);
    exportStats.textContent = `${data.length - 1} lignes • ${data[0].length} colonnes`;
}

function exportFile(data, fileName, format) {
    try {
        // Prepare data with apostrophe prefix for large numbers (SIRET, phone numbers, etc.)
        // This prevents Excel from converting them to scientific notation or truncating them
        const formattedData = data.map((row) =>
            row.map(cell => {
                if (cell === null || cell === undefined) return '';
                const value = String(cell).trim();
                // Ajouter apostrophe uniquement pour les nombres très longs (11 chiffres ou plus)
                // Cela protège le SIRET (14 chiffres) mais laisse les téléphones (9-10 chiffres) tels quels
                if (/^\d{11,}$/.test(value)) {
                    return `'${value}`;
                }
                return cell;
            })
        );

        const ws = XLSX.utils.aoa_to_sheet(formattedData);

        // Forcer le format texte pour les cellules avec de grands nombres (commençant par ')
        const range = XLSX.utils.decode_range(ws['!ref']);
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
                const cell = ws[cellAddress];
                if (cell && cell.v && typeof cell.v === 'string' && cell.v.startsWith("'")) {
                    cell.t = 's'; // Force string type
                    // On peut aussi essayer de définir le format de nombre en texte
                    cell.z = '@';
                }
            }
        }

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

        const wbout = XLSX.write(wb, writeOptions);

        const blobParts = [wbout];
        if (format === 'csv') {
            // Add BOM for Excel to recognize UTF-8
            const bom = new Uint8Array([0xEF, 0xBB, 0xBF]);
            blobParts.unshift(bom);
        }

        const blob = new Blob(blobParts, {
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
    initNavigation();
    initImportMerge();
    initReorganize();
    initExport();

    console.log('📊 Excel Manager initialized');
});
