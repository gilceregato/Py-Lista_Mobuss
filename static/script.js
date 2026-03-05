document.addEventListener('DOMContentLoaded', () => {
    // Elements
    const fileInput = document.getElementById('excel-file');
    const dropZone = document.getElementById('drop-zone');
    const fileNameDisplay = document.getElementById('selected-file-name');
    const uploadBtn = document.getElementById('upload-btn');
    const uploadLoader = document.getElementById('upload-loader');

    const step2 = document.getElementById('step-2');
    const step3 = document.getElementById('step-3');
    const step4 = document.getElementById('step-4');
    const phasesContainer = document.getElementById('fases-container');
    const disciplinesContainer = document.getElementById('disciplinas-container');
    const badge = document.getElementById('project-name-badge');

    const previewTableBody = document.querySelector('#preview-table tbody');
    const tableContainer = document.getElementById('table-container');

    // Default dates
    const dateInit = document.getElementById('data-inicial');
    const dateFinal = document.getElementById('data-final');

    const today = new Date();
    const nextYear = new Date();
    nextYear.setFullYear(today.getFullYear() + 1);
    const lastYear = new Date();
    lastYear.setFullYear(today.getFullYear() - 1);

    // Format YYYY-MM-DD safely
    const formatYMD = (d) => {
        const mm = String(d.getMonth() + 1).padStart(2, '0');
        const dd = String(d.getDate()).padStart(2, '0');
        return `${d.getFullYear()}-${mm}-${dd}`;
    };

    dateInit.value = formatYMD(today);
    dateInit.min = formatYMD(lastYear);

    dateFinal.value = formatYMD(nextYear);

    // State
    let currentFileId = null;

    // File Drop & Select
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });

    dropZone.addEventListener('dragleave', () => {
        dropZone.classList.remove('dragover');
    });

    dropZone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelect();
        }
    });

    fileInput.addEventListener('change', handleFileSelect);

    function handleFileSelect() {
        if (fileInput.files.length > 0) {
            const file = fileInput.files[0];
            if (file.name.endsWith('.xlsx')) {
                fileNameDisplay.textContent = file.name;
                fileNameDisplay.style.color = '#38bdf8';
                uploadBtn.disabled = false;
            } else {
                fileNameDisplay.textContent = 'Apenas arquivos .xlsx são permitidos.';
                fileNameDisplay.style.color = '#fca5a5';
                uploadBtn.disabled = true;
            }
        }
    }

    // Upload Action
    uploadBtn.addEventListener('click', async () => {
        const file = fileInput.files[0];
        if (!file) return;

        const formData = new FormData();
        formData.append('file', file);

        uploadBtn.disabled = true;
        uploadLoader.style.display = 'flex';
        fileNameDisplay.textContent = 'Enviando documento e lendo informações...';

        // Remove old errors if any
        const existingError = dropZone.parentElement.querySelector('.alert-error');
        if (existingError) existingError.remove();

        try {
            const res = await fetch('/upload', {
                method: 'POST',
                body: formData
            });

            const data = await res.json();

            if (data.success) {
                currentFileId = data.file_id;

                // Populate phases
                phasesContainer.innerHTML = '';
                data.fases.forEach((fase, idx) => {
                    const label = document.createElement('label');
                    label.className = 'checkbox-item';
                    label.innerHTML = `
                        <input type="checkbox" value="${fase}" id="fase-${idx}" checked> 
                        <span for="fase-${idx}">${fase}</span>
                    `;
                    phasesContainer.appendChild(label);
                });

                if (data.fases.length === 0) {
                    phasesContainer.innerHTML = '<span style="color:#fca5a5;">Nenhuma Fase encontrada no arquivo.</span>';
                }

                // Populate disciplines
                disciplinesContainer.innerHTML = '';
                if (data.disciplinas) {
                    data.disciplinas.forEach((disciplina, idx) => {
                        const label = document.createElement('label');
                        label.className = 'checkbox-item';
                        label.innerHTML = `
                            <input type="checkbox" value="${disciplina}" id="disciplina-${idx}" checked> 
                            <span for="disciplina-${idx}">${disciplina}</span>
                        `;
                        disciplinesContainer.appendChild(label);
                    });

                    if (data.disciplinas.length === 0) {
                        disciplinesContainer.innerHTML = '<span style="color:#fca5a5;">Nenhuma Disciplina encontrada no arquivo.</span>';
                    }
                }

                // Update Project Name
                badge.textContent = `Empreendimento: ${data.nome_empreendimento}`;
                badge.style.display = 'inline-block';

                // Pre-fill the oldest date from the file as default data inicial
                if (data.data_inicial_padrao) {
                    dateInit.value = data.data_inicial_padrao;
                }

                // Enable Step 2 and Step 3
                step2.classList.remove('disabled');
                step3.classList.remove('disabled');
                step4.classList.remove('disabled');

                // Wire quick-select buttons for fases
                document.getElementById('fases-select-all').addEventListener('click', () => {
                    phasesContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
                });
                document.getElementById('fases-select-none').addEventListener('click', () => {
                    phasesContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
                });

                // Wire quick-select buttons for disciplinas
                document.getElementById('disciplinas-select-all').addEventListener('click', () => {
                    disciplinesContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = true);
                });
                document.getElementById('disciplinas-select-none').addEventListener('click', () => {
                    disciplinesContainer.querySelectorAll('input[type="checkbox"]').forEach(cb => cb.checked = false);
                });

                uploadBtn.innerHTML = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg> Arquivo Lido com Sucesso';
                uploadBtn.classList.replace('btn-primary', 'btn-success');
                fileNameDisplay.textContent = `Carregado: ${file.name}`;
                fileNameDisplay.style.color = '#6ee7b7';

            } else {
                showError(dropZone.parentElement, data.error || 'Erro ao processar arquivo.');
                uploadBtn.disabled = false;
                fileNameDisplay.textContent = file.name;
                fileNameDisplay.style.color = '#38bdf8';
            }

        } catch (err) {
            showError(dropZone.parentElement, 'Erro de conexão com o servidor.');
            uploadBtn.disabled = false;
            fileNameDisplay.textContent = file.name;
            fileNameDisplay.style.color = '#38bdf8';
        } finally {
            uploadLoader.style.display = 'none';
        }
    });

    function showError(parent, msg) {
        const div = document.createElement('div');
        div.className = 'alert alert-error mt-1';
        div.innerHTML = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="flex-shrink:0;"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg> <span>${msg}</span>`;
        parent.appendChild(div);
    }

    // Preview Action
    const previewBtn = document.getElementById('preview-btn');
    const previewLoader = document.getElementById('preview-loader');

    previewBtn.addEventListener('click', async () => {
        if (!currentFileId) return;

        const selectedFases = Array.from(phasesContainer.querySelectorAll('input:checked')).map(i => i.value);
        const selectedDisciplines = Array.from(disciplinesContainer.querySelectorAll('input:checked')).map(i => i.value);

        if (selectedFases.length === 0) {
            alert('Atenção: Por favor, selecione pelo menos uma fase de projeto para continuar.');
            return;
        }

        if (selectedDisciplines.length === 0) {
            alert('Atenção: Por favor, selecione pelo menos uma disciplina para continuar.');
            return;
        }

        const payload = {
            file_id: currentFileId,
            fases: selectedFases,
            disciplinas_selecionadas: selectedDisciplines,
            data_inicial: dateInit.value,
            data_final: dateFinal.value
        };

        previewBtn.disabled = true;
        previewLoader.style.display = 'flex';
        tableContainer.style.display = 'none';

        try {
            const res = await fetch('/preview', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const data = await res.json();

            if (data.success) {
                renderTable(data.data);
                tableContainer.style.display = 'block';
                document.getElementById('generate-btn').style.display = 'inline-flex';
                // Scroll smoothly to step 4
                step4.scrollIntoView({ behavior: 'smooth' });
            } else {
                alert(data.error || 'Erro ao gerar pré-visualização.');
            }
        } catch (err) {
            alert('Erro de conexão ao gerar a pré-visualização.');
        } finally {
            previewBtn.disabled = false;
            previewLoader.style.display = 'none';
        }
    });

    // Color helper for Liberado selects (SIM=green, NÃO=red)
    function applyLiberadoColor(select) {
        select.classList.remove('val-sim', 'val-nao');
        if (select.value === 'SIM') select.classList.add('val-sim');
        else if (select.value === 'NÃO') select.classList.add('val-nao');
    }

    function renderTable(dataArray) {
        previewTableBody.innerHTML = '';

        if (dataArray.length === 0) {
            previewTableBody.innerHTML = '<tr><td colspan="14" style="text-align: center; color: #94a3b8;">Nenhum dado encontrado com os filtros selecionados.</td></tr>';
            return;
        }

        dataArray.forEach((row, index) => {
            const tr = document.createElement('tr');
            tr.dataset.index = index;

            tr.innerHTML = `
                <td class="row-check col-check"><input type="checkbox" class="row-selector"></td>
                <td>${row['Disciplina'] || ''}</td>
                <td>${row['Fase'] || ''}</td>
                <td>${row['Código'] || ''}</td>
                <td>${row['Revisão'] || ''}</td>
                <td title="${row['Documento'] || ''}">${String(row['Documento'] || '')}</td>
                <td>${row['Extensão'] || ''}</td>
                <td>${row['Situação'] || ''}</td>
                <td title="${row['Título'] || ''}">${String(row['Título'] || '')}</td>

                <td class="col-liberado">
                    <select class="edit-orcam">
                        <option value="">-</option>
                        <option value="SIM">SIM</option>
                        <option value="NÃO">NÃO</option>
                    </select>
                </td>
                <td class="col-liberado">
                    <select class="edit-materiais">
                        <option value="">-</option>
                        <option value="SIM">SIM</option>
                        <option value="NÃO">NÃO</option>
                    </select>
                </td>
                <td><input type="text" class="edit-obs" placeholder="Observações" value=""></td>
                <td><input type="date" class="edit-prazo-entrega" value=""></td>
                <td><input type="date" class="edit-prazo-analise" value=""></td>
            `;

            previewTableBody.appendChild(tr);

            // Color-code the liberado selects
            const orcamSel = tr.querySelector('.edit-orcam');
            const matSel = tr.querySelector('.edit-materiais');
            orcamSel.addEventListener('change', () => applyLiberadoColor(orcamSel));
            matSel.addEventListener('change', () => applyLiberadoColor(matSel));
        });

        // Wire row-selector checkboxes for bulk selection
        previewTableBody.querySelectorAll('.row-selector').forEach(cb => {
            cb.addEventListener('change', updateBulkBar);
        });
    }

    // Bulk-edit logic
    const bulkEditBar = document.getElementById('bulk-edit-bar');
    const bulkCount = document.getElementById('bulk-count');
    const bulkApplyBtn = document.getElementById('bulk-apply-btn');
    const selectAllRows = document.getElementById('select-all-rows');

    function updateBulkBar() {
        const selected = previewTableBody.querySelectorAll('.row-selector:checked');
        const total = previewTableBody.querySelectorAll('.row-selector');

        // Highlight selected rows
        previewTableBody.querySelectorAll('tr[data-index]').forEach(tr => {
            const cb = tr.querySelector('.row-selector');
            tr.classList.toggle('row-selected', cb && cb.checked);
        });

        if (selected.length > 0) {
            bulkEditBar.style.display = 'flex';
            bulkCount.textContent = `${selected.length} linha(s) selecionada(s)`;
        } else {
            bulkEditBar.style.display = 'none';
        }

        // Sync the select-all state
        selectAllRows.indeterminate = (selected.length > 0 && selected.length < total.length);
        selectAllRows.checked = (selected.length === total.length && total.length > 0);
    }

    // Select-all header checkbox
    if (selectAllRows) {
        selectAllRows.addEventListener('change', () => {
            previewTableBody.querySelectorAll('.row-selector').forEach(cb => {
                const row = cb.closest('tr');
                if (row && row.style.display !== 'none') {
                    cb.checked = selectAllRows.checked;
                }
            });
            updateBulkBar();
        });
    }

    // Apply all 5 bulk fields at once to selected rows
    bulkApplyBtn && bulkApplyBtn.addEventListener('click', () => {
        const orcamVal = document.getElementById('bulk-orcam').value;
        const materiaisVal = document.getElementById('bulk-materiais').value;
        const obsVal = document.getElementById('bulk-obs').value;
        const prazoEntVal = document.getElementById('bulk-prazo-entrega').value;
        const prazoAnaVal = document.getElementById('bulk-prazo-analise').value;

        previewTableBody.querySelectorAll('tr[data-index]').forEach(tr => {
            const cb = tr.querySelector('.row-selector');
            if (!cb || !cb.checked) return;

            if (orcamVal) { const s = tr.querySelector('.edit-orcam'); s.value = orcamVal; applyLiberadoColor(s); }
            if (materiaisVal) { const s = tr.querySelector('.edit-materiais'); s.value = materiaisVal; applyLiberadoColor(s); }
            if (obsVal !== '') tr.querySelector('.edit-obs').value = obsVal;
            if (prazoEntVal) tr.querySelector('.edit-prazo-entrega').value = prazoEntVal;
            if (prazoAnaVal) tr.querySelector('.edit-prazo-analise').value = prazoAnaVal;
        });
    });

    // Generate Action
    const generateBtn = document.getElementById('generate-btn');
    const generateLoader = document.getElementById('generate-loader');
    const downloadSection = document.getElementById('download-section');
    const downloadLink = document.getElementById('download-link');

    generateBtn.addEventListener('click', async () => {
        if (!currentFileId) return;

        // Collect phases and disciplines
        const selectedFases = Array.from(phasesContainer.querySelectorAll('input:checked')).map(i => i.value);
        const selectedDisciplines = Array.from(disciplinesContainer.querySelectorAll('input:checked')).map(i => i.value);

        // Collect Edited Table Data (only visible rows, in DOM order)
        const rows = previewTableBody.querySelectorAll('tr[data-index]');
        const tabelaEditada = [];

        rows.forEach(tr => {
            tabelaEditada.push({
                'Liberado para Orçamento': tr.querySelector('.edit-orcam').value,
                'Liberado para compra de materiais?': tr.querySelector('.edit-materiais').value,
                'Observações': tr.querySelector('.edit-obs').value,
                'Prazo de entrega dos projetos': formatDateForExcel(tr.querySelector('.edit-prazo-entrega').value),
                'Prazo de análise dos projetos': formatDateForExcel(tr.querySelector('.edit-prazo-analise').value)
            });
        });

        // Helper func to convert HTML date to DD/MM/YYYY for nicer Excel output (optional)
        function formatDateForExcel(dateStr) {
            if (!dateStr) return '';
            const [y, m, d] = dateStr.split('-');
            return `${d}/${m}/${y}`;
        }

        const payload = {
            file_id: currentFileId,
            fases: selectedFases,
            disciplinas_selecionadas: selectedDisciplines,
            data_inicial: dateInit.value,
            data_final: dateFinal.value,
            tabela_editada: tabelaEditada
        };

        generateBtn.disabled = true;
        generateLoader.style.display = 'flex';
        downloadSection.style.display = 'none';

        try {
            const res = await fetch('/process', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            });

            const data = await res.json();

            if (data.success) {
                downloadLink.href = data.download_url;
                downloadSection.style.display = 'block';
            } else {
                alert(data.error || 'Erro ao gerar lista.');
            }
        } catch (err) {
            alert('Erro de conexão ao gerar a lista mestra.');
        } finally {
            generateBtn.disabled = false;
            generateLoader.style.display = 'none';
        }
    });

    // Table Sorting and Filter Dropdowns Logic
    const sortHeaders = document.querySelectorAll('.sortable');
    let sortDirection = 1; // 1 for asc, -1 for desc

    // Store active filters for each column: { '0': new Set(['ARQ', 'EST']), '1': new Set([...]) }
    let activeFilters = {};

    function getUniqueValuesInColumn(colIdx) {
        const rows = Array.from(previewTableBody.querySelectorAll('tr[data-index]'));
        const values = new Set();
        // +1 offset because col 0 is now the checkbox
        const actualColIdx = parseInt(colIdx) + 1;
        rows.forEach(r => {
            const cell = r.children[actualColIdx];
            const cellText = cell ? cell.textContent.trim() : '';
            if (cellText) values.add(cellText);
        });
        return Array.from(values).sort();
    }

    function applyFilters() {
        const rows = Array.from(previewTableBody.querySelectorAll('tr[data-index]'));

        rows.forEach(row => {
            let shouldShow = true;

            // Check against each column that has active filters
            for (const colIdx in activeFilters) {
                const allowedValues = activeFilters[colIdx];
                if (allowedValues && allowedValues.size > 0) {
                    // +1 offset because col 0 is now the checkbox
                    const actualColIdx = parseInt(colIdx) + 1;
                    const cell = row.children[actualColIdx];
                    const cellText = cell ? cell.textContent.trim() : '';
                    if (!allowedValues.has(cellText)) {
                        shouldShow = false;
                        break;
                    }
                }
            }

            row.style.display = shouldShow ? '' : 'none';
        });

        // Update active state of filter icons visually
        document.querySelectorAll('.filter-icon').forEach(icon => {
            const cIdx = icon.getAttribute('data-col');
            if (activeFilters[cIdx] && activeFilters[cIdx].size > 0) {
                icon.classList.add('active');
            } else {
                icon.classList.remove('active');
            }
        });
    }

    // Single floating filter panel
    const filterPanel = document.getElementById('filter-panel');
    let activePanelColIdx = null;

    function openFilterPanel(icon) {
        const colIdx = icon.getAttribute('data-col');

        // If clicking the same column, toggle close
        if (activePanelColIdx === colIdx && filterPanel.style.display === 'flex') {
            filterPanel.style.display = 'none';
            activePanelColIdx = null;
            icon.classList.remove('active');
            return;
        }

        activePanelColIdx = colIdx;

        // Position the panel below the icon
        const rect = icon.getBoundingClientRect();
        filterPanel.style.top = (rect.bottom + 4) + 'px';
        filterPanel.style.left = rect.left + 'px';

        // Build content
        const uniqueValues = getUniqueValuesInColumn(colIdx);
        const isFiltered = activeFilters[colIdx] && activeFilters[colIdx].size > 0;

        filterPanel.innerHTML = `
            <div class="filter-dropdown-actions">
                <button type="button" id="fp-select-all">Tudo</button>
                <button type="button" id="fp-select-none">Nenhum</button>
            </div>
            <div class="filter-list" id="fp-list"></div>
            <button type="button" id="fp-apply" class="filter-apply-btn">Aplicar</button>
        `;

        const listContainer = filterPanel.querySelector('#fp-list');
        uniqueValues.forEach(val => {
            const id = 'fp-cb-' + val.replace(/[^a-zA-Z0-9]/g, '_');
            const isChecked = !isFiltered || activeFilters[colIdx].has(val);
            const label = document.createElement('label');
            label.innerHTML = `<input type="checkbox" id="${id}" value="${val}" ${isChecked ? 'checked' : ''}> <span>${val}</span>`;
            listContainer.appendChild(label);
        });

        filterPanel.querySelector('#fp-select-all').addEventListener('click', (e) => {
            e.stopPropagation();
            activeFilters[colIdx] = new Set();
            applyFilters();
            filterPanel.style.display = 'none';
            activePanelColIdx = null;
        });

        filterPanel.querySelector('#fp-select-none').addEventListener('click', (e) => {
            e.stopPropagation();
            activeFilters[colIdx] = new Set(['__NENHUM__']);
            applyFilters();
            filterPanel.style.display = 'none';
            activePanelColIdx = null;
        });

        filterPanel.querySelector('#fp-apply').addEventListener('click', (e) => {
            e.stopPropagation();
            const selected = new Set();
            filterPanel.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => selected.add(cb.value));
            const total = filterPanel.querySelectorAll('input[type="checkbox"]').length;
            activeFilters[colIdx] = (selected.size === total) ? new Set() : selected;
            applyFilters();
            filterPanel.style.display = 'none';
            activePanelColIdx = null;
        });

        filterPanel.style.display = 'flex';

        // Update icon active state
        document.querySelectorAll('.filter-icon').forEach(ic => ic.classList.remove('active'));
        icon.classList.add('active');
    }

    // Attach click listeners to all filter icons
    document.querySelectorAll('.filter-icon').forEach(icon => {
        icon.addEventListener('click', (e) => {
            e.stopPropagation();
            openFilterPanel(icon);
        });
    });

    // Close panel on outside click
    document.addEventListener('click', (e) => {
        if (!e.target.closest('#filter-panel') && !e.target.closest('.filter-icon')) {
            filterPanel.style.display = 'none';
            activePanelColIdx = null;
            document.querySelectorAll('.filter-icon').forEach(ic => ic.classList.remove('active'));
        }
    });

    // Sorting functionality
    sortHeaders.forEach(header => {
        header.addEventListener('click', (e) => {
            // Prevent if clicked on filter icon or resizer
            if (e.target.closest('.filter-icon') || e.target.closest('.resize-handle') || e.target.closest('.filter-dropdown')) return;

            // +1 offset for checkbox column
            const colIdx = parseInt(header.getAttribute('data-col')) + 1;

            // Remove classes from others
            sortHeaders.forEach(h => {
                if (h !== header) {
                    h.classList.remove('asc', 'desc');
                }
            });

            // Toggle direction
            if (header.classList.contains('asc')) {
                header.classList.remove('asc');
                header.classList.add('desc');
                sortDirection = -1;
            } else {
                header.classList.remove('desc');
                header.classList.add('asc');
                sortDirection = 1;
            }

            const rows = Array.from(previewTableBody.querySelectorAll('tr[data-index]'));

            rows.sort((a, b) => {
                const cellA = a.children[colIdx].textContent.trim();
                const cellB = b.children[colIdx].textContent.trim();

                // Compare alfanumericamente
                return cellA.localeCompare(cellB, undefined, { numeric: true, sensitivity: 'base' }) * sortDirection;
            });

            // Reappend in new order
            rows.forEach(row => previewTableBody.appendChild(row));
        });
    });

    // Column Resizing functionality
    const headers = document.querySelectorAll('th');
    headers.forEach(th => th.classList.add('resizable')); // Make all resizable

    let currentResizer;
    let isResizing = false;
    let rect;

    headers.forEach(th => {
        const resizer = document.createElement('div');
        resizer.classList.add('resize-handle');
        th.appendChild(resizer);

        resizer.addEventListener('mousedown', function (e) {
            currentResizer = e.target;
            isResizing = true;

            const th = resizer.parentElement;
            rect = th.getBoundingClientRect();

            th.classList.add('active');
            currentResizer.classList.add('active');

            // Disable text selection while dragging
            document.body.style.userSelect = 'none';
            document.body.style.cursor = 'col-resize';
        });
    });

    document.addEventListener('mousemove', function (e) {
        if (!isResizing) return;

        const th = currentResizer.parentElement;
        const newWidth = e.clientX - rect.left;

        // Apply minimum width
        if (newWidth > 50) {
            th.style.width = newWidth + 'px';
            th.style.minWidth = newWidth + 'px';
        }
    });

    document.addEventListener('mouseup', function () {
        if (isResizing) {
            isResizing = false;
            currentResizer.classList.remove('active');
            currentResizer.parentElement.classList.remove('active');
            currentResizer = null;

            // Restore selection and cursor
            document.body.style.userSelect = '';
            document.body.style.cursor = '';
        }
    });

});
