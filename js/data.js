// --- ESTADO GLOBAL ---
    let GLOBAL_DATA = { 
        nuevo: [], htr: [], 
        excelTotals: { crudo: null, htr: null, global: null, ne: null },
        filterStatus: { crudoFiltered: false, htrFiltered: false },
        // emptyGroups stores placeholders for empty blocks that should still render
        emptyGroups: { crudo: [], htr: [] },
        // Balance2 datasets (opcional)
        balance2: { datos: [], produccion: [], solverRows: [] },
        currentTab: 'crudo'
    };

    // Mapeo cliente -> certificaciAn por defecto (puedes extenderlo)
    const CLIENT_DEFAULT_CERT = {
        'LLL': 'OCS'
    };

    function getClientCert(client) {
        if (!client) return null;
        const key = String(client).toUpperCase().trim();
        // Si existe mapeo explAcito, devolverlo
        if (CLIENT_DEFAULT_CERT[key]) return CLIENT_DEFAULT_CERT[key];
        // Por compatibilidad con la lAgica anterior, por defecto devolver GOTS
        return 'GOTS';
    }

    // Helpers para detectar variantes de 'ORG' / 'ORGANICO' y aAadir certificaciAn
    function isOrganicoText(s) {
        if (!s) return false;
        return /\b(ORGANICO|ORGANIC|ORG)\b/i.test(String(s));
    }

    function addCertToOrganico(text, certShort) {
        if (!text) return text;
        if (!certShort) return text;
        // Si ya contiene (OCS) o (GOTS), no modificar
        if (/\(OCS\)|\(GOTS\)/i.test(text)) return text;
        // Reemplazar primera apariciAn de ORG/ORGANICO/ORGANIC aAadiendo la certificaciAn
        return String(text).replace(/\b(ORGANICO|ORGANIC|ORG)\b/i, `$1 (${certShort})`);
    }

    // ID Generator for Drag & Drop
    function generateId() {
        return 'row_' + Math.random().toString(36).substr(2, 9);
    }

    function getCellBackgroundColor(workbook, sheetName, cellAddress, excelJsWorkbook) {
        try {
            if(!excelJsWorkbook) return null;
            
            const worksheet = excelJsWorkbook.getWorksheet(sheetName);
            if(!worksheet) return null;
            
            // Parsear direcciAn de celda (ej: "A1")
            const cell = worksheet.getCell(cellAddress);
            if(!cell || !cell.fill) return null;
            
            // ExcelJS proporciona fill con propiedades reales
            if(cell.fill.type === 'pattern' || cell.fill.type === 'solid') {
                const fgColor = cell.fill.fgColor;
                if(fgColor && fgColor.argb) {
                    // Formato ARGB - tomar Altimos 6 caracteres para RGB
                    let argb = String(fgColor.argb);
                    if(argb.length >= 6) {
                        return '#' + argb.slice(-6);
                    }
                }
            }
            
            return null;
        } catch(e) {
            return null;
        }
    }

    function switchModule(modId) {
        document.querySelectorAll('.module-view').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.module-btn').forEach(el => el.classList.remove('active'));
        document.getElementById(modId).classList.add('active');
        const btns = document.querySelectorAll('.module-btn');
        if(modId === 'mod-pcp') btns[0].classList.add('active');
        else if(modId === 'mod-balance') btns[1].classList.add('active');
        else if(modId === 'mod-titulo') btns[2].classList.add('active');
        else if(modId === 'mod-balance2') btns[3].classList.add('active');

        if(modId === 'mod-titulo') renderTituloModule();
        if(modId === 'mod-balance2') renderBalance2Module && renderBalance2Module();
    }

    function switchSubTab(modulePrefix, type) {
        if (modulePrefix === 'pcp') {
            document.getElementById('pcp-crudo-view').style.display = 'none';
            document.getElementById('pcp-htr-view').style.display = 'none';
            document.getElementById(`pcp-${type}-view`).style.display = 'block';
            const container = document.querySelector('#mod-pcp .sub-tabs');
            container.querySelectorAll('button').forEach(b => b.classList.remove('active'));
            container.querySelector(`.t-${type}`).classList.add('active');
        } else if (modulePrefix === 'bal') {
            document.getElementById('bal-crudo-view').style.display = 'none';
            document.getElementById('bal-htr-view').style.display = 'none';
            document.getElementById(`bal-${type}-view`).style.display = 'block';
            const container = document.querySelector('#mod-balance .sub-tabs');
            container.querySelectorAll('button').forEach(b => b.classList.remove('active'));
            container.querySelector(`.t-${type}`).classList.add('active');
            GLOBAL_DATA.currentTab = type;
            renderBalanceModule();
        } else if (modulePrefix === 'bal2') {
            // Sub-tabs para el nuevo mAdulo 'Balance' (mod-balance2)
            document.getElementById('bal2-crudo-view').style.display = 'none';
            document.getElementById('bal2-htr-view').style.display = 'none';
            document.getElementById(`bal2-${type}-view`).style.display = 'block';
            const container2 = document.querySelector('#mod-balance2 .sub-tabs');
            if(container2) {
                container2.querySelectorAll('button').forEach(b => b.classList.remove('active'));
                container2.querySelector(`.t-${type}`) && container2.querySelector(`.t-${type}`).classList.add('active');
            }
            GLOBAL_DATA.currentTab = type;
        } else if (modulePrefix === 'tit') {
            document.getElementById('tit-crudo-view').style.display = 'none';
            document.getElementById('tit-mezcla-view').style.display = 'none';
            document.getElementById('tit-htr-view').style.display = 'none';
            document.getElementById(`tit-${type}-view`).style.display = 'block';
            const container = document.querySelector('#mod-titulo .sub-tabs');
            container.querySelectorAll('button').forEach(b => b.classList.remove('active'));
            container.querySelector(`.t-${type}`).classList.add('active');
            GLOBAL_DATA.currentTab = (type === 'htr') ? 'htr' : 'crudo';
            renderTituloModule();
        }
    }

    function switchInnerTab(innerType) {
        document.querySelectorAll('.nested-content').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.nested-btn').forEach(el => el.classList.remove('active'));
        document.getElementById(`inner-${innerType}-view`).classList.add('active');
        const btns = document.querySelectorAll('.nested-btn');
        if(innerType === 'pure') btns[0].classList.add('active');
        else if(innerType === 'mix') btns[1].classList.add('active');
        else if(innerType === 'summary') btns[2].classList.add('active');
    }

    document.getElementById('fileUpload').addEventListener('change', handleFile, false);

    function handleFile(e) {
        e.preventDefault();
        e.stopPropagation();
        const files = e.target.files || (e.dataTransfer && e.dataTransfer.files);
        if (!files || files.length === 0) return;
        
        // Validar que sea un archivo Excel
        const file = files[0];
        const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                        file.type === 'application/vnd.ms-excel' ||
                        file.name.endsWith('.xlsx') || 
                        file.name.endsWith('.xls');
        
        if (!isExcel) {
            console.warn('Por favor, selecciona un archivo Excel vAlido (.xlsx o .xls)');
            return;
        }
        
        const reader = new FileReader();
        reader.onload = function(e) {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array', cellFormula: true, cellStyles: true});
                
                // TambiAn leer con ExcelJS para obtener estilos
                const excelJsWorkbook = new ExcelJS.Workbook();
                excelJsWorkbook.xlsx.load(e.target.result).then(() => {
                    processData(workbook, excelJsWorkbook);
                }).catch(err => {
                    console.error('Error cargando con ExcelJS:', err);
                    processData(workbook, null);
                });
            } catch(err) {
                console.error('Error al procesar el archivo:', err);
            }
            // Resetear el input file para permitir cargar el mismo archivo otra vez
            document.getElementById('fileUpload').value = '';
        };
        reader.onerror = function(err) {
            console.error('Error al leer el archivo:', err);
            document.getElementById('fileUpload').value = '';
        };
        reader.readAsArrayBuffer(file);
    }

    // PREVENIR NAVEGACIAN ACCIDENTAL POR ARRASTRE DE ARCHIVOS A NIVEL GLOBAL
    // Esto previene que el navegador intente abrir archivos directamente
    document.addEventListener('dragover', function(e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'none';
        return false;
    }, false);

    document.addEventListener('dragleave', function(e) {
        e.preventDefault();
        return false;
    }, false);

    document.addEventListener('drop', function(e) {
        e.preventDefault();
        e.stopPropagation();
        
        // Solo procesar si son archivos Excel
        if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
            const file = e.dataTransfer.files[0];
            const isExcel = file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' || 
                            file.type === 'application/vnd.ms-excel' ||
                            file.name.endsWith('.xlsx') || 
                            file.name.endsWith('.xls');
            
            if (isExcel) {
                // Simular cambio de input file
                const input = document.getElementById('fileUpload');
                const dataTransfer = new DataTransfer();
                dataTransfer.items.add(file);
                input.files = dataTransfer.files;
                handleFile({target: {files: dataTransfer.files}});
            }
        }
        return false;
    }, false);

    function cleanNumber(val) {
        if (typeof val === 'number') return val;
        if (!val) return 0;
        let str = String(val).replace(/,/g, '').trim(); 
        return parseFloat(str) || 0;
    }

    function round0(num) { return num; } // Devolver valor sin modificar para cAlculos

    function findHeaderRow(jsonData) {
        for(let r=0; r<Math.min(jsonData.length, 30); r++){
            if(!jsonData[r]) continue;
            let ordenIdx = -1;
            let neIdx = -1;
            let obsIndices = []; // Buscar TODAS las columnas OBS
            
            // Buscar ORDEN, Ne y TODAS las columnas OBS en la misma fila
            for(let c=0; c<jsonData[r].length; c++){
                const cellUpper = String(jsonData[r][c]).toUpperCase().trim();
                if(cellUpper === "ORDEN") ordenIdx = c;
                if(cellUpper.includes("OBS") || cellUpper.includes("OBSERVACIONES")) {
                    obsIndices.push(c);
                }
                if(cellUpper === "NE" || cellUpper.startsWith("NE.")) {
                    neIdx = c;
                }
            }
            
            if(ordenIdx !== -1) {
                return {r: r, c: ordenIdx, obsIndices: obsIndices, neIdx: neIdx, found: true};
            }
        }
        return {r:0, c:0, obsIndices: [], neIdx: -1, found:false};
    }

    function processData(workbook, excelJsWorkbook) {
        GLOBAL_DATA = { 
            nuevo: [], htr: [], 
            nuevoOriginal: [], htrOriginal: [], // Guardar orden original para PCP
            excelTotals: { crudo: null, htr: null, global: null, ne: null },
            filterStatus: { crudoFiltered: false, htrFiltered: false },
            // Mantener lista de bloques vacAos que el usuario decidiA conservar
            emptyGroups: { crudo: [], htr: [] },
            currentTab: 'crudo'
        };
        
        let calcSumCrudo = 0;
        let calcSumHtr = 0;

        workbook.SheetNames.forEach(sheetName => {
            const upperName = sheetName.toUpperCase();
            if (upperName.includes("NUEVO") || upperName.includes("HTR")) {
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet, {header: 1});
                const anchor = findHeaderRow(jsonData);
                if(!anchor.found) return;

                const isHtr = upperName.includes("HTR");
                const idxOrden = anchor.c;
                const idxKg = idxOrden + 7;
                const idxNe = anchor.neIdx; // Usar el Andice encontrado en el header
                
                // Seleccionar la columna OBS que tiene datos
                let idxObs = -1;
                if(anchor.obsIndices && anchor.obsIndices.length > 0) {
                    // Si hay mAltiples columnas OBS, elegir la que tiene mAs datos
                    let maxDataCount = 0;
                    for(let obsCol of anchor.obsIndices) {
                        let dataCount = 0;
                        for(let i = anchor.r + 1; i < jsonData.length && i < anchor.r + 50; i++) {
                            const row = jsonData[i];
                            if(row[obsCol] && String(row[obsCol]).trim() !== "") {
                                dataCount++;
                            }
                        }
                        if(dataCount > maxDataCount) {
                            maxDataCount = dataCount;
                            idxObs = obsCol;
                        }
                    }
                }

                for(let i = anchor.r + 1; i < jsonData.length; i++) {
                    const row = jsonData[i];
                    let celdaOrden = row[idxOrden] ? String(row[idxOrden]).toUpperCase().trim() : "";

                    if (celdaOrden.includes("KG CRUDO") || celdaOrden.includes("TOTAL CRUDO")) {
                        GLOBAL_DATA.excelTotals.crudo = cleanNumber(row[idxKg]); continue;
                    }
                    if (celdaOrden.includes("KG HTR") || celdaOrden.includes("TOTAL HTR")) {
                        GLOBAL_DATA.excelTotals.htr = cleanNumber(row[idxKg]);
                        continue;
                    }
                    if (celdaOrden.includes("KG TOTAL") || celdaOrden.includes("TOTAL GENERAL") || celdaOrden.includes("SUMA TOTAL")) {
                        // Buscar Ne en las prAximas 5 filas
                        for(let nextRowIdx = i + 1; nextRowIdx < Math.min(i + 6, jsonData.length); nextRowIdx++) {
                            const nextRow = jsonData[nextRowIdx];
                            
                            // Buscar "Ne. =" especAficamente
                            let neEqualColIdx = -1;
                            for(let col = 0; col < nextRow.length; col++) {
                                const cellStr = nextRow[col] ? String(nextRow[col]).toUpperCase().trim() : "";
                                if(cellStr.includes("NE") && cellStr.includes("=")) {
                                    neEqualColIdx = col;
                                    break;
                                }
                            }
                            
                            if(neEqualColIdx >= 0) {
                                // Tomar el valor de la siguiente columna no vacAa
                                for(let col = neEqualColIdx + 1; col < nextRow.length; col++) {
                                    const cellVal = nextRow[col];
                                    if(cellVal !== null && cellVal !== undefined && cellVal !== "") {
                                        const numVal = cleanNumber(cellVal);
                                        const roundedNe = Math.round(numVal);
                                        if(numVal > 0) {
                                            GLOBAL_DATA.excelTotals.ne = roundedNe;
                                            break;
                                        }
                                    }
                                }
                                break;
                            }
                        }
                        continue;
                    }
                    if (celdaOrden.includes("TOTAL") || celdaOrden === "ORDEN" || celdaOrden === "" || celdaOrden.includes("RESUMEN")) continue;

                    let kg = cleanNumber(row[idxKg]);
                    let rawHilado = row[idxOrden+5] ? String(row[idxOrden+5]).trim() : "";
                    let color = row[idxOrden+6] ? String(row[idxOrden+6]).trim() : "";
                    if(color && (color.includes("OCS") || color.includes("BCI")) && !rawHilado.includes(color)) {
                        rawHilado = rawHilado + " (" + color + ")";
                    }
                    
                    // Extraer observaciones de la columna seleccionada CON COLOR
                    let obs = "";
                    let obsBgColor = null;
                    if(idxObs !== -1 && row[idxObs]) {
                        obs = String(row[idxObs]).trim();
                        
                        // Intentar extraer color de fondo de la celda
                        try {
                            const cellAddress = XLSX.utils.encode_cell({r: i, c: idxObs});
                            obsBgColor = getCellBackgroundColor(workbook, sheetName, cellAddress, excelJsWorkbook);
                        } catch(e) {
                            // Ignorar errores de extracciAn de estilos
                        }
                    }
                    
                    // Extraer color de la celda del MATERIAL (PCP FUENTE)
                    let materialBgColor = null;
                    try {
                        const cellAddress = XLSX.utils.encode_cell({r: i, c: idxOrden+5});
                        materialBgColor = getCellBackgroundColor(workbook, sheetName, cellAddress, excelJsWorkbook);
                    } catch(e) {
                        // Ignorar errores de extracciAn de estilos
                    }

                    // --- NUEVA REGLA: Insertar (OCS) o (GOTS) inmediatamente despuAs de "ORGANICO" ---
                    try {
                        const clienteStr = (row[idxOrden+1] || "").toString().trim();
                        const certShort = getClientCert(clienteStr); // 'OCS' or 'GOTS'
                        // Append certification after ORG/ORGANICO/ORGANIC where appropriate
                        rawHilado = addCertToOrganico(rawHilado, certShort);
                    } catch(e) {}
                    // -------------------------------------------------------------------------------

                    // --- ASIGNACIAN DE GRUPO INICIAL ---
                    let cleanRaw = rawHilado.toUpperCase().replace(/^\d+\/\d+\s+/, '').trim();
                    
                    // Detectar mezcla: si tiene porcentajes de participaciAn (con o sin parAntesis)
                    let isMezcla = /\d{1,3}\/\d{1,3}\s*%/.test(cleanRaw);
                    
                    let groupKey = cleanRaw;
                    if(!isMezcla) {
                        // CRUDO: usar token base
                        const baseToken = getBaseMaterialToken(rawHilado);
                        if (baseToken) {
                            groupKey = baseToken;
                        } else {
                            groupKey = "TANGUIS"; // fallback
                        }
                    } else {
                        // MEZCLA: usar nombre canAnico
                        groupKey = getCanonicalGroupName(cleanRaw);
                    }

                    let item = {
                        _id: generateId(),
                        group: groupKey,
                        // TAtulo (ej: "30/1") extraAdo del hilado para mAdulo TAtulo
                        titulo: (rawHilado.match(/^\s*(\d+\/\d+)/) ? rawHilado.match(/^\s*(\d+\/\d+)/)[1] : 'SIN TITULO'),
                        isMezcla: isMezcla,
                        orden: row[idxOrden],
                        cliente: row[idxOrden+1],
                        temporada: row[idxOrden+2],
                        rsv: row[idxOrden+3], 
                        op: row[idxOrden+4],
                        hilado: rawHilado, 
                        colorText: color,
                        tipo: row[idxOrden+6],
                        kg: kg,
                        obs: obs,
                        obsBgColor: obsBgColor,
                        materialBgColor: materialBgColor,
                        highlight: false
                    };

                    if(isHtr) {
                        GLOBAL_DATA.htr.push(item);
                        GLOBAL_DATA.htrOriginal.push({...item}); // Copia para mantener orden original
                        calcSumHtr += kg;
                    } else {
                        GLOBAL_DATA.nuevo.push(item);
                        GLOBAL_DATA.nuevoOriginal.push({...item}); // Copia para mantener orden original
                        calcSumCrudo += kg;
                    }
                }
            }
        });

        // 1. PRIMERO CALCULAMOS EL SEMAFORO (HIGHLIGHT) EN EL ORDEN ORIGINAL
        GLOBAL_DATA.excelTotals.global = (GLOBAL_DATA.excelTotals.crudo || 0) + (GLOBAL_DATA.excelTotals.htr || 0);
        identifyIncludedRows(GLOBAL_DATA.nuevo, GLOBAL_DATA.excelTotals.crudo);
        identifyIncludedRows(GLOBAL_DATA.htr, GLOBAL_DATA.excelTotals.htr);

        // 2. COPIAR HIGHLIGHTS AL ARRAY ORIGINAL (PCP USA ESTE)
        for(let i=0; i<GLOBAL_DATA.nuevo.length; i++) {
            GLOBAL_DATA.nuevoOriginal[i].highlight = GLOBAL_DATA.nuevo[i].highlight;
        }
        for(let i=0; i<GLOBAL_DATA.htr.length; i++) {
            GLOBAL_DATA.htrOriginal[i].highlight = GLOBAL_DATA.htr[i].highlight;
        }

        // 3. LUEGO ORDENAMOS LOS DATOS PARA LA VISTA (BALANCE SOLO)
        sortDataArray(GLOBAL_DATA.nuevo);
        sortDataArray(GLOBAL_DATA.htr);

        renderPCPModule();
        renderBalanceModule();
    }

    function sortDataArray(arr) {
        const groupStats = {};
        arr.forEach(item => {
            if(!groupStats[item.group]) groupStats[item.group] = { minReal: 9e9, minEmitir: 9e9 };
            let itemOrden = parseFloat(String(item.orden).replace(/[^\d.]/g, '')) || 9e9;
            const opStr = String(item.op || "").toUpperCase();
            if (!opStr.includes("EMITIR")) {
                if (itemOrden < groupStats[item.group].minReal) groupStats[item.group].minReal = itemOrden;
            } else {
                if (itemOrden < groupStats[item.group].minEmitir) groupStats[item.group].minEmitir = itemOrden;
            }
        });

        arr.sort((a, b) => {
            const statA = groupStats[a.group];
            const statB = groupStats[b.group];
            const scoreA = (statA.minReal < 9e9) ? statA.minReal : (statA.minEmitir + 1e9);
            const scoreB = (statB.minReal < 9e9) ? statB.minReal : (statB.minEmitir + 1e9);
            if (scoreA !== scoreB) return scoreA - scoreB;

            const opA = String(a.op).toUpperCase().includes("EMITIR");
            const opB = String(b.op).toUpperCase().includes("EMITIR");
            if (opA !== opB) return opA ? 1 : -1;

            let oA = parseFloat(String(a.orden).replace(/[^\d.]/g, '')) || 0;
            let oB = parseFloat(String(b.orden).replace(/[^\d.]/g, '')) || 0;
            return oA - oB;
        });
    }

    function identifyIncludedRows(rows, targetTotal) {
        if (!targetTotal || targetTotal === 0) {
            rows.forEach(r => r.highlight = false); 
            return;
        }
        rows.forEach(r => r.highlight = false);
        const epsilon = 1.0; 
        
        let totalSum = rows.reduce((a,b) => a + b.kg, 0);
        // Si coincide, NO resaltar nada (Todo OK)
        if (Math.abs(totalSum - targetTotal) < epsilon) return;

        // Si NO coincide, buscar subconjunto y pintarlo
        let found = false;
        let sum = 0;
        for (let i = 0; i < rows.length; i++) {
            sum += rows[i].kg;
            if (Math.abs(sum - targetTotal) < epsilon) {
                for(let j=0; j<=i; j++) rows[j].highlight = true;
                found = true;
                break;
            }
        }
        if (!found) {
            for (let start = 0; start < rows.length; start++) {
                let currentSum = 0;
                for (let end = start; end < rows.length; end++) {
                    currentSum += rows[end].kg;
                    if (Math.abs(currentSum - targetTotal) < epsilon) {
                        for(let k=start; k<=end; k++) rows[k].highlight = true;
                        found = true;
                        break;
                    }
                    if (currentSum > targetTotal + epsilon) break;
                }
                if(found) break;
            }
        }
    }

    function renderPCPModule() {
        // Usar los arrays ORIGINALES (orden del Excel) para PCP
        document.getElementById('pcp-table-crudo').innerHTML = generatePCPTable(GLOBAL_DATA.nuevoOriginal, 'pcp-nuevo');
        document.getElementById('pcp-table-htr').innerHTML = generatePCPTable(GLOBAL_DATA.htrOriginal, 'pcp-htr');

        const et = GLOBAL_DATA.excelTotals;
        document.getElementById('pcp-summary').innerHTML = `
            <div class="kpi-box"><div class="kpi-lbl">Total Kg Crudo</div><div class="kpi-val">${fmt(et.crudo)}</div></div>
            <div class="kpi-box"><div class="kpi-lbl">Total Kg HTR</div><div class="kpi-val">${fmt(et.htr)}</div></div>
            <div class="kpi-box" style="border-color:#3b82f6"><div class="kpi-lbl">Total General</div><div class="kpi-val">${fmt(et.global)}</div></div>
            <div class="kpi-box"><div class="kpi-lbl">Ne</div><div class="kpi-val">${et.ne || '-'}</div></div>
        `;
        if (typeof applyTableFilters === 'function') applyTableFilters();
    }

    function generatePCPTable(rows, className) {
        if(rows.length === 0) return '<p class="empty-msg">Sin datos.</p>';
        let html = `<table class="${className}"><thead><tr><th>ORDEN</th><th>CLIENTE</th><th>TEMP</th><th>OP</th><th>HILADO</th><th>COLOR</th><th>KG SOL.</th><th style="width:150px;">OBSERVACIONES</th></tr></thead><tbody>`;
        rows.forEach(r => {
            let rowClass = r.highlight ? 'highlight-row' : '';
            const obsText = (r.obs && r.obs.trim() !== '') ? r.obs : '-';
            const obsBgStyle = r.obsBgColor ? `background-color: ${r.obsBgColor};` : '';
            const obsTextColor = r.obsBgColor ? 'color: #000; font-weight: 600;' : 'color:#475569;';
            const colorText = r.colorText || '-';
            // Hacer editable el hilado mediante un prompt para facilitar ediciAn rApida
            const safeHilado = (r.hilado || '').toString().replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            html += `<tr class="${rowClass}">
                <td>${r.orden || '-'}</td>
                <td>${r.cliente || '-'}</td>
                <td>${r.temporada || '-'}</td>
                <td>${r.op || '-'}</td>
                <td class="hilado-cell">
                    <div style="padding-right:28px;">${r.hilado || '-'}</div>
                    <button class="edit-pencil" title="Editar" onclick="promptEditHilado('${r._id}')">&#x270F;&#xFE0F;</button>
                </td>
                <td style="font-size:10px;">${colorText}</td>
                <td style="font-weight:bold;">${fmtDecimal(r.kg)}</td>
                <td style="text-align:left; font-size:11px; ${obsTextColor} white-space:normal; word-break:break-word; ${obsBgStyle}">${obsText}</td>
            </tr>`;
        });
        html += '</tbody></table>';
        return html;
    }

    // Abrir prompt para editar el hilado de una fila (rApido). Propaga cambios a todos los arrays y recalcula.
    function promptEditHilado(id) {
        try {
            // Buscar el item en alguno de los arrays
            let item = GLOBAL_DATA.nuevoOriginal.find(x => x._id === id) || GLOBAL_DATA.nuevo.find(x => x._id === id) || GLOBAL_DATA.htrOriginal.find(x => x._id === id) || GLOBAL_DATA.htr.find(x => x._id === id);
            if(!item) return alert('No se encontrA la fila.');
            const oldVal = item.hilado || '';
            const newVal = prompt('Editar Hilado', oldVal);
            if(newVal === null) return; // Cancel
            const trimmed = (newVal || '').toString().trim();
            if(trimmed === oldVal) return; // Sin cambios
            applyHiladoChange(oldVal, trimmed);
        } catch(e) { console.error('promptEditHilado error', e); }
    }

    // Recalcula y aplica el nuevo hilado a todas las ocurrencias iguales en los arrays globales
    function applyHiladoChange(oldHilado, newHilado) {
        if(!oldHilado || !newHilado) return;
        let updatedCount = 0;
        function updateArray(arr) {
            for(let i=0;i<arr.length;i++){
                if(String(arr[i].hilado || '') === String(oldHilado)){
                    arr[i].hilado = newHilado;
                    // Recalcular campos derivados
                    recalcItemFields(arr[i]);
                    updatedCount++;
                }
            }
        }
        updateArray(GLOBAL_DATA.nuevo);
        updateArray(GLOBAL_DATA.nuevoOriginal);
        updateArray(GLOBAL_DATA.htr);
        updateArray(GLOBAL_DATA.htrOriginal);

        // Re-evaluar highlights usando totales existentes
        identifyIncludedRows(GLOBAL_DATA.nuevo, GLOBAL_DATA.excelTotals.crudo);
        identifyIncludedRows(GLOBAL_DATA.htr, GLOBAL_DATA.excelTotals.htr);

        // Reordenar y volver a renderizar todo
        sortDataArray(GLOBAL_DATA.nuevo);
        sortDataArray(GLOBAL_DATA.htr);
        renderPCPModule();
        renderBalanceModule();
        renderTituloModule();
        updateFooterTotals();

        alert('Hilado actualizado en ' + updatedCount + ' filas.');
    }

    // Recalcula `group`, `isMezcla` y `titulo` a partir de `hilado` en el item
    function recalcItemFields(item) {
        try {
            const rawHilado = (item.hilado || '').toString();
            let cleanRaw = rawHilado.toUpperCase().replace(/^\d+\/\d+\s+/, '').trim();
            let isMez = /\d{1,3}\/\d{1,3}\s*%/.test(cleanRaw);
            item.isMezcla = isMez;
            if(!isMez) {
                const baseToken = getBaseMaterialToken(rawHilado);
                item.group = baseToken ? baseToken : 'TANGUIS';
            } else {
                item.group = getCanonicalGroupName(cleanRaw);
            }
            // titulo
            const m = rawHilado.match(/^\s*(\d+\/\d+)/);
            item.titulo = m ? m[1] : 'SIN TITULO';
        } catch(e) { console.error('recalcItemFields', e); }
    }

    // Normalizar tAtulos para agrupamiento en mAdulo TAtulo
    function normalizeTitulo(t) {
        if (!t) return 'SIN TITULO';
        const s = String(t).toUpperCase().trim();
        // Reglas especAficas solicitadas:
        // 40/1 VI -> agrupar como 36/1
        if (/^40\/1\b.*\bVI\b/i.test(s)) return '36/1';
        // 50/1 IV -> agrupar como 44/1
        if (/^50\/1\b.*\bIV\b/i.test(s)) return '44/1';
        // Quitar sufijos de calidad numerales raros y devolver bAsico (ej: '40/1 A' -> '40/1')
        const m = s.match(/^(\d+\/\d+)/);
        if (m) return m[1];
        return s;
    }

    // Extraer valor Ne desde un item (aplica mapeos especAficos y luego extracciAn)
    function getNeFromItem(it) {
        try {
            // Priorizar `hilado` porque contiene sufijos (ej. 'VI') que `titulo` pierde.
            const rawSrc = (it && ((it.hilado && it.hilado.toString()) || it.titulo || '')) || '';
            const s = String(rawSrc).toUpperCase().trim();

            // Reglas explAcitas del usuario (coberturas con y sin '/1')
            // 40/1 VI  -> 36
            if (/^\s*40(?:\/1)!\b.*\bVI\b/i.test(s)) return 36;
            // 50 IV -> 44
            if (/^\s*50(?:\/1)!\b.*\bIV\b/i.test(s)) return 44;

            // Intentar extraer formato 'NN/NN' -> devolver primer nAmero
            const m = s.match(/(\d+)\s*\/\s*\d+/);
            if (m) return parseInt(m[1], 10);

            // Si no hay '/', tomar primer nAmero de 1 a 3 dAgitos
            const m2 = s.match(/(\d{1,3})/);
            if (m2) return parseInt(m2[1], 10);
        } catch (e) { console.error('getNeFromItem', e); }
        return null;
    }

    function computeWeightedNe(arr) {
        let sumNeKg = 0, sumKg = 0;
        (arr || []).forEach(it => {
            const ne = getNeFromItem(it);
            const kg = Number(it.kg || 0);
            if (ne && kg > 0) {
                sumNeKg += ne * kg;
                sumKg += kg;
            }
        });
        if (sumKg === 0) return null;
        return sumNeKg / sumKg;
    }

    function isAlgodonText(value) {
        if (!value) return false;
        const norm = String(value)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toUpperCase();
        return /\b(ALGODON|PIMA|TANGUIS|UPLAND|COP|ORGANICO|OCS|GOTS|BCI|USTCP)\b/.test(norm);
    }

    function renderOtrosBlock(isHtr, allRows) {
        const rows = (allRows || []).filter(r => (r.group || '').toString().toUpperCase() === 'OTROS');
        const hasRows = rows.length > 0;
        // Detectar si hay COP ORGANICO LLL entre las filas OTROS (solo crudo)
        const groupHasCopOrgLll = (!isHtr) && rows.some(r => (getClientCert((r.cliente||'').toString().trim()) === 'OCS') && /COP\s*(ORGANICO|ORG|ORGANIC)/i.test(r.hilado || ''));
        // Calcular nAmero de columnas del encabezado para usar en colspan cuando estA vacAo
        let baseCols = 10; // ORDEN, CLIENTE, TEMP, RSV, OP, HILADO, COLOR, KG SOL., KG REQ., QQ REQ.
        if (groupHasCopOrgLll) baseCols += 2; // columnas adicionales para ORGANICO/TANGUIS

        const tableType = isHtr ? 'htr' : 'pure';
        let html = `<div class="table-wrap"><div class="bal-header-row"><span class="bal-title">MATERIAL: OTROS</span>`;
        if (!hasRows) {
            html += `<button style="margin-left:auto; padding:6px 12px; font-size:11px; background:#ef4444; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="confirmDeleteGroup('OTROS', '${tableType}')">Eliminar bloque vacAo</button>`;
        }
        html += `</div>`;
        html += `<table><thead><tr>
                    <th class="th-base">ORDEN</th><th class="th-base">CLIENTE</th><th class="th-base">TEMP</th><th class="th-base">RSV</th><th class="th-base">OP</th>
                    <th class="th-base">HILADO</th><th class="th-base">COLOR</th>
                    <th class="th-base">KG SOL.</th><th class="th-base">KG REQ.</th><th class="th-base">QQ REQ.</th>`;
        if (groupHasCopOrgLll) {
            html += `<th class="th-comp-2">QQ REQ ORGANICO 80%</th><th class="th-comp-1">QQ REQ TANGUIS 20%</th>`;
        }
        html += `</tr></thead><tbody>`;

        if (!hasRows) {
            html += `<tr ondragover="allowDrop(event)" ondragenter="this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="handleDropToOtros(event)">
                        <td colspan="${baseCols}"><div class="otros-drop-hint">Arrastra aquA filas para asignarlas a OTROS</div></td>
                    </tr>`;
        } else {
            rows.forEach(r => {
                const kgSol = round0(r.kg);
                const factor = isHtr ? 0.60 : 0.65;
                const kgReq = round0(kgSol / factor);
                const isAlgodon = isAlgodonText((r.hilado || r.group || ''));
                const qqReq = isAlgodon ? round0(kgReq / 46) : null;
                const dragAttrs = `draggable="true" ondragstart="handleDragStart(event, '${r._id}')" ondrop="handleDrop(event, '${r._id}')" ondragover="allowDrop(event)"`;
                const colorText = r.colorText || '-';

                const cli = (r.cliente || '').toString().trim().toUpperCase();
                const isCopOrgLll = (!isHtr) && (getClientCert(cli) === 'OCS') && /COP\s*(?:ORGANICO|ORG|ORGANIC)/i.test(r.hilado || '');
                let qqOrg = '';
                let qqTan = '';
                if (isCopOrgLll) {
                    qqOrg = round0(qqReq * 0.8);
                    qqTan = round0(qqReq * 0.2);
                }

                html += `<tr ${dragAttrs}>
                    <td>${r.orden || '-'}</td><td>${r.cliente || '-'}</td><td>${r.temporada || '-'}</td><td>${r.rsv || ''}</td><td>${r.op || ''}</td>
                    <td class="hilado-cell">${r.hilado || '-'}</td><td style="font-size:10px;">${colorText}</td>
                    <td>${fmtDecimal(kgSol)}</td>
                    <td style="background:#f8fafc; font-weight:bold;">${fmtDecimal(kgReq)}</td>
                    <td style="background:#f0f9ff;">${fmtDecimal(qqReq)}</td>`;
                if (groupHasCopOrgLll) html += `<td style="color:#15803d;">${qqOrg !== '' ? fmtDecimal(qqOrg) : ''}</td><td style="color:#b45309;">${qqTan !== '' ? fmtDecimal(qqTan) : ''}</td>`;
                html += `</tr>`;
            });
        }

        html += `</tbody>`;
        
        // Agregar fila TOTAL si hay filas
        if (hasRows) {
            let totalKgSol = 0;
            let totalKgReq = 0;
            let totalQQReq = 0;
            let totalQQOrg = 0;
            let totalQQTan = 0;
            let totalQQHas = false;
            
            rows.forEach(r => {
                const kgSol = round0(r.kg);
                const factor = isHtr ? 0.60 : 0.65;
                const kgReq = round0(kgSol / factor);
                const isAlgodon = isAlgodonText((r.hilado || r.group || ''));
                const qqReq = isAlgodon ? round0(kgReq / 46) : null;
                totalKgSol += kgSol;
                totalKgReq += kgReq;
                if (qqReq !== null && qqReq !== undefined) {
                    totalQQReq += qqReq;
                    totalQQHas = true;
                }
                
                const cli = (r.cliente || '').toString().trim();
                const isCopOrgLll = (!isHtr) && (getClientCert(cli) === 'OCS') && /COP\s*ORGANICO/i.test(r.hilado || '');
                if (isCopOrgLll) {
                    totalQQOrg += round0(qqReq * 0.8);
                    totalQQTan += round0(qqReq * 0.2);
                }
            });
            
            const totalQQDisplay = totalQQHas ? fmt(totalQQReq) : '-';
            html += `<tr class="bal-subtotal"><td colspan="7" style="text-align:right;">TOTAL OTROS:</td><td>${fmt(totalKgSol)}</td><td>${fmt(totalKgReq)}</td><td>${totalQQDisplay}</td>`;
            if (groupHasCopOrgLll) html += `<td>${fmt(totalQQOrg)}</td><td>${fmt(totalQQTan)}</td>`;
            html += `</tr>`;
        }
        
        html += `</table></div>`;
        return html;
    }

    

    
function fmt(n) {
        if (n === null || n === undefined) return '-';
        const num = Number(n);
        if (!isFinite(num)) return '-';
        const rounded = Math.round(num);
        if (rounded === 0) return '-';
        return rounded.toLocaleString('es-PE');
    }

    function fmtDecimal(n) {
        if (n === null || n === undefined) return '-';
        const num = Number(n);
        if (!isFinite(num)) return '-';
        const rounded = Math.round(num);
        if (rounded === 0) return '-';
        return rounded.toLocaleString('es-PE');
    }

    // Filtros por columna en tablas
    function applyTableFilters(root) {
        const moduleRoot = document.getElementById('mod-balance2');
        const scope = root || moduleRoot;
        if (!scope) return;
        const tables = scope.querySelectorAll('table');
        tables.forEach(table => {
            if (moduleRoot && !moduleRoot.contains(table)) return;
            const filterRow = table.querySelector('thead .table-filter-row');
            if (filterRow) filterRow.remove();
            if (table.dataset.filters) delete table.dataset.filters;
        });
    }

    const HEADER_FILTER_STATE = new WeakMap();
    let headerFiltersBound = false;

    function applyHeaderFilters(root) {
        const moduleRoot = document.getElementById('mod-balance2');
        const scope = root || moduleRoot;
        if (!scope) return;
        const tables = scope.querySelectorAll('table.data-filter');
        tables.forEach(table => {
            if (table.dataset.headerFilters === 'true') return;
            const thead = table.querySelector('thead');
            const headerRow = thead ? thead.querySelector('tr') : null;
            if (!headerRow) return;
            const ths = headerRow.querySelectorAll('th');
            ths.forEach((th, idx) => {
                if (th.querySelector('.th-filter-btn')) return;
                th.dataset.colIndex = String(idx);
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'th-filter-btn';
                btn.innerHTML = '&#9662;';
                btn.addEventListener('click', (e) => {
                    e.stopPropagation();
                    toggleHeaderFilterMenu(table, th, idx);
                });
                th.appendChild(btn);
                const menu = document.createElement('div');
                menu.className = 'filter-menu';
                menu.dataset.colIndex = String(idx);
                menu.style.display = 'none';
                menu.addEventListener('click', (e) => e.stopPropagation());
                th.appendChild(menu);
            });
            table.dataset.headerFilters = 'true';
        });

        if (!headerFiltersBound) {
            document.addEventListener('click', closeHeaderFilterMenus);
            headerFiltersBound = true;
        }
    }

    function toggleHeaderFilterMenu(table, th, idx) {
        const menu = th.querySelector('.filter-menu');
        if (!menu) return;
        const isOpen = menu.style.display === 'block';
        closeHeaderFilterMenus();
        if (isOpen) {
            menu.style.display = 'none';
            return;
        }
        buildHeaderFilterMenu(table, idx, menu, th);
        menu.style.display = 'block';
    }

    function closeHeaderFilterMenus() {
        const menus = document.querySelectorAll('#mod-balance2 .filter-menu');
        menus.forEach(m => { m.style.display = 'none'; });
    }

    function getHeaderFilterState(table) {
        let state = HEADER_FILTER_STATE.get(table);
        if (!state) {
            state = {};
            HEADER_FILTER_STATE.set(table, state);
        }
        return state;
    }

    function collectColumnValues(table, idx) {
        const values = [];
        const seen = new Set();
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const cell = row.cells[idx];
            if (!cell) return;
            const text = (cell.textContent || '').trim();
            if (seen.has(text)) return;
            seen.add(text);
            values.push(text);
        });
        return values;
    }

    function buildHeaderFilterMenu(table, idx, menu, th) {
        menu.innerHTML = '';
        const values = collectColumnValues(table, idx);
        const state = getHeaderFilterState(table);
        let selected = state[idx] ? new Set(state[idx]) : new Set(values);
        selected = new Set(Array.from(selected).filter(v => values.includes(v)));
        if (selected.size === 0 && values.length) selected = new Set(values);

        const title = document.createElement('div');
        title.className = 'filter-title';
        title.textContent = 'Filtrar';
        menu.appendChild(title);

        const allLabel = document.createElement('label');
        const allCb = document.createElement('input');
        allCb.type = 'checkbox';
        allCb.checked = selected.size === values.length;
        allCb.addEventListener('change', () => {
            if (allCb.checked) {
                selected = new Set(values);
            } else {
                selected = new Set();
            }
            state[idx] = Array.from(selected);
            syncHeaderFilterChecks(menu, selected);
            applyHeaderFiltersToTable(table);
            updateHeaderFilterIndicator(th, selected, values);
        });
        allLabel.appendChild(allCb);
        allLabel.appendChild(document.createTextNode('Todos'));
        menu.appendChild(allLabel);
        const hr = document.createElement('hr');
        menu.appendChild(hr);

        values.forEach(val => {
            const label = document.createElement('label');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.value = val;
            cb.checked = selected.has(val);
            cb.addEventListener('change', () => {
                if (cb.checked) selected.add(val);
                else selected.delete(val);
                state[idx] = Array.from(selected);
                allCb.checked = selected.size === values.length;
                applyHeaderFiltersToTable(table);
                updateHeaderFilterIndicator(th, selected, values);
            });
            const display = val === '' ? '-' : val;
            label.appendChild(cb);
            label.appendChild(document.createTextNode(display));
            menu.appendChild(label);
        });

        state[idx] = Array.from(selected);
        updateHeaderFilterIndicator(th, selected, values);
    }

    function syncHeaderFilterChecks(menu, selected) {
        const checks = menu.querySelectorAll('label input[type="checkbox"]');
        checks.forEach((cb, i) => {
            if (i === 0) return;
            cb.checked = selected.has(cb.value);
        });
    }

    function updateHeaderFilterIndicator(th, selected, values) {
        if (!th) return;
        if (selected.size === 0 || selected.size < values.length) th.classList.add('filter-active');
        else th.classList.remove('filter-active');
    }

    function applyHeaderFiltersToTable(table) {
        const state = getHeaderFilterState(table);
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            let show = true;
            for (const key in state) {
                const idx = Number(key);
                const allowed = new Set(state[key] || []);
                if (allowed.size === 0) { show = false; break; }
                const cell = row.cells[idx];
                const text = cell ? (cell.textContent || '').trim() : '';
                if (!allowed.has(text)) { show = false; break; }
            }
            row.style.display = show ? '' : 'none';
        });
    }

    function filterTableRows(table) {
        const inputs = table.querySelectorAll('thead .table-filter-input');
        const filters = Array.from(inputs).map(i => (i.value || '').toString().trim().toLowerCase());
        const rows = table.querySelectorAll('tbody tr');
        rows.forEach(row => {
            const cells = row.querySelectorAll('td');
            let show = true;
            for (let i = 0; i < filters.length; i++) {
                const f = filters[i];
                if (!f) continue;
                const cell = cells[i];
                const text = cell ? cell.textContent.toLowerCase() : '';
                if (!text.includes(f)) { show = false; break; }
            }
            row.style.display = show ? '' : 'none';
        });
    }

    // Modal functions
    function closeDetailModal() {
        document.getElementById('detailModal').classList.remove('show');
    }

    document.addEventListener('DOMContentLoaded', () => {
        if (typeof applyTableFilters === 'function') applyTableFilters();
    });

    window.onclick = function(event) {
        const modal = document.getElementById('detailModal');
        if (event.target == modal) {
            modal.classList.remove('show');
        }
    };
