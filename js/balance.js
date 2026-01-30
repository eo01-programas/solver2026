// Control de ventanas para el modulo Balance (mod-balance2)
function switchInnerTabBal2(innerType) {
    // ocultar todas las secciones especificas de balance2
    document.querySelectorAll('.nested2-content').forEach(el => {
        el.classList.remove('active');
        el.style.display = 'none';
    });
    // desactivar botones
    const btns = document.querySelectorAll('#mod-balance2 .nested-btn');
    btns.forEach(b => b.classList.remove('active'));

    // mostrar la solicitada
    const viewId = `inner2-${innerType}-view`;
    const view = document.getElementById(viewId);
    if (view) {
        view.classList.add('active');
        view.style.display = 'block';
    }

    // activar el boton correspondiente
    const btnMap = { datos: 0, produccion: 1, solver: 2 };
    const idx = btnMap[innerType];
    if (typeof idx === 'number') {
        if (btns[idx]) btns[idx].classList.add('active');
    }
}

function renderBalance2Module() {
    // renderizar contenido antes de mostrar tabs
    renderBalance2Datos();
    renderBalance2Produccion();
    renderBalance2Solver();
    // inicializar vista por defecto
    switchInnerTabBal2('datos');
    if (typeof applyTableFilters === 'function') applyTableFilters();
}

function renderBalance2Datos() {
    const el = document.getElementById('bal2-datos-area');
    if (!el) return;
    // Mantener el HTML estatico ya definido en index.html si no hay datos externos
    if (el.dataset.locked === 'true') return;
    if (GLOBAL_DATA && GLOBAL_DATA.balance2 && Array.isArray(GLOBAL_DATA.balance2.datos) && GLOBAL_DATA.balance2.datos.length) {
        el.dataset.locked = 'true';
        el.innerHTML = renderSimpleTable(GLOBAL_DATA.balance2.datos);
        if (typeof applyTableFilters === 'function') applyTableFilters(el);
    }
    if (typeof applyHeaderFilters === 'function') applyHeaderFilters(el);
}

const BAL2_PRODUCCION_COLUMNS = [
    { key: 'rpm', label: 'RPM' },
    { key: 'cd', label: 'CD' },
    { key: 'rpmHusos', label: 'RPM HUSOS' },
    { key: 'observaciones', label: 'OBSERVACIONES' },
    { key: 'velcd', label: 'VELCD' },
    { key: 'pinon', label: 'PIÑON' },
    { key: 'codigo', label: 'CODIGO' },
    { key: 'dia', label: 'DIA' },
    { key: 'ne', label: 'NE' },
    { key: 'material', label: 'MATERIAL' },
    { key: 'maqCount', label: 'Nª DE MAQ' },
    { key: 'htrsDisp', label: 'HTRS. DISP.' },
    { key: 'tipoWW', label: 'TIPO WW' },
    { key: 'huww', label: 'HUWW' },
    { key: 'humaq', label: 'HUMAQ' },
    { key: 'numHusHumaq', label: 'Nª DE HUS. HUMAQ' },
    { key: 'inicial', label: 'CONTADOR INICIAL' },
    { key: 'final', label: 'CONTADOR FINAL' },
    { key: 'puntosLeidos', label: 'PUNTOS LEIDOS' },
    { key: 'puntosHilan', label: 'PUNTOS HILAN(R)' },
    { key: 'factorHImp', label: 'FACTOR H. IMPRD' },
    { key: 'torsionXPulg', label: 'TOSION X/PULG' },
    { key: 'factorTorsion', label: 'FACTOR TORSION' },
    { key: 'factorContrc', label: 'FACTOR CONTRC' },
    { key: 'factorKgProd', label: 'FACTOR KG PROD' },
    { key: 'factorKg100', label: 'FACTOR KG 100%' },
    { key: 'kilosProd', label: 'KILOS PROD' },
    { key: 'kilosAlCien', label: 'KILOS AL CIEN' },
    { key: 'rendtComerc', label: 'RENDT COMERC' },
    
];

// Datos de Producción provistos por el usuario (TSV). Cada línea corresponde a una fila
const BAL2_PRODUCCION_TSV = `
152	0	12000			26	40COP1B	1			INGTD.(1)	1	1	572	572	572	404853	405677													1
210	0	12000			27	20COP1B				INGTD.(1)																		
168	0	12000			30	40PTN1B	1			INGTD.(2)	1	1	572	572	572	404853	405677													1
166	0	12000			30	36PTN1B				INGTD.(2)																		
185	0	12000			27	30COP1B	1			INGTD.(3)	1	1	572	572	572	404853	405677													1
210	0	12000			27	20COP1B				INGTD.(3)																		
154	0	11000			30	30LWCHB	1			INGTD.(4)	1	1	572	572	572	404853	405677													1
160	0								INGTD.(4)																		
130	0	11000			25	32POWOB	1			INGTD.(5)	1	1	572	572	572	404853	405677													1
196	0	11000			25	24OLW1B	1			INGTD.(5)																		
238		10200			58	14PM0OB	1			INGTD.(6)	1	1	572	572	572	404853	405677													1
222		8100			58	12PM0OB	1			INGTD.(6)																		
167	0	12500			30	40PS75B	1			INGTD.(7)	1	1	572	572	572	404853	405677													1
164	0	12000			30				INGTD.(7)																		
230	0	12000			29	24PM0OB	1			INGTD.(8)	1	1	572	572	572	404853	405677													1
170	0	12500			29	40PM0OB	1			INGTD.(8)	1	1	572	572	572	404853	405677													1
142		12500			40	55PM0OB	1			INGTD.(9)	1	1	572	572	572	404853	405677													1
239		11000			46	20PS01B				INGTD.(9)																		
142		12500			40	55PS75B	1			INGTD.(10)	1	1	572	572	572	404853	405677													1
146		12000						INGTD.(10)																		
142		11000			49	40PT01B	1			INGTD.(11)	1	1	620	620	620	754364	755115													1
168					53				INGTD.(11)																		
		11000			51	30COC1B	1			INGTD.(12)	1	1	620	620	620	754364	755115													1
								INGTD.(12)																		
153		12500			42	50PM0OB	1			INGTD.(13)	1	1	620	620	620	754364	755115													1
								INGTD.(13)																		
152		12500			40	44CPT1B	1			INGTD.(14)	1	1	620	620	620	754364	755115													1
152		12500			40	44PM0OB																		
		11000			44	40P401B	1			INGTD.(15)	1	1	620	620	620	754364	755115													1
				181.09				INGTD.(15)																		
		13000			60	55PS7HB	1			ZINZER  (16)	1	1	620	620	620	366936	373697													1
				8.90	52				ZINZER  (16)																		
		13000			23	40PS7HB	1			ZINZER  (17)	1	1	620	620	620	366936	373697													1
		12500			23	40PS7HB				ZINZER  (17)																		
		11000			24	40PX01B	1			ZINZER  (18)	1	1	620	620	620	366936	373697													1
								ZINZER  (18)																		
		11000			54	40M501B	1			ZINZER  (19)	1	1	620	620	620	366936	373697													1
				18.55				ZINZER  (19)																		
		11500			54	60PM0HB	1			ZINZER  (20)	1	1	620	620	620	366936	373697													1
				15.80	32				ZINZER  (20)																		
		13000			60	55PM0HB	1			ZINZER  (21)	1	1	620	620	620	366936	373697													1
		12500			60	55PM0HB				ZINZER  (21)																		
		13000			23	40PM0HB	1			ZINZER  (22)	1	1	620	620	620	366936	373697													1
		11000			23	40PM0HB				ZINZER  (22)																		
		11000			58	40CFL1B	1			ZINZER  (23)	1	1	620	620	620	366936	373697													1
				14.50	27				ZINZER  (23)																		
		11000			24	36PM0OB	1			ZINZER  (24)	1	1	620	620	620	366936	373697													1
		13000			26	36PM0OB	1			ZINZER  (24)																		
		12000			27	30A100B	1			ZINZER  (25)	1	1	620	620	620	366936	373697													1
		13000			28	30ABN3B				ZINZER  (25)																		
`;

function parseProductionTSV(tsv, cols) {
    const lines = tsv.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    const rows = lines.map(line => {
        const parts = line.split(/\t/);
        const obj = {};
        cols.forEach((c, i) => {
            obj[c.key] = (parts[i] !== undefined) ? parts[i].trim() : '';
        });
        return obj;
    });
    return rows;
}

const BAL2_PRODUCCION_ROWS = parseProductionTSV(BAL2_PRODUCCION_TSV, BAL2_PRODUCCION_COLUMNS);

function formatProductionValue(val) {
    if (val === null || val === undefined || val === '') return '';
    // If the value is a string that contains any letters, treat it as text
    if (typeof val === 'string') {
        const s = val.trim();
        if (/[A-Za-z]/.test(s)) return s;
        // only parse numeric-like strings (digits, commas, dots, minus)
        if (!/^[\d,\.\-\s]+$/.test(s)) return s;
    }
    const num = toNumber(val);
    if (num === null) return String(val);
    // Mostrar hasta 2 decimales como máximo, sin ceros finales innecesarios
    return num.toFixed(2).replace(/\.0+$/, '').replace(/(\.[0-9]*?)0+$/, '$1');
}

function getBalance2ProduccionRows() {
    if (GLOBAL_DATA && GLOBAL_DATA.balance2 && Array.isArray(GLOBAL_DATA.balance2.produccion) && GLOBAL_DATA.balance2.produccion.length) {
        return GLOBAL_DATA.balance2.produccion;
    }
    if (typeof BAL2_PRODUCCION_CSV_ROWS !== 'undefined' && Array.isArray(BAL2_PRODUCCION_CSV_ROWS) && BAL2_PRODUCCION_CSV_ROWS.length) {
        return BAL2_PRODUCCION_CSV_ROWS;
    }
    return BAL2_PRODUCCION_ROWS;
}

function renderBalance2Produccion() {
    const el = document.getElementById('bal2-produccion-area');
    if (!el) return;
    // Mantener el HTML estatico ya definido en index.html si no hay datos externos
    if (el.dataset.locked === 'true') return;
    const rows = getBalance2ProduccionRows();
    // calcular campos derivados para cada fila antes de renderizar
    const computedRows = (rows || []).map(r => {
        const row = Object.assign({}, r);
        try {
            // VELCD
            if (!row.velcd || String(row.velcd).trim() === '') {
                row.velcd = calculateVelcd(row.rpm || row.rpmCd, row.cd || row.cd);
            }
            // NE
            if (!row.ne || String(row.ne).trim() === '') {
                row.ne = extractNeFromCode(row);
            }
            // MATERIAL
            if (!row.material || String(row.material).trim() === '') {
                row.material = getMaterialFromCode(row);
            }
            // PUNTOS LEIDOS
            if (!row.puntosLeidos || String(row.puntosLeidos).trim() === '') {
                row.puntosLeidos = calculatePuntosLeidos(row.final || row.contadorFinal || row.contador_final, row.inicial || row.contadorInicial || row.contador_inicial);
            }
            // PUNTOS HILAN
            if (!row.puntosHilan || String(row.puntosHilan).trim() === '') {
                row.puntosHilan = calculatePuntosHilan(row.puntosLeidos, row.tipoWW || row.tipo_ww, row.humaq);
            }
            // FACTOR H. IMPRD (look up from TIPOS DE MATERIAL using EXTRAE)
            if (!row.factorHImp || String(row.factorHImp).trim() === '') {
                row.factorHImp = getFactorHImpFromCode(row);
            }
            // TORSION X/PULG
            if (!row.torsionXPulg || String(row.torsionXPulg).trim() === '') {
                row.torsionXPulg = calculateTorsionXPulg(row.pinon);
            }
            // FACTOR TORSION
            if (!row.factorTorsion || String(row.factorTorsion).trim() === '') {
                row.factorTorsion = calculateFactorTorsion(row.torsionXPulg, row.ne);
            }
            // FACTOR CONTRC
            if (!row.factorContrc || String(row.factorContrc).trim() === '') {
                row.factorContrc = calculateFactorContrc(row.factorTorsion);
            }
            // FACTORES KG
            if (!row.factorKgProd || String(row.factorKgProd).trim() === '') {
                row.factorKgProd = calculateFactorKgProd(row.factorHImp, row.factorContrc, row.numHusHumaq || row.numHus || row.num_hus, row.ne);
            }
            if (!row.factorKg100 || String(row.factorKg100).trim() === '') {
                row.factorKg100 = calculateFactorKg100(row.factorContrc, row.numHusHumaq || row.numHus || row.num_hus, row.ne);
            }
            // KILOS
            if (!row.kilosProd || String(row.kilosProd).trim() === '') {
                row.kilosProd = calculateKilosProd(row.puntosLeidos, row.factorKgProd);
            }
            if (!row.kilosAlCien || String(row.kilosAlCien).trim() === '') {
                row.kilosAlCien = calculateKilosAlCien(row.velcd, row.tipoWW, row.factorKg100);
            }
            // RENDT COMERC
            if (!row.rendtComerc || String(row.rendtComerc).trim() === '') {
                row.rendtComerc = calculateRendtComerc(row.kilosProd, row.kilosAlCien);
            }
            // PARA CONTAR MAQ removed per user request
        } catch (e) {
            console.error('Error calculando campos produccion', e);
        }
        return row;
    });

    let html = '<div class="table-wrap"><table class="production-table" style="min-width: 1600px;">';

    html += '<thead><tr>';
    BAL2_PRODUCCION_COLUMNS.forEach(col => {
        html += `<th class="th-base">${col.label}</th>`;
    });
    html += '</tr></thead>';

    html += '<tbody>';
    if (!computedRows.length) {
        html += `<tr><td colspan="${BAL2_PRODUCCION_COLUMNS.length}" style="text-align:center; padding:20px; color:#999;">Produccion: no hay datos cargados.</td></tr>`;
    } else {
        computedRows.forEach(row => {
            html += '<tr>';
            BAL2_PRODUCCION_COLUMNS.forEach(col => {
                // ensure CODIGO column shows original CSV value if available
                if (col.key === 'codigo') {
                    html += `<td>${formatProductionValue(row._rawCodigo || row[col.key])}</td>`;
                } else {
                    html += `<td>${formatProductionValue(row[col.key])}</td>`;
                }
            });
            html += '</tr>';
        });
    }
    html += '</tbody>';

    html += '</table></div>';

    el.innerHTML = html;

    if (typeof applyTableFilters === 'function') applyTableFilters(el);
}



// Funciones auxiliares para calculos de produccion
function getMaterialFromCode(rowOrCodigo) {
    // BUSCARV(EXTRAE(CODIGO;3;5);LISTAMATERIALES;2;0)
    // Accept either the codigo string or the parsed row object (which may contain _rawLine/_rawCodigo)
    let codigoStr = '';
    if (typeof rowOrCodigo === 'string') codigoStr = rowOrCodigo;
    else if (rowOrCodigo && typeof rowOrCodigo === 'object') codigoStr = rowOrCodigo._rawCodigo || rowOrCodigo.codigo || '';

    // if codigoStr is missing or too short, try to reconstruct from raw CSV line
    if ((!codigoStr || codigoStr.length < 4) && rowOrCodigo && rowOrCodigo._rawLine) {
        const raw = rowOrCodigo._rawLine;
        const m = raw.match(/\d{1,2}[A-Za-z0-9]{3,}/);
        if (m) codigoStr = m[0];
    }

    if (!codigoStr || codigoStr.length < 5) return '';
    const extracted = codigoStr.substring(2, 7);

    const materialTable = document.getElementById('tbl-materiales');
    if (!materialTable) return '';

    const rows = materialTable.querySelectorAll('tbody tr');
    for (let row of rows) {
        const codigo = row.querySelector('.material-code')?.textContent?.trim() || '';
        if (codigo === extracted) {
            const material = row.querySelector('.material-name')?.textContent?.trim() || '';
            return material;
        }
    }
    return '';
}

function getFactorHImpFromCode(rowOrCodigo) {
    // BUSCARV(EXTRAE(CODIGO;3;5);TIPOS DE MATERIAL;3;0)
    let codigoStr = '';
    if (typeof rowOrCodigo === 'string') codigoStr = rowOrCodigo;
    else if (rowOrCodigo && typeof rowOrCodigo === 'object') codigoStr = rowOrCodigo._rawCodigo || rowOrCodigo.codigo || '';

    if ((!codigoStr || codigoStr.length < 4) && rowOrCodigo && rowOrCodigo._rawLine) {
        const raw = rowOrCodigo._rawLine;
        const m = raw.match(/\d{1,2}[A-Za-z0-9]{3,}/);
        if (m) codigoStr = m[0];
    }

    if (!codigoStr || codigoStr.length < 5) return 0.6;
    const extracted = codigoStr.substring(2, 7);

    // Primero buscar en LISTA DE MATERIALES (#tbl-materiales)
    const listTable = document.getElementById('tbl-materiales');
    if (listTable) {
        const rowsList = listTable.querySelectorAll('tbody tr');
        for (let row of rowsList) {
            const codigo = row.querySelector('.material-code')?.textContent?.trim() || '';
            if (codigo === extracted) {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 3) {
                    const factorText = cells[2]?.textContent?.trim() || '';
                    if (factorText && factorText !== '-') return parseFloat(factorText) || 0.6;
                }
            }
        }
    }

    // Si no está en la lista principal, buscar en TIPOS DE MATERIAL (#tbl-materiales-tipo)
    const materialTable = document.getElementById('tbl-materiales-tipo');
    if (materialTable) {
        const rows = materialTable.querySelectorAll('tbody tr');
        for (let row of rows) {
            const codigo = row.querySelector('.material-code')?.textContent?.trim() || '';
            if (codigo === extracted) {
                const cells = row.querySelectorAll('td');
                if (cells.length >= 3) {
                    const factorText = cells[2]?.textContent?.trim() || '';
                    if (factorText && factorText !== '-') return parseFloat(factorText) || 0.6;
                }
            }
        }
    }

    return 0.6;
}

function extractNeFromCode(rowOrCodigo) {
    // Extraer los primeros 2 numeros de izquierda del CODIGO
    let codigoStr = '';
    if (typeof rowOrCodigo === 'string') codigoStr = rowOrCodigo;
    else if (rowOrCodigo && typeof rowOrCodigo === 'object') codigoStr = rowOrCodigo._rawCodigo || rowOrCodigo.codigo || '';
    if ((!codigoStr || codigoStr.length < 2) && rowOrCodigo && rowOrCodigo._rawLine) {
        const raw = rowOrCodigo._rawLine;
        const m = raw.match(/\d{1,2}[A-Za-z0-9]{3,}/);
        if (m) codigoStr = m[0];
    }
    const match = codigoStr.match(/^\d{2}/);
    return match ? parseInt(match[0]) : null;
}

function calculateVelcd(rpm, cd) {
    // VELCD = (RPM+CD)*0.025*PI()
    const rpmNum = toNumber(rpm) || 0;
    const cdNum = toNumber(cd) || 0;
    return (rpmNum + cdNum) * 0.025 * Math.PI;
}

function calculatePuntosLeidos(counterFinal, counterInitial) {
    // PUNTOS LEIDOS = (CONTADOR FINAL - CONTADOR INICIAL)*10
    const finalNum = toNumber(counterFinal) || 0;
    const initialNum = toNumber(counterInitial) || 0;
    return (finalNum - initialNum) * 10;
}

function calculatePuntosHilan(puntosLeidos, tipoWW, humaq) {
    // PUNTOS HILAN(R) = (PUNTOS LEIDOS * TIPO WW) / HUMAQ
    const puntosNum = toNumber(puntosLeidos) || 0;
    const tipoNum = toNumber(tipoWW) || 1;
    const humaqNum = toNumber(humaq) || 1;
    return humaqNum > 0 ? (puntosNum * tipoNum) / humaqNum : 0;
}

function calculateTorsionXPulg(pinon) {
    // TORSION X/PULG = 683/PINON
    const pinonNum = toNumber(pinon) || 1;
    return pinonNum > 0 ? 683 / pinonNum : 0;
}

function calculateFactorTorsion(torsionXPulg, ne) {
    // FACTOR TORSION = TORSION X/PULG / RAIZ(NE)
    const torsionNum = toNumber(torsionXPulg) || 0;
    const neNum = toNumber(ne) || 1;
    return neNum > 0 ? torsionNum / Math.sqrt(neNum) : 0;
}

function calculateFactorContrc(factorTorsion) {
    // FACTOR CONTRC = SI(FACTOR TORSION < 3.6; 1.5; SI(FACTOR TORSION >= 4; 2.5; 2))
    const factorNum = toNumber(factorTorsion) || 0;
    if (factorNum < 3.6) return 1.5;
    if (factorNum >= 4) return 2.5;
    return 2;
}

function calculateFactorKgProd(factorHImp, factorContrc, numHusHumaq, ne) {
    // FACTOR KG PROD = ((FACTOR H.IMP + FACTOR CONTRC - 100) / (-100)) * 0.592 * N DE HUS. HUMAQ / (NE * 1000)
    const himpNum = toNumber(factorHImp) || 0;
    const contrNum = toNumber(factorContrc) || 0;
    const husNum = toNumber(numHusHumaq) || 1;
    const neNum = toNumber(ne) || 1;
    
    const numerator = (himpNum + contrNum - 100) / (-100);
    return numerator * 0.592 * husNum / (neNum * 1000);
}

function calculateFactorKg100(factorContrc, numHusHumaq, ne) {
    // FACTOR KG 100% = ((FACTOR CONTRC - 100) / (-100)) * 0.592 * N DE HUS. HUMAQ / (NE * 1000)
    const contrNum = toNumber(factorContrc) || 0;
    const husNum = toNumber(numHusHumaq) || 1;
    const neNum = toNumber(ne) || 1;
    
    const numerator = (contrNum - 100) / (-100);
    return numerator * 0.592 * husNum / (neNum * 1000);
}

function calculateKilosProd(puntosLeidos, factorKgProd) {
    // KILOS PROD = PUNTOS LEIDOS * FACTOR KG PROD
    const puntosNum = toNumber(puntosLeidos) || 0;
    const factorNum = toNumber(factorKgProd) || 0;
    return puntosNum * factorNum;
}

function calculateKilosAlCien(velcd, tipoWW, factorKg100) {
    // KILOS AL CIEN = VELCD * TIPO WW * 60 * FACTOR KG 100%
    const velcdNum = toNumber(velcd) || 0;
    const tipoNum = toNumber(tipoWW) || 1;
    const factorNum = toNumber(factorKg100) || 0;
    return velcdNum * tipoNum * 60 * factorNum;
}

function calculateRendtComerc(kilosProd, kilosAlCien) {
    // RENDT COMERC = KILOS PROD / KILOS AL CIEN
    const prodNum = toNumber(kilosProd) || 0;
    const cientoNum = toNumber(kilosAlCien) || 1;
    return cientoNum > 0 ? prodNum / cientoNum : 0;
}

function renderBalance2Solver() {
    const el = document.getElementById('bal2-solver-area');
    if (!el) return;

    const rows = buildSolverRows();
    if (!rows || rows.length === 0) {
        el.innerHTML = '<p class="empty-msg">Solver: no hay datos cargados.</p>';
        return;
    }

    el.innerHTML = renderSolverTable(rows);
    if (typeof applyTableFilters === 'function') applyTableFilters(el);
}

function buildSolverRows() {
    // 1) Si hay datos ya preparados para solver, usarlos
    if (GLOBAL_DATA && GLOBAL_DATA.balance2 && Array.isArray(GLOBAL_DATA.balance2.solverRows) && GLOBAL_DATA.balance2.solverRows.length) {
        return GLOBAL_DATA.balance2.solverRows;
    }

    // 2) Caso por defecto: generar filas desde Agrup TÍTULO (no por material)
    const titles = getSolverTitleList();
    if (!titles.length) return [];

    const rows = titles.map(entry => {
        const name = (entry && typeof entry === 'object') ? entry.title : entry;
        const tipo = (entry && typeof entry === 'object' && entry.type) ? entry.type : 'CRUDO';
        const neVal = getNeForTitle(name, tipo);
        const info = getSolverTitleInfo(name, tipo);
        const prod = getProductionRowForTitle(name);
        // Preparar primer material limpio y NE ponderado para búsquedas
        const firstRawTmp = (info && info.items && info.items.length && (info.items[0].hilado || info.items[0].titulo || info.items[0].material)) ? (info.items[0].hilado || info.items[0].titulo || info.items[0].material) : name;
        const firstMaterialTmp = stripTituloFromHilado(firstRawTmp);
        const normTitleTmp = (typeof normalizeTitulo === 'function') ? normalizeTitulo(name) : name;
        let neWeightedTmp = null;
        if (normTitleTmp === '36/1') neWeightedTmp = 36;
        else if (normTitleTmp === '44/1') neWeightedTmp = 44;
        else if (info && info.items && info.items.length) neWeightedTmp = computeWeightedNe(info.items);

        // TPP = FACTOR H. IMPRD (lookup from production)
            let tpp = '';
            const rpm = getProductionValue(prod, ['rpm', 'velocidad', 'r.p.m']);
            const rpmHusos = getProductionValue(prod, ['rpm husos', 'rpmhusos', 'rpm_husos', 'rpm h.usos', 'rpmh']);
            try {
                const factorH = getProductionValue(prod, ['factor h. imprd', 'factor h imprd', 'factorhimprd', 'factor h.imp', 'factorhimpr']);
                if (factorH !== null && factorH !== undefined && String(factorH).trim() !== '') {
                    tpp = toNumber(factorH) !== null ? Math.round(toNumber(factorH) * 100) / 100 : '';
                } else {
                    tpp = '';
                }
            } catch (e) {
                console.error('Error obteniendo FACTOR H. IMPRD para TPP en buildSolverRows', e);
                tpp = '';
            }
            const porMin = getProductionValue(prod, ['metros por minuto', 'm/min', 'm por minuto', 'por minuto']);
            const hora100 = getProductionValue(prod, ['prd hora 100', 'prod hora 100', 'hora 100', 'hora100']);

            // COF = FACTOR TORSION: preferir tabla de referencia, luego Producción. NO calcular si no hay datos.
            let cof = '';
            try {
                let cofVal = null;
                
                // PASO 1: Buscar en tabla de referencia (PRIMERO Y PRINCIPAL)
                try {
                    const refResult = lookupFactorTorsion(name, firstMaterialTmp);
                    if (refResult && refResult.factorTorsion) {
                        cofVal = refResult.factorTorsion;
                        console.log(`COF obtenido de tabla de referencia: ${cofVal} (NE=${refResult.ne}, TIPO=${refResult.tipo})`);
                    }
                } catch(e) { console.error('Error buscando COF en tabla de referencia', e); }
                
                // PASO 2: Si no está en tabla, buscar en datos de Producción
                if (cofVal === null) {
                    try {
                        const neSearch = (neWeightedTmp !== null && !isNaN(neWeightedTmp)) ? Math.round(neWeightedTmp) : (neVal !== null && neVal !== undefined ? (isNaN(neVal) ? null : Math.round(neVal)) : null);
                        let rowByNeMat = null;
                        if (neSearch !== null) {
                            rowByNeMat = getProductionRowByNeAndMaterial(neSearch, firstMaterialTmp);
                        }
                        if (rowByNeMat) {
                            const cand = getProductionValue(rowByNeMat, ['factor torsion', 'factortorsion', 'factor tors', 'factor_torsion']);
                            if (cand !== null && cand !== undefined && String(cand).trim() !== '') cofVal = toNumber(cand);
                        }
                        if (cofVal === null) {
                            const cofCandidate = getProductionValue(prod, ['factor torsion', 'factortorsion', 'factor tors', 'factor_torsion']);
                            if (cofCandidate !== null && cofCandidate !== undefined && String(cofCandidate).trim() !== '') {
                                cofVal = toNumber(cofCandidate);
                            }
                        }
                    } catch(e) { console.error('Error buscando COF en producción', e); }
                }
                
                // PASO 3: NO CALCULAR - Si no encontramos en tabla o producción, dejar vacío
                // (Cambio: antes calculaba, ahora simplemente no lo calcula)
                
                // Formatear a 2 decimales si se encontró valor
                if (cofVal !== null && cofVal !== undefined && !isNaN(cofVal)) {
                    cof = cofVal.toFixed(2); // Muestra con 2 decimales
                } else {
                    cof = ''; // Vacío si no hay datos
                }
            } catch (e) {
                console.error('Error procesando COF en buildSolverRows', e);
                cof = '';
            }

        const hora100Num = toNumber(hora100);
        const eficPct = 78;
        const eficFactor = 0.78;
        const horaEfectiva = (hora100Num !== null) ? (hora100Num * eficFactor) : null;
        const diariaKg = (horaEfectiva !== null) ? (horaEfectiva * 24) : null;

        const kgSol = (info && info.kgSol !== null) ? info.kgSol : null;
        const kgReq = (info && info.kgReq !== null) ? info.kgReq : null;
        const diariaReq = (kgReq !== null) ? (kgReq / 24) : null;
        const horasMaquina = (diariaReq !== null && horaEfectiva) ? (diariaReq / horaEfectiva) : null;
        const maquinas = (diariaReq !== null && diariaKg) ? (diariaReq / diariaKg) : null;
        const prod24 = (diariaReq !== null) ? (diariaReq * 24) : null;

        // NE: usar lista de NEs del agrupamiento por título (si hay varios, mostrar como lista separada por comas)
        // Obtener NE ponderado desde Agrup Título (computeWeightedNe sobre los items del título)
        const normTitle = (typeof normalizeTitulo === 'function') ? normalizeTitulo(name) : name;
        let neWeighted = null;
        if (normTitle === '36/1') neWeighted = 36;
        else if (normTitle === '44/1') neWeighted = 44;
        else if (info && info.items && info.items.length) neWeighted = computeWeightedNe(info.items);
        // Mostrar en columna NE el valor proveniente de Agrup Título (redondeado). Si no existe, intentar valores discretos.
        const neList = getNEsForTitle(name, tipo);
        const neDisplay = (neWeighted !== null && !isNaN(neWeighted)) ? Math.round(neWeighted) : ((neList && neList.length === 1) ? neList[0] : '');
        // MATERIAL: tomar el primer hilado/título del bloque si existe
        const firstRaw = (info && info.items && info.items.length && (info.items[0].hilado || info.items[0].titulo || info.items[0].material)) ? (info.items[0].hilado || info.items[0].titulo || info.items[0].material) : name;
        const firstMaterial = stripTituloFromHilado(firstRaw);

        return {
            ne: neDisplay,
            neValue: (neWeighted !== null && !isNaN(neWeighted)) ? Math.round(neWeighted) : null,
            tipo: tipo,
            material: firstMaterial,
            kgSol: kgSol,
            cof: cof,
            tpp: tpp || '',
            cd: rpm || '',
            rpm: rpm || '',
            husos: rpmHusos || '',
            porMin: porMin || '',
            hora100: hora100 || '',
            efic: `${eficPct}%`,
            horaEfectiva: horaEfectiva,
            diariaKg: diariaKg,
            diariaReq: diariaReq,
            horasMaquina: horasMaquina,
            maquinas: maquinas,
            prod24: prod24,
            obs: ''
        };
    });
    // Ordenar filas por NE ascendente: preferir `neValue` numérico, si no existe extraer primer número de `ne`
    try {
        rows.sort((a, b) => {
            const getNum = r => {
                if (!r) return Infinity;
                if (r.neValue !== undefined && r.neValue !== null && !isNaN(Number(r.neValue))) return Number(r.neValue);
                const s = String(r.ne || '');
                const m = s.match(/(\d{1,3})/);
                return m ? Number(m[1]) : Infinity;
            };
            return getNum(a) - getNum(b);
        });
    } catch (e) { console.error('Error sorting solver rows by NE', e); }

    scaleMachinesToLimit(rows, 24);
    return rows;
}

function scaleMachinesToLimit(rows, limit) {
    try {
        if (!rows || !rows.length) return;
        const items = rows.map((r, idx) => ({
            idx,
            val: toNumber(r.maquinas)
        })).filter(x => x.val !== null && x.val > 0);
        if (!items.length) return;
        const total = items.reduce((s, x) => s + x.val, 0);
        if (total <= limit) return;

        const factor = limit / total;
        const scaled = items.map(x => {
            const scaledVal = x.val * factor;
            return {
                idx: x.idx,
                scaled: scaledVal,
                floor: Math.floor(scaledVal),
                frac: scaledVal - Math.floor(scaledVal)
            };
        });
        let used = scaled.reduce((s, x) => s + x.floor, 0);
        let remaining = Math.max(0, Math.round(limit - used));
        scaled.sort((a, b) => b.frac - a.frac);
        for (let i = 0; i < scaled.length && remaining > 0; i++) {
            scaled[i].floor += 1;
            remaining -= 1;
        }
        scaled.forEach(x => {
            rows[x.idx].maquinas = x.floor;
        });
    } catch (e) {
        console.error('scaleMachinesToLimit error', e);
    }
}

function getSolverMaterialInfo(materialName, tipo) {
    try {
        if (typeof getGroupsFromData !== 'function') return null;
        const crudoAll = (GLOBAL_DATA && GLOBAL_DATA.nuevo) ? GLOBAL_DATA.nuevo : [];
        const htrAll = (GLOBAL_DATA && GLOBAL_DATA.htr) ? GLOBAL_DATA.htr : [];
        const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
        const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;

        const groupedCrudo = getGroupsFromData(activeCrudo);
        const groupedHtr = getGroupsFromData(activeHtr);

        let group = null;
        let isHtr = false;
        const tipoNorm = String(tipo || '').toUpperCase();
        if (tipoNorm === 'HTR') {
            group = groupedHtr.find(g => (g.name || '').toString() === (materialName || '').toString());
            isHtr = true;
        } else if (tipoNorm === 'MEZCLA') {
            group = groupedCrudo.find(g => g.isMezcla && (g.name || '').toString() === (materialName || '').toString());
        } else if (tipoNorm === 'CRUDO') {
            group = groupedCrudo.find(g => !g.isMezcla && (g.name || '').toString() === (materialName || '').toString());
        }
        if (!group) {
            group = groupedCrudo.find(g => (g.name || '').toString() === (materialName || '').toString());
        }
        if (!group) {
            group = groupedHtr.find(g => (g.name || '').toString() === (materialName || '').toString());
            isHtr = true;
        }
        if (!group) return null;
        const kgSol = (group.items || []).reduce((s, it) => s + (Number(it.kg || 0)), 0);
        const factor = isHtr ? 0.60 : 0.65;
        const kgReq = (kgSol) ? (kgSol / factor) : null;
        return { kgSol, kgReq, isHtr };
    } catch (e) {
        console.error('getSolverMaterialInfo error', e);
        return null;
    }
}

function getProductionRowForMaterial(materialName) {
    try {
        const prod = getBalance2ProduccionRows() || [];
        if (!prod.length) return null;
        const sample = prod[0] || {};
        const keys = Object.keys(sample);
        const matKey = keys.find(k => k.toString().toLowerCase().includes('material')) || keys.find(k => k.toString().toLowerCase().includes('hilado'));
        if (!matKey) return null;
        const targetNorm = normalizeSolverKey(materialName);
         
         // 1) Intentar coincidencia exacta en columna MATERIAL
         let result = prod.find(r => normalizeSolverKey(r[matKey]) === targetNorm) || null;
         if (result) return result;
         
         // 2) Fallback: búsqueda aproximada por CODIGO
         // Buscar CODIGO que contenga o sea contenido por el material name
         const codigoKey = keys.find(k => k.toString().toLowerCase().includes('codigo'));
         if (codigoKey) {
             // Intentar: si material = "TANGUIS ORGANICO", buscar codigo que incluya "TA0OB" (aproximado)
             result = prod.find(r => {
                 const codigo = String(r[codigoKey] || '').trim();
                 const matNorm = normalizeSolverKey(String(r[matKey] || ''));
                 // coincidencia si el codigo (últimos 3-5 chars) incluye parte del nombre normalizado
                 // o si el nombre normalizado incluye fragmentos del código
                 if (codigo && targetNorm) {
                     const codNorm = normalizeSolverKey(codigo);
                     return codNorm.includes(targetNorm) || targetNorm.includes(codNorm);
                 }
                 return false;
             }) || null;
         }
         return result;
    } catch (e) {
        console.error('getProductionRowForMaterial error', e);
        return null;
    }
}

function normalizeCompareKey(value) {
    if (!value) return '';
    return String(value)
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toUpperCase()
        .replace(/[^A-Z0-9]+/g, '');
}

function getProductionValue(row, candidates) {
    if (!row) return null;
    const keys = Object.keys(row);
    const normCandidates = (candidates || []).map(c => normalizeCompareKey(c));
    for (const key of keys) {
        const normKey = normalizeCompareKey(key);
        for (const cand of normCandidates) {
            if (cand && normKey.includes(cand)) return row[key];
        }
    }
    return null;
}

function normalizeSolverKey(value) {
    return normalizeCompareKey(value);
}

function toNumber(val) {
    if (val === null || val === undefined || val === '') return null;
    const str = String(val).replace(/,/g, '');
    const match = str.match(/-?\d+(?:\.\d+)?/);
    if (!match) return null;
    const num = Number(match[0]);
    return isFinite(num) ? num : null;
}

// Eliminar prefijos de TÍTULO/hilado del texto de hilado para mostrar en MATERIAL
function stripTituloFromHilado(value) {
    if (value === null || value === undefined) return '';
    let s = String(value).trim();
    // Quitar prefijos como '40/1 VI', '40 VI', '50/1 IV', '36/1', etc. al inicio
    s = s.replace(/^\s*\d+(?:\/\d+)?(?:\s*(?:IV|VI|V|II|III|I))?\b\s*/i, '');
    return s.trim();
}

function getSolverMaterialList() {
    try {
        if (typeof getGroupsFromData !== 'function') return [];
        const crudoAll = (GLOBAL_DATA && GLOBAL_DATA.nuevo) ? GLOBAL_DATA.nuevo : [];
        const htrAll = (GLOBAL_DATA && GLOBAL_DATA.htr) ? GLOBAL_DATA.htr : [];

        const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
        const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;

        const groupedCrudo = getGroupsFromData(activeCrudo);
        const pureGroups = groupedCrudo.filter(g => !g.isMezcla && (g.name || '').toString().toUpperCase() !== 'OTROS');
        const mixGroups = groupedCrudo.filter(g => g.isMezcla && (g.name || '').toString().toUpperCase() !== 'OTROS');

        const groupedHtr = getGroupsFromData(activeHtr);
        const htrGroups = groupedHtr.filter(g => (g.name || '').toString().toUpperCase() !== 'OTROS');

        const ordered = [
            ...pureGroups.map(g => ({ name: g.name, type: 'CRUDO' })),
            ...mixGroups.map(g => ({ name: g.name, type: 'MEZCLA' })),
            ...htrGroups.map(g => ({ name: g.name, type: 'HTR' }))
        ];
        const names = [];
        const seen = new Set();
        ordered.forEach(g => {
            const n = (g && g.name) ? g.name : ''.toString().trim();
            const t = (g && g.type) ? g.type : '';
            if (!n || !t) return;
            const key = `${t}||${n}`;
            if (!seen.has(key)) { seen.add(key); names.push({ name: n, type: t }); }
        });

        return names;
    } catch (e) {
        console.error('getSolverMaterialList error', e);
        return [];
    }
}

// --- Nuevas utilidades para agrupar y resolver por TÍTULO ---
function getSolverTitleList() {
    try {
        const crudoAll = (GLOBAL_DATA && GLOBAL_DATA.nuevo) ? GLOBAL_DATA.nuevo : [];
        const htrAll = (GLOBAL_DATA && GLOBAL_DATA.htr) ? GLOBAL_DATA.htr : [];
        const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
        const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;
        const combined = (activeCrudo || []).concat(activeHtr || []);
        const seen = new Set();
        const list = [];
        combined.forEach(it => {
            try {
                const rawTitle = (it && (it.titulo || it.hilado)) ? (it.titulo || it.hilado) : 'SIN TITULO';
                const norm = normalizeTitulo(rawTitle);
                if (!norm) return;
                if (seen.has(norm)) return;
                seen.add(norm);
                let type = 'CRUDO';
                if (it && it.isMezcla) type = 'MEZCLA';
                if (activeHtr.includes(it)) type = 'HTR';
                list.push({ title: norm, type });
            } catch (e) {}
        });
        return list;
    } catch (e) { console.error('getSolverTitleList error', e); return []; }
}

function getNeForTitle(title, tipo) {
    try {
        const crudoAll = GLOBAL_DATA && GLOBAL_DATA.nuevo ? GLOBAL_DATA.nuevo : [];
        const htrAll = GLOBAL_DATA && GLOBAL_DATA.htr ? GLOBAL_DATA.htr : [];
        const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
        const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;
        const combined = (activeCrudo || []).concat(activeHtr || []);
        const items = (combined || []).filter(it => normalizeTitulo(it.titulo || it.hilado || '') === normalizeTitulo(title));
        if (!items.length) return null;
        return computeWeightedNe(items);
    } catch (e) { console.error('getNeForTitle error', e); return null; }
}

function getNEsForTitle(title, tipo) {
    try {
        const info = getSolverTitleInfo(title, tipo) || { items: [] };
        const normTitle = (typeof normalizeTitulo === 'function') ? normalizeTitulo(title) : title;
        // reglas especiales: si el título normalizado es 36/1 o 44/1 forzar esos NE
        if (normTitle === '36/1') return [36];
        if (normTitle === '44/1') return [44];
        const items = info.items || [];
        const neSet = new Set();
        items.forEach(it => {
            const n = getNeFromItem(it);
            if (n !== null && n !== undefined && !isNaN(n)) neSet.add(Math.round(n));
        });
        const arr = Array.from(neSet).sort((a,b) => a - b);
        return arr; // array of numbers
    } catch (e) { console.error('getNEsForTitle error', e); return []; }
}

function getSolverTitleInfo(title, tipo) {
    try {
        const crudoAll = GLOBAL_DATA && GLOBAL_DATA.nuevo ? GLOBAL_DATA.nuevo : [];
        const htrAll = GLOBAL_DATA && GLOBAL_DATA.htr ? GLOBAL_DATA.htr : [];
        const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
        const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;
        const combined = (activeCrudo || []).concat(activeHtr || []);
        const items = (combined || []).filter(it => normalizeTitulo(it.titulo || it.hilado || '') === normalizeTitulo(title));
        const kgSol = items.reduce((s, it) => s + (Number(it.kg || 0) || 0), 0);
        const kgReq = kgSol;
        return { kgSol, kgReq, items };
    } catch (e) { console.error('getSolverTitleInfo error', e); return { kgSol: null, kgReq: null, items: [] }; }
}

function getProductionRowForTitle(title) {
    try {
        const prod = getBalance2ProduccionRows() || [];
        if (!prod.length) return null;
        const sample = prod[0] || {};
        const keys = Object.keys(sample);
        const tituloKey = keys.find(k => k.toString().toLowerCase().includes('titulo'))
            || keys.find(k => k.toString().toLowerCase().includes('hilado'))
            || keys.find(k => k.toString().toLowerCase().includes('material'))
            || null;
        if (tituloKey) {
            const target = normalizeTitulo(title);
            let found = prod.find(r => normalizeTitulo(r[tituloKey] || r.hilado || r.titulo || '') === target) || null;
            if (found) return found;
        }
        // Fallback por CODIGO (búsqueda aproximada)
        const codigoKey = keys.find(k => k.toString().toLowerCase().includes('codigo'));
        if (codigoKey) {
            const tnorm = normalizeSolverKey(title);
            const found2 = prod.find(r => {
                const codigo = String(r[codigoKey] || '') || '';
                if (!codigo) return false;
                const codNorm = normalizeSolverKey(codigo);
                return codNorm.includes(tnorm) || tnorm.includes(codNorm);
            }) || null;
            return found2;
        } 
        return null;
    } catch (e) { console.error('getProductionRowForTitle error', e); return null; }
}

function extractProductionNe(row) {
    if (!row) return null;
    const keys = Object.keys(row || {});
    const neKey = keys.find(k => k.toString().toLowerCase().includes('ne'));
    if (!neKey) return null;
    const v = row[neKey];
    const n = toNumber(v);
    return n;
}

function getProductionRowByNeAndMaterial(ne, material) {
    try {
        if (ne === null || ne === undefined) return null;
        const prod = getBalance2ProduccionRows() || [];
        if (!prod.length) return null;
        const keysSample = Object.keys(prod[0] || {});
        const matKey = keysSample.find(k => k.toString().toLowerCase().includes('material')) || keysSample.find(k => k.toString().toLowerCase().includes('hilado')) || keysSample.find(k => k.toString().toLowerCase().includes('titulo')) || null;
        const normMat = normalizeCompareKey(material || '');
        // First try exact NE + material contains
        for (const r of prod) {
            const rNe = extractProductionNe(r);
            if (rNe === null) continue;
            if (Number(rNe) !== Number(ne)) continue;
            if (!matKey) return r;
            const rMat = String(r[matKey] || '');
            const rMatNorm = normalizeCompareKey(rMat);
            if (!normMat) return r;
            if (rMatNorm.includes(normMat) || normMat.includes(rMatNorm)) return r;
        }
        // fallback: any row with matching NE
        for (const r of prod) {
            const rNe = extractProductionNe(r);
            if (rNe === null) continue;
            if (Number(rNe) === Number(ne)) return r;
        }
        return null;
    } catch (e) { console.error('getProductionRowByNeAndMaterial error', e); return null; }
}

function getNeForMaterial(materialName, tipo) {
    try {
        if (typeof computeWeightedNe !== 'function' || typeof getGroupsFromData !== 'function') return null;
        const crudoAll = GLOBAL_DATA && GLOBAL_DATA.nuevo ? GLOBAL_DATA.nuevo : [];
        const htrAll = GLOBAL_DATA && GLOBAL_DATA.htr ? GLOBAL_DATA.htr : [];
        const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
        const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;
        const groupedCrudo = getGroupsFromData(activeCrudo);
        const groupedHtr = getGroupsFromData(activeHtr);
        let group = null;
        const tipoNorm = String(tipo || '').toUpperCase();
        if (tipoNorm === 'HTR') {
            group = groupedHtr.find(g => (g.name || '').toString() === (materialName || '').toString());
        } else if (tipoNorm === 'MEZCLA') {
            group = groupedCrudo.find(g => g.isMezcla && (g.name || '').toString() === (materialName || '').toString());
        } else if (tipoNorm === 'CRUDO') {
            group = groupedCrudo.find(g => !g.isMezcla && (g.name || '').toString() === (materialName || '').toString());
        }
        if (!group) {
            const combined = activeCrudo.concat(activeHtr);
            const groups = getGroupsFromData(combined);
            group = groups.find(g => (g.name || '').toString() === (materialName || '').toString());
        }
        if (!group || !group.items) return null;
        return computeWeightedNe(group.items || []);
    } catch (e) {
        console.error('getNeForMaterial error', e);
        return null;
    }
}

function renderSolverTable(rows) {
    const cols = [
        { key: 'ne', label: 'NE' },
        { key: 'tipo', label: 'TIPO' },
        { key: 'material', label: 'MATERIAL' },
        { key: 'kgSol', label: 'KG<br>SOL' },
        { key: 'cof', label: 'COF. E' },
        { key: 'tpp', label: 'T. P. P.' },
        { key: 'cd', label: 'C.D' },
        { key: 'rpm', label: 'R.P.M.' },
        { key: 'husos', label: 'HUSOS' },
        { key: 'porMin', label: 'POR<br>MINUTO' },
        { key: 'hora100', label: 'HORA 100%<br>KG.' },
        { key: 'efic', label: 'EFIC.' },
        { key: 'horaEfectiva', label: 'HORA<br>EFECTIVA' },
        { key: 'diariaKg', label: 'DIARIA<br>KG.' },
        { key: 'diariaReq', label: 'DIARIA<br>REQUERIDA' },
        { key: 'horasMaquina', label: 'HORAS<br>MAQUINA' },
        { key: 'maquinas', label: 'MAQUINAS' },
        { key: 'prod24', label: 'PROD. CON 24<br>DIAS (KG.)' },
        { key: 'obs', label: 'OBSERVACIONES' }
    ];

    let html = '<div class="table-wrap"><div class="bal-header-row"><div class="bal-title">CONTINUAS</div></div><table class="solver-table"><thead><tr>';
    cols.forEach(c => {
        html += `<th class=\"th-base\">${c.label}</th>`;
    });
    html += '</tr></thead><tbody>';
    rows.forEach(r => {
        html += '<tr>';
        cols.forEach(c => {
            const val = (r && r[c.key] !== undefined && r[c.key] !== null) ? r[c.key] : '';
            const cellClass = c.key === 'material' ? ' class="solver-material"' : '';
            // COF.E debe mostrar SIEMPRE 2 decimales: 4.00, 3.60, etc.
            const displayVal = (c.key === 'cof') ? val : formatSolverValue(val);
            html += `<td${cellClass}>${displayVal}</td>`;
        });
        html += '</tr>';
    });
    // Ne Promedio: usar `neValue` numérico si está disponible, fallback a `ne` si es numérico
    const neVals = rows.map(r => {
        if (r && r.neValue !== undefined && r.neValue !== null) return toNumber(r.neValue);
        return toNumber(r.ne);
    }).filter(v => v !== null);
    const neProm = neVals.length ? (neVals.reduce((a,b)=>a+b,0) / neVals.length) : null;
    const totalKgSol = rows.map(r => toNumber(r.kgSol)).filter(v => v !== null).reduce((a,b) => a + b, 0);
    const totalProd24 = rows.map(r => toNumber(r.prod24)).filter(v => v !== null).reduce((a,b) => a + b, 0);
    html += `<tr class="bal-subtotal">`;
    cols.forEach((c, idx) => {
        let cellVal = '';
        if (idx === 0) cellVal = neProm;
        else if (c.key === 'material') cellVal = 'NE PROMEDIO';
        else if (c.key === 'kgSol') cellVal = totalKgSol;
        else if (c.key === 'prod24') cellVal = totalProd24;
        html += `<td>${formatSolverValue(cellVal)}</td>`;
    });
    html += `</tr>`;

    html += '</tbody></table></div>';
    return html;
}

function renderSimpleTable(rows) {
    if (!rows || rows.length === 0) return '<p class="empty-msg">Sin datos.</p>';
    const cols = Object.keys(rows[0] || {});
    let html = '<div class="table-wrap"><table><thead><tr>';
    cols.forEach(c => { html += `<th class=\"th-base\">${c}</th>`; });
    html += '</tr></thead><tbody>';
    rows.forEach(r => {
        html += '<tr>';
        cols.forEach(c => { html += `<td>${formatSolverValue(r[c])}</td>`; });
        html += '</tr>';
    });
    html += '</tbody></table></div>';
    return html;
}

function formatSolverValue(val) {
    if (val === null || val === undefined) return '';
    if (typeof val === 'number') {
        if (typeof fmt === 'function') return fmt(val);
        // mostrar hasta 2 decimales, sin ceros finales
        return Number(val).toFixed(2).replace(/\.0+$/, '').replace(/(\.[0-9]*?)0+$/, '$1');
    }
    if (typeof val === 'string') {
        const trimmed = val.trim();
        if (trimmed && /^[\d,.\-]+$/.test(trimmed)) {
            const num = toNumber(trimmed);
            if (num !== null) {
                if (typeof fmt === 'function') return fmt(num);
                return Number(num).toFixed(2).replace(/\.0+$/, '').replace(/(\.[0-9]*?)0+$/, '$1');
            }
        }
    }
    return String(val);
}

// (Opciones de edicion/historial deshabilitadas por solicitud)

// --- CSV de Producción (alternativa): datos enviados por el usuario
const BAL2_PRODUCCION_CSV = `RPM,CD,RPM HUSOS,OBSERVACIONES,VELCD,PIÑON,CODIGO,DIA,NE,MATERIAL,Nª DE MAQ,HTRS. DISP.,TIPO WW,HUWW,HUMAQ,Nª DE HUS. HUMAQ,CONTADOR INICIAL,CONTADOR FINAL,PUNTOS LEIDOS,PUNTOS HILAN(R),FACTOR H. IMPRD,TOSION X/PULG,FACTOR TORSION,FACTOR CONTRC,FACTOR KG PROD,FACTOR KG 100%,KILOS PROD,KILOS AL CIEN,RENDT COMERC,PARA CONTAR MAQ
152,0,12000,,,26,40COP1B,1,,,INGTD.(1),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
210,0,12000,,,27,20COP1B,,,,INGTD.(1),,,,,,,,,,,,,,,,,,,
168,0,12000,,,30,40PTN1B,1,,,INGTD.(2),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
166,0,12000,,,30,36PTN1B,,,,INGTD.(2),,,,,,,,,,,,,,,,,,,
185,0,12000,,,27,30COP1B,1,,,INGTD.(3),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
210,0,12000,,,27,20COP1B,,,,INGTD.(3),,,,,,,,,,,,,,,,,,,
154,0,11000,,,30,30LWCHB,1,,,INGTD.(4),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
160,0,,,,,,,"",,INGTD.(4),,,,,,,,,,,,,,,,,,,
130,0,11000,,,25,32POWOB,1,,,INGTD.(5),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
196,0,11000,,,25,24OLW1B,1,,,INGTD.(5),,,,,,,,,,,,,,,,,,,
238,,10200,,,58,14PM0OB,1,,,INGTD.(6),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
222,,8100,,,58,12PM0OB,1,,,INGTD.(6),,,,,,,,,,,,,,,,,,,
167,0,12500,,,30,40PS75B,1,,,INGTD.(7),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
164,0,12000,,,30,,,,INGTD.(7),,,,,,,,,,,,,,,,,,,
230,0,12000,,,29,24PM0OB,1,,,INGTD.(8),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
170,0,12500,,,29,40PM0OB,1,,,INGTD.(8),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
142,,12500,,,40,55PM0OB,1,,,INGTD.(9),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
239,,11000,,,46,20PS01B,,,,INGTD.(9),,,,,,,,,,,,,,,,,,,
142,,12500,,,40,55PS75B,1,,,INGTD.(10),1,1,572,572,572,404853,405677,,,,,,,,,,,,1
146,,12000,,,,,,,INGTD.(10),,,,,,,,,,,,,,,,,,,
142,,11000,,,49,40PT01B,1,,,INGTD.(11),1,1,620,620,620,754364,755115,,,,,,,,,,,,1
168,,,,,53,,,,INGTD.(11),,,,,,,,,,,,,,,,,,,
,,11000,,,51,30COC1B,1,,,INGTD.(12),1,1,620,620,620,754364,755115,,,,,,,,,,,,1
,,,,,,,,,,INGTD.(12),,,,,,,,,,,,,,,,,,,
153,,12500,,,42,50PM0OB,1,,,INGTD.(13),1,1,620,620,620,754364,755115,,,,,,,,,,,,1
,,,,,,,,,,INGTD.(13),,,,,,,,,,,,,,,,,,,
152,,12500,,,40,44CPT1B,1,,,INGTD.(14),1,1,620,620,620,754364,755115,,,,,,,,,,,,1
152,,12500,,,40,44PM0OB,,,,,,,,,,,,,,,,,,,,,
,,,,,,,,,,INGTD.(14),,,,,,,,,,,,,,,,,,,
,,11000,,,44,40P401B,1,,,INGTD.(15),1,1,620,620,620,754364,755115,,,,,,,,,,,,1
,,,"181.09",,,,,,INGTD.(15),,,,,,,,,,,,,,,,,,,
,,13000,,,60,55PS7HB,1,,,ZINZER  (16),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,,"8.90",,,52,,,,ZINZER  (16),,,,,,,,,,,,,,,,,,,
,,13000,,,23,40PS7HB,1,,,ZINZER  (17),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,12500,,,23,40PS7HB,,,,ZINZER  (17),,,,,,,,,,,,,,,,,,,
,,11000,,,24,40PX01B,1,,,ZINZER  (18),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,,,,,,,,,ZINZER  (18),,,,,,,,,,,,,,,,,,,
,,11000,,,54,40M501B,1,,,ZINZER  (19),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,,"18.55",,,,,,ZINZER  (19),,,,,,,,,,,,,,,,,,,
,,11500,,,54,60PM0HB,1,,,ZINZER  (20),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,,"15.80",,,32,,,,ZINZER  (20),,,,,,,,,,,,,,,,,,,
,,13000,,,60,55PM0HB,1,,,ZINZER  (21),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,12500,,,60,55PM0HB,,,,ZINZER  (21),,,,,,,,,,,,,,,,,,,
,,13000,,,23,40PM0HB,1,,,ZINZER  (22),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,11000,,,23,40PM0HB,,,,ZINZER  (22),,,,,,,,,,,,,,,,,,,
,,11000,,,58,40CFL1B,1,,,ZINZER  (23),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,,"14.50",,,27,,,,ZINZER  (23),,,,,,,,,,,,,,,,,,,
,,11000,,,24,36PM0OB,1,,,ZINZER  (24),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,13000,,,26,36PM0OB,1,,,ZINZER  (24),,,,,,,,,,,,,,,,,,,
,,12000,,,27,30A100B,1,,,ZINZER  (25),1,1,620,620,620,366936,373697,,,,,,,,,,,,1
,,13000,,,28,30ABN3B,,,,ZINZER  (25),,,,,,,,,,,,,,,,,,,
`;

function parseProductionCSV(csv) {
    const lines = csv.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
    if (!lines.length) return [];
    const parseLine = (line) => {
        const out = [];
        let cur = '';
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') {
                if (inQuotes && line[i+1] === '"') { cur += '"'; i++; }
                else inQuotes = !inQuotes;
            } else if (ch === ',' && !inQuotes) { out.push(cur); cur = ''; }
            else cur += ch;
        }
        out.push(cur);
        return out.map(s => s.trim());
    };

    const headers = parseLine(lines[0]).map(h => h.normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z0-9]+/ig, '').toUpperCase());
    const colMap = {};
    BAL2_PRODUCCION_COLUMNS.forEach(c => {
        const norm = (c.label || c.key).toString().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^A-Z0-9]+/ig, '').toUpperCase();
        colMap[norm] = c.key;
    });

    const rows = lines.slice(1).map(line => {
        const parts = parseLine(line);
        const obj = {};
        BAL2_PRODUCCION_COLUMNS.forEach(c => obj[c.key] = '');
        parts.forEach((p, i) => {
            const h = headers[i] || '';
            const key = colMap[h];
            if (key) obj[key] = p;
            else {
                for (const norm in colMap) {
                    if (norm.includes(h) || h.includes(norm)) { obj[colMap[norm]] = p; break; }
                }
            }
        });
        // preserve raw codigo field and raw line for diagnostics and reconstruction
        if (obj.codigo && obj.codigo !== '') obj._rawCodigo = obj.codigo;
        obj._rawLine = line;
        return obj;
    });
    return rows;
}

const BAL2_PRODUCCION_CSV_ROWS = parseProductionCSV(BAL2_PRODUCCION_CSV);

// ===== TABLA DE REFERENCIA DE FACTOR TORSIÓN (COF.E) =====
// Solo usa datos explícitos de la tabla de Producción
let TORSION_REFERENCE_TABLE = [];

// ===== FUNCIÓN PARA EXTRAER NE DEL FORMATO "XX/1" =====
function extractNEFromTitulo(titulo) {
    if (!titulo) return null;
    const match = String(titulo).trim().match(/^(\d+)\/\d+/);
    return match ? parseInt(match[1], 10) : null;
}

// ===== FUNCIÓN PARA CLASIFICAR TIPO DE MATERIAL (PEINADO/CARDADO) =====
function classifyMaterialType(hilado) {
    if (!hilado) return null;
    const upper = String(hilado).toUpperCase();
    if (upper.includes('COP') && !upper.includes('COC')) {
        return 'PEINADO';  // COP = Peinado
    } else if (upper.includes('COC')) {
        return 'CARDADO';   // COC = Cardado
    }
    return null;
}

// ===== FUNCIÓN PARA BUSCAR FACTOR TORSIÓN EN LA TABLA DE REFERENCIA =====
function lookupFactorTorsion(titulo, hilado) {
    const ne = extractNEFromTitulo(titulo);
    const tipo = classifyMaterialType(hilado);
    
    if (!ne || !tipo) {
        console.warn(`lookupFactorTorsion: No se pudo extraer NE (${ne}) o TIPO (${tipo}) de "${titulo}" / "${hilado}"`);
        return null;
    }
    
    // Normalizar el hilado para extraer el material base
    const hilNorm = normalizeHilado(hilado).toUpperCase();
    
    // PASO 1: Buscar en tabla de referencia (tabla vacía inicialmente, se llena manualmente)
    let match = TORSION_REFERENCE_TABLE.find(row => {
        if (row.ne !== ne || row.tipo !== tipo) return false;
        const tabMat = row.material.toUpperCase();
        return tabMat === hilNorm;
    });
    
    if (match) {
        console.log(`lookupFactorTorsion ENCONTRADO (tabla ref): NE=${ne}, TIPO=${tipo}, MATERIAL=${match.material}, COF.E=${match.factorTorsion}`);
        return {
            factorTorsion: match.factorTorsion,
            pinon: match.pinon,
            material: match.material,
            tipo: tipo,
            ne: ne
        };
    }
    
    // PASO 2: Buscar en datos de Producción explícitamente cargados
    const prodRows = getBalance2ProduccionRows() || [];
    const prodMatch = prodRows.find(row => {
        const rowNe = toNumber(row.ne);
        const rowFactor = String(row.factorTorsion || '').trim();
        
        // Solo si tiene Factor Torsion explícito
        if (!rowFactor || rowNe !== ne) return false;
        
        const rowMat = String(row.material || '').trim().toUpperCase();
        const rowMatNorm = normalizeHilado(rowMat).toUpperCase();
        
        // Verificar si el tipo coincide
        let rowTipo = 'PEINADO';
        if (rowMat.includes('COC') || rowMat.includes('CARDADO')) {
            rowTipo = 'CARDADO';
        }
        
        return rowTipo === tipo && rowMatNorm === hilNorm;
    });
    
    if (prodMatch && prodMatch.factorTorsion) {
        const factorVal = toNumber(prodMatch.factorTorsion);
        if (factorVal !== null) {
            console.log(`lookupFactorTorsion ENCONTRADO (producción): NE=${ne}, TIPO=${tipo}, COF.E=${factorVal}`);
            return {
                factorTorsion: factorVal,
                pinon: toNumber(prodMatch.pinon),
                material: String(prodMatch.material || ''),
                tipo: tipo,
                ne: ne
            };
        }
    }
    
    console.warn(`lookupFactorTorsion NO ENCONTRADO: NE=${ne}, TIPO=${tipo}, HILADO="${hilado}"`);
    return null;
}

// ===== FUNCIÓN PÚBLICA PARA CONSULTA RÁPIDA DE FACTOR TORSIÓN =====
window.consultarCOF = function(titulo, hilado) {
    console.log(`\n========== CONSULTA COF.E ==========`);
    console.log(`Título: "${titulo}"`);
    console.log(`Hilado: "${hilado}"`);
    
    const ne = extractNEFromTitulo(titulo);
    const tipo = classifyMaterialType(hilado);
    
    console.log(`NE extraído: ${ne}`);
    console.log(`TIPO clasificado: ${tipo}`);
    
    const result = lookupFactorTorsion(titulo, hilado);
    
    if (result) {
        console.log(`\n✓ RESULTADO ENCONTRADO:`);
        console.log(`  NE: ${result.ne}`);
        console.log(`  TIPO: ${result.tipo}`);
        console.log(`  MATERIAL: ${result.material}`);
        console.log(`  COF.E (FACTOR TORSIÓN): ${result.factorTorsion}`);
        console.log(`  PIÑÓN: ${result.pinon}`);
        console.log(`====================================\n`);
        return result;
    } else {
        console.log(`\n✗ NO SE ENCONTRÓ EN LA TABLA DE REFERENCIA`);
        console.log(`====================================\n`);
        return null;
    }
};

// Ejemplo de uso en consola:
// consultarCOF("40/1", "40/1 COP ORGANICO")
// consultarCOF("30/1", "30/1 COC TANGUIS")
