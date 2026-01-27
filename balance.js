// Control de ventanas para el mÃ³dulo Balance (mod-balance2)
    function switchInnerTabBal2(innerType) {
        // ocultar todas las secciones especÃ­ficas de balance2
        document.querySelectorAll('.nested2-content').forEach(el => {
            el.classList.remove('active');
            el.style.display = 'none';
        });
        // desactivar botones
        const allBtns = document.querySelectorAll('#mod-balance2 .nested-btn');
        allBtns.forEach(b => b.classList.remove('active'));

        // mostrar la solicitada
        const viewId = `inner2-${innerType}-view`;
        const view = document.getElementById(viewId);
        if (view) {
            view.classList.add('active');
            view.style.display = 'block';
        }

        // activar el botÃ³n correspondiente
        const btnMap = { datos: 0, produccion: 1, solver: 2 };
        const idx = btnMap[innerType];
        if (typeof idx === 'number') {
            const btns = document.querySelectorAll('#mod-balance2 .nested-btn');
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
        // Mantener el HTML estÃ¡tico ya definido en index.html si no hay datos externos
        if (el.dataset.locked === 'true') return;
        if (GLOBAL_DATA && GLOBAL_DATA.balance2 && Array.isArray(GLOBAL_DATA.balance2.datos) && GLOBAL_DATA.balance2.datos.length) {
            el.dataset.locked = 'true';
            el.innerHTML = renderSimpleTable(GLOBAL_DATA.balance2.datos);
            if (typeof applyTableFilters === 'function') applyTableFilters(el);
        }
    }

    function renderBalance2Produccion() {
        const el = document.getElementById('bal2-produccion-area');
        if (!el) return;
        if (GLOBAL_DATA && GLOBAL_DATA.balance2 && Array.isArray(GLOBAL_DATA.balance2.produccion) && GLOBAL_DATA.balance2.produccion.length) {
            el.innerHTML = renderSimpleTable(GLOBAL_DATA.balance2.produccion);
        } else {
            el.innerHTML = '<p class="empty-msg">ProducciÃ³n: no hay datos cargados.</p>';
        }
        if (typeof applyTableFilters === 'function') applyTableFilters(el);
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

        // 2) Caso por defecto: generar filas desde Agrup Material
        const materials = getSolverMaterialList();
        if (!materials.length) return [];

        const rows = materials.map(name => {
            const neVal = getNeForMaterial(name);
            const info = getSolverMaterialInfo(name);
            const prod = getProductionRowForMaterial(name);

            const tpp = getProductionValue(prod, ['torsion', 'tpp', 't.p.p']);
            const rpm = getProductionValue(prod, ['velocidad', 'rpm', 'r.p.m']);
            const porMin = getProductionValue(prod, ['metros por minuto', 'm/min', 'm por minuto', 'por minuto']);
            const hora100 = getProductionValue(prod, ['prd hora 100', 'prod hora 100', 'hora 100', 'hora100']);

            const hora100Num = toNumber(hora100);
            const eficPct = 78;
            const eficFactor = 0.78;
            const horaEfectiva = (hora100Num !== null) ? (hora100Num * eficFactor) : null;
            const diariaKg = (horaEfectiva !== null) ? (horaEfectiva * 24) : null;

            const kgSol = info && info.kgSol !== null ? info.kgSol : null;
            const kgReq = info && info.kgReq !== null ? info.kgReq : null;
            const diariaReq = (kgReq !== null) ? (kgReq / 24) : null;
            const horasMaquina = (diariaReq !== null && horaEfectiva) ? (diariaReq / horaEfectiva) : null;
            const maquinas = (diariaReq !== null && diariaKg) ? (diariaReq / diariaKg) : null;
            const prod24 = (diariaReq !== null) ? (diariaReq * 24) : null;

            return {
                maquina: '',
                ne: neVal !== null && !isNaN(neVal) ? Math.round(neVal) : '',
                material: name,
                kgSol: kgSol,
                cof: '',
                e: '',
                tpp: tpp || '',
                cd: '',
                rpm: rpm || '',
                husos: '',
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

    function getSolverMaterialInfo(materialName) {
        try {
            if (typeof getGroupsFromData !== 'function') return null;
            const crudoAll = GLOBAL_DATA && GLOBAL_DATA.nuevo ? GLOBAL_DATA.nuevo : [];
            const htrAll = GLOBAL_DATA && GLOBAL_DATA.htr ? GLOBAL_DATA.htr : [];
            const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
            const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;

            const groupedCrudo = getGroupsFromData(activeCrudo);
            const groupedHtr = getGroupsFromData(activeHtr);

            let group = groupedCrudo.find(g => (g.name || '').toString() === (materialName || '').toString());
            let isHtr = false;
            if (!group) {
                group = groupedHtr.find(g => (g.name || '').toString() === (materialName || '').toString());
                isHtr = true;
            }
            if (!group) return null;
            const kgSol = (group.items || []).reduce((s, it) => s + (Number(it.kg || 0)), 0);
            const factor = isHtr ? 0.60 : 0.65;
            const kgReq = kgSol ? (kgSol / factor) : null;
            return { kgSol, kgReq, isHtr };
        } catch (e) {
            console.error('getSolverMaterialInfo error', e);
            return null;
        }
    }

    function getProductionRowForMaterial(materialName) {
        try {
            const prod = GLOBAL_DATA && GLOBAL_DATA.balance2 && Array.isArray(GLOBAL_DATA.balance2.produccion)
                ? GLOBAL_DATA.balance2.produccion
                : [];
            if (!prod.length) return null;
            const sample = prod[0] || {};
            const keys = Object.keys(sample);
            const matKey = keys.find(k => k.toString().toLowerCase().includes('material')) || keys.find(k => k.toString().toLowerCase().includes('hilado'));
            if (!matKey) return null;
            const targetNorm = normalizeSolverKey(materialName);
            return prod.find(r => normalizeSolverKey(r[matKey]) === targetNorm) || null;
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
        const match = str.match(/-?\d+(\.\d+)?/);
        if (!match) return null;
        const num = Number(match[0]);
        return isFinite(num) ? num : null;
    }

    function getSolverMaterialList() {
        try {
            if (typeof getGroupsFromData !== 'function') return [];
            const crudoAll = GLOBAL_DATA && GLOBAL_DATA.nuevo ? GLOBAL_DATA.nuevo : [];
            const htrAll = GLOBAL_DATA && GLOBAL_DATA.htr ? GLOBAL_DATA.htr : [];

            const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
            const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;

            const groupedCrudo = getGroupsFromData(activeCrudo);
            const pureGroups = groupedCrudo.filter(g => !g.isMezcla && (g.name || '').toString().toUpperCase() !== 'OTROS');
            const mixGroups = groupedCrudo.filter(g => g.isMezcla && (g.name || '').toString().toUpperCase() !== 'OTROS');

            const groupedHtr = getGroupsFromData(activeHtr);
            const htrGroups = groupedHtr.filter(g => (g.name || '').toString().toUpperCase() !== 'OTROS');

            const ordered = pureGroups.concat(mixGroups).concat(htrGroups);
            const names = [];
            const seen = new Set();
            ordered.forEach(g => {
                const n = (g && g.name ? g.name : '').toString().trim();
                if (!n) return;
                if (!seen.has(n)) { seen.add(n); names.push(n); }
            });

            return names;
        } catch (e) {
            console.error('getSolverMaterialList error', e);
            return [];
        }
    }

    function getNeForMaterial(materialName) {
        try {
            if (typeof computeWeightedNe !== 'function' || typeof getGroupsFromData !== 'function') return null;
            const crudoAll = GLOBAL_DATA && GLOBAL_DATA.nuevo ? GLOBAL_DATA.nuevo : [];
            const htrAll = GLOBAL_DATA && GLOBAL_DATA.htr ? GLOBAL_DATA.htr : [];
            const activeCrudo = (crudoAll.some(r => r.highlight)) ? crudoAll.filter(r => r.highlight) : crudoAll;
            const activeHtr = (htrAll.some(r => r.highlight)) ? htrAll.filter(r => r.highlight) : htrAll;
            const combined = activeCrudo.concat(activeHtr);
            const groups = getGroupsFromData(combined);
            const group = groups.find(g => (g.name || '').toString() === (materialName || '').toString());
            if (!group || !group.items) return null;
            return computeWeightedNe(group.items || []);
        } catch (e) {
            console.error('getNeForMaterial error', e);
            return null;
        }
    }

    function renderSolverTable(rows) {
        const cols = [
            { key: 'maquina', label: 'MAQUINA' },
            { key: 'ne', label: 'Ne' },
            { key: 'material', label: 'MATERIAL' },
            { key: 'kgSol', label: 'KG SOL' },
            { key: 'cof', label: 'COF.' },
            { key: 'e', label: 'e' },
            { key: 'tpp', label: 'T. P. P.' },
            { key: 'cd', label: 'C.D' },
            { key: 'rpm', label: 'R.P.M.' },
            { key: 'husos', label: 'HUSOS' },
            { key: 'porMin', label: 'POR MINUTO' },
            { key: 'hora100', label: 'HORA 100% KG.' },
            { key: 'efic', label: 'EFIC.' },
            { key: 'horaEfectiva', label: 'HORA Efectiva' },
            { key: 'diariaKg', label: 'DIARIA Kg.' },
            { key: 'diariaReq', label: 'DIARIA REQUERIDA' },
            { key: 'horasMaquina', label: 'HORAS MAQUINA' },
            { key: 'maquinas', label: 'MAQUINAS' },
            { key: 'prod24', label: 'PROD. CON 24 DIAS (kG.)' },
            { key: 'obs', label: 'OBSERVACIONES' }
        ];

        let html = '<div class="table-wrap"><table><thead><tr>';
        cols.forEach(c => {
            html += `<th class=\"th-base\">${c.label}</th>`;
        });
        html += '</tr></thead><tbody>';
        html += `<tr class="bal-subtotal"><td style="text-align:left;">CONTINUAS</td>` + `<td></td>`.repeat(cols.length - 1) + `</tr>`;
        rows.forEach(r => {
            html += '<tr>';
            cols.forEach(c => {
                const val = (r && r[c.key] !== undefined && r[c.key] !== null) ? r[c.key] : '';
                html += `<td>${formatSolverValue(val)}</td>`;
            });
            html += '</tr>';
        });
        // Ne Promedio
        const neVals = rows.map(r => toNumber(r.ne)).filter(v => v !== null);
        const neProm = neVals.length ? (neVals.reduce((a,b)=>a+b,0) / neVals.length) : null;
        html += `<tr class="bal-subtotal"><td></td><td>${formatSolverValue(neProm)}</td><td>NE PROMEDIO</td>` + `<td></td>`.repeat(cols.length - 3) + `</tr>`;

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
            return String(Math.round(val));
        }
        if (typeof val === 'string') {
            const trimmed = val.trim();
            if (trimmed && /^[\d,.\-]+$/.test(trimmed)) {
                const num = toNumber(trimmed);
                if (num !== null) {
                    if (typeof fmt === 'function') return fmt(num);
                    return String(Math.round(num));
                }
            }
        }
        return String(val);
    }

    // (Opciones de ediciÃ³n/historial deshabilitadas por solicitud)
