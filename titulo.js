// MODULO BALANCE TITULO
    function renderTituloModule() {
        const activeView = document.querySelector('#mod-titulo .bal-tab-content:not([style*="display: none"])');
        const isCrudo = activeView && activeView.id === 'tit-crudo-view';

        // Usar los arrays programmados si existen highlights (PCP Fuente), si no usar todos
        let activeCrudo = GLOBAL_DATA.nuevo;
        if (GLOBAL_DATA.nuevo && GLOBAL_DATA.nuevo.some(r => r.highlight)) activeCrudo = GLOBAL_DATA.nuevo.filter(r => r.highlight);
        let activeHtr = GLOBAL_DATA.htr;
        if (GLOBAL_DATA.htr && GLOBAL_DATA.htr.some(r => r.highlight)) activeHtr = GLOBAL_DATA.htr.filter(r => r.highlight);

        if (isCrudo) {
            document.getElementById('tit-table-crudo').innerHTML = generateTituloTable(activeCrudo, 'crudo') + renderOtrosBlock(false, activeCrudo.filter(r => !r.isMezcla));
        } else {
            document.getElementById('tit-table-htr').innerHTML = generateTituloTable(activeHtr, 'htr') + renderOtrosBlock(true, activeHtr);
        }
        updateTituloFooter();
        if (typeof applyTableFilters === 'function') applyTableFilters();
    }

    function generateTituloTable(rows, tipo) {
        if (!rows || rows.length === 0) return '<p class="empty-msg">Sin datos.</p>';
        const grupos = {};
        // Helper: detectar tÃ­tulos especiales a partir del campo `hilado` y devolver clave de grupo.
        function getTituloGroupKey(r) {
            const hil = (r.hilado || '').toString().toUpperCase();
            // 40/1 VI (o "40/1 VI COP ...") -> grupo '40/1 VI'
            let m = hil.match(/\b(40(?:\/1)?)\b[^\n\r]*\bVI\b/i);
            if (m) return `${m[1]} VI`;
            // 50 IV (o "50/1 IV ...") -> grupo '50/1 IV' o '50 IV' segÃºn captura
            m = hil.match(/\b(50(?:\/1)?)\b[^\n\r]*\bIV\b/i);
            if (m) return `${m[1]} IV`;
            // Por defecto usar la normalizaciÃ³n existente (basada en `titulo`)
            const rawTitulo = (r.titulo || r.hilado || 'SIN TITULO');
            return normalizeTitulo(rawTitulo);
        }

        rows.forEach(r => {
            const rawTitulo = (r.titulo || 'SIN TITULO');
            const titulo = getTituloGroupKey(r);
            if (!grupos[titulo]) { grupos[titulo] = { items: [], totalKg: 0, rawNames: new Set() }; }
            grupos[titulo].items.push(r);
            grupos[titulo].rawNames.add(rawTitulo);
            grupos[titulo].totalKg += r.kg;
        });
        // Ordenar por nÃºmero principal (ej: 36, 40, 44) y dejar 'SIN TITULO' al final
        const extractLeadingNum = t => { const mm = String(t).match(/(\d{1,3})/); return mm ? parseInt(mm[1],10) : 9999; };
        const titulosOrdenados = Object.keys(grupos).sort((a, b) => {
            if (a === "SIN TITULO") return 1; if (b === "SIN TITULO") return -1;
            return extractLeadingNum(a) - extractLeadingNum(b);
        });
        let html = '';
        // AÃ±adir placeholders de bloques vacÃ­os para TÃ­tulo (si existen)
        try {
            const placeholders = GLOBAL_DATA.emptyGroups && GLOBAL_DATA.emptyGroups.crudo ? GLOBAL_DATA.emptyGroups.crudo : [];
            placeholders.forEach(p => {
                if (!grupos[p.name]) {
                    grupos[p.name] = { items: [], totalKg: 0 };
                }
            });
        } catch(e) {}

        titulosOrdenados.forEach(titulo => {
            const grupo = grupos[titulo];
            const safeTitle = (titulo || '').toString().replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            // Mostrar el tÃ­tulo normalizado en el encabezado; si hay nombres originales diferentes,
            // mostrar aviso de inclusiÃ³n (ej: '36/1 (incluye: 40/1 VI)')
            let headerLabel = titulo;
            try {
                const raws = Array.from(grupo.rawNames || []);
                const uniqueRaw = raws.filter(rn => normalizeTitulo(rn) !== titulo);
                // Omitir el texto "(incluye: ...)" para tÃ­tulos especiales con sufijos romanos (ej. 'VI','IV')
                const isSpecialSuffix = /\b(IV|VI|V|II|III|I)\b/i.test(String(titulo));
                if (uniqueRaw.length > 0 && !isSpecialSuffix) {
                    headerLabel = `${titulo} (incluye: ${uniqueRaw.join(', ')})`;
                }
            } catch(e) {}
            // Permitir soltar sobre el bloque de tÃ­tulo para mover filas a este TÃTULO (cambia `titulo` del item)
            const isMezGroup = (grupo.items || []).some(it => !!it.isMezcla);
            html += `<div class="table-wrap" ondragover="allowDrop(event)" ondragenter="this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="handleDropToTitle(event, '${safeTitle}', ${isMezGroup ? 'true' : 'false'}, ${tipo === 'htr' ? 'true' : 'false'})"><div class="bal-header-row"><span class="bal-title">TITULO: ${headerLabel}</span>`;
            if (!grupo.items || grupo.items.length === 0) {
                html += `<button style="margin-left:auto; padding:6px 12px; font-size:11px; background:#ef4444; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="confirmDeleteGroup('${safeTitle}', '${tipo === 'crudo' ? 'pure' : 'htr'}')">Eliminar bloque vacÃ­o</button>`;
            }
            html += `</div><table><thead><tr><th class="th-base">ORDEN</th><th class="th-base">CLIENTE</th><th class="th-base">TEMPORADA</th><th class="th-base">RSV</th><th class="th-base">OP</th><th class="th-base">HILADO</th><th class="th-base">NE</th><th class="th-base">TIPO</th><th class="th-base">KG SOLICITADOS</th></tr></thead><tbody>`;
            if (!grupo.items || grupo.items.length === 0) {
                html += `<tr ondragover="allowDrop(event)" ondragenter="this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="handleDropToTitle(event, '${safeTitle}', ${isMezGroup ? 'true' : 'false'}, ${tipo === 'htr' ? 'true' : 'false'})"><td colspan="9" style="text-align:center; padding:30px; color:#999;"><em>Este bloque estÃ¡ vacÃ­o</em></td></tr>`;
            }
            grupo.items.forEach(r => {
                const dragAttrs = `draggable="true" ondragstart="handleDragStart(event, '${r._id}')" ondrop="handleDrop(event, '${r._id}')" ondragover="allowDrop(event)"`;
                const neVal = getNeFromItem(r);
                const neDisplay = (neVal !== null && !isNaN(neVal)) ? Math.round(neVal) : '-';
                html += `<tr ${dragAttrs}><td>${r.orden || '-'}</td><td>${r.cliente || '-'}</td><td>${r.temporada || '-'}</td><td>${r.rsv || ''}</td><td>${r.op || ''}</td><td class="hilado-cell">${r.hilado || '-'}</td><td style="font-weight:600;">${neDisplay}</td><td>${r.tipo || ''}</td><td style="font-weight:600;">${fmtDecimal(r.kg)}</td></tr>`;
            });
            // Cambiar etiqueta inferior de SUBTOTAL a TOTAL para mantener consistencia con Balance Material
            html += `<tr class="bal-subtotal"><td colspan="8" style="text-align:right;">TOTAL ${titulo}:</td><td>${fmt(grupo.totalKg)}</td></tr></tbody></table></div>`;
        });
        return html;
    }

    function updateTituloFooter() {
        // Calcular totales basados en la vista programada (highlights) para que coincida con PCP Fuente
        const activeCrudo = (GLOBAL_DATA.nuevo && GLOBAL_DATA.nuevo.some(r => r.highlight)) ? GLOBAL_DATA.nuevo.filter(r => r.highlight) : GLOBAL_DATA.nuevo;
        const activeHtr = (GLOBAL_DATA.htr && GLOBAL_DATA.htr.some(r => r.highlight)) ? GLOBAL_DATA.htr.filter(r => r.highlight) : GLOBAL_DATA.htr;
        const totalCrudo = round0((activeCrudo || []).reduce((sum, r) => sum + r.kg, 0));
        const totalHTR = round0((activeHtr || []).reduce((sum, r) => sum + r.kg, 0));
        const totalGlobal = totalCrudo + totalHTR;
        const container = document.getElementById('tit-footer-grid');
        // Calcular NE para la vista actual (ponderado por KG solicitados)
        const neCrudoVal = computeWeightedNe(activeCrudo || []);
        const neHtrVal = computeWeightedNe(activeHtr || []);
        const neTotalVal = computeWeightedNe(((activeCrudo || []).concat(activeHtr || [])));
        const neCrudoDisplay = (neCrudoVal !== null && !isNaN(neCrudoVal)) ? fmt(Math.round(neCrudoVal)) : '-';
        const neHtrDisplay = (neHtrVal !== null && !isNaN(neHtrVal)) ? fmt(Math.round(neHtrVal)) : '-';
        const neTotalDisplay = (neTotalVal !== null && !isNaN(neTotalVal)) ? fmt(Math.round(neTotalVal)) : '-';

        container.innerHTML = `<div class="footer-item"><label>TOTAL KG CRUDO</label><div class="val val-crudo">${fmt(totalCrudo)}</div></div><div class="footer-item"><label>TOTAL KG HTR</label><div class="val val-mezcla">${fmt(totalHTR)}</div></div><div class="footer-item" style="border-left: 1px solid rgba(255,255,255,0.2); padding-left: 20px;"><label>TOTAL KG (CRUDO + HTR)</label><div class="val val-total">${fmt(totalGlobal)}</div></div>
            <div style="flex-basis:100%; height:8px;"></div>
            <div class="footer-item"><label>Ne CRUDO</label><div class="val">${neCrudoDisplay}</div></div>
            <div class="footer-item"><label>Ne HTR</label><div class="val">${neHtrDisplay}</div></div>
            <div class="footer-item"><label>Ne PROMEDIO</label><div class="val">${neTotalDisplay}</div></div>`;
    }
