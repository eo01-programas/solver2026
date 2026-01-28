// MODULO BALANCE TITULO
    function renderTituloModule() {
        // Usar los arrays programmados si existen highlights (PCP Fuente), si no usar todos
        let activeCrudo = GLOBAL_DATA.nuevo;
        if (GLOBAL_DATA.nuevo && GLOBAL_DATA.nuevo.some(r => r.highlight)) activeCrudo = GLOBAL_DATA.nuevo.filter(r => r.highlight);
        
        // Separar CRUDO y MEZCLA
        const actCrudoPuro = activeCrudo.filter(r => !r.isMezcla);
        const actCrudoMezcla = activeCrudo.filter(r => r.isMezcla);
        
        let activeHtr = GLOBAL_DATA.htr;
        if (GLOBAL_DATA.htr && GLOBAL_DATA.htr.some(r => r.highlight)) activeHtr = GLOBAL_DATA.htr.filter(r => r.highlight);

        // Debug: verificar qué datos se están usando
        console.log('renderTituloModule - CRUDO items:', actCrudoPuro.length, actCrudoPuro);
        console.log('renderTituloModule - MEZCLA items:', actCrudoMezcla.length, actCrudoMezcla);
        console.log('renderTituloModule - HTR items:', activeHtr.length, activeHtr);

        // Renderizar las 3 vistas
        document.getElementById('tit-table-crudo').innerHTML = generateTituloTable(actCrudoPuro, 'crudo') + renderOtrosBlock(false, actCrudoPuro);
        document.getElementById('tit-table-mezcla').innerHTML = generateTituloTable(actCrudoMezcla, 'mezcla') + renderOtrosBlock(false, actCrudoMezcla);

        // Para HTR: tomar exactamente los grupos usados en Agrup Material (excluyendo OTROS)
        // y luego agrupar esos items por TÍTULO para mostrarlos en el módulo Título.
        let groupedHtr = [];
        try {
            groupedHtr = getGroupsFromData(activeHtr || []);
        } catch (e) { groupedHtr = []; }
        const groupedHtrForTable = (groupedHtr || []).filter(g => ((g.name||'').toString().toUpperCase() !== 'OTROS'));
        // Aplanar items para pasarlos a la vista de TÍTULO (agrupamiento por titulo lo hace generateTituloTable)
        const htrItemsForTitle = groupedHtrForTable.flatMap(g => g.items || []);
        document.getElementById('tit-table-htr').innerHTML = generateTituloTable(htrItemsForTitle, 'htr') + renderOtrosBlock(true, (activeHtr || []).filter(r => (r.group||'').toString().toUpperCase() === 'OTROS'));
        
        updateTituloFooter();
        if (typeof applyTableFilters === 'function') applyTableFilters();
    }

    function generateTituloTable(rows, tipo) {
        if (!rows || rows.length === 0) return '<p class="empty-msg">Sin datos.</p>';
        const grupos = {};
        // Helper: detectar tAtulos especiales a partir del campo `hilado` y devolver clave de grupo.
        function getTituloGroupKey(r) {
            const hil = (r.hilado || '').toString().toUpperCase();
            // 40/1 VI (o "40/1 VI COP ...") -> grupo '40/1 VI'
            let m = hil.match(/\b(40\/1)\b[^\n\r]*\bVI\b/i);
            if (m) return `${m[1]} VI`;
            // 50 IV (o "50/1 IV ...") -> grupo '50/1 IV' o '50 IV' segAn captura
            m = hil.match(/\b(50\/1)\b[^\n\r]*\bIV\b/i);
            if (m) return `${m[1]} IV`;
            // Por defecto usar la normalizaciAn existente (basada en `titulo`)
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
        // Ordenar por nAmero principal (ej: 36, 40, 44) y dejar 'SIN TITULO' al final
        const extractLeadingNum = t => { const mm = String(t).match(/(\d{1,3})/); return mm ? parseInt(mm[1],10) : 9999; };
        const titulosOrdenados = Object.keys(grupos).sort((a, b) => {
            if (a === "SIN TITULO") return 1; if (b === "SIN TITULO") return -1;
            return extractLeadingNum(a) - extractLeadingNum(b);
        });
        let html = '';
        // AAadir placeholders de bloques vacAos para TAtulo (si existen)
        try {
            const placeholders = (GLOBAL_DATA.emptyGroups && GLOBAL_DATA.emptyGroups.crudo) ? GLOBAL_DATA.emptyGroups.crudo : [];
            placeholders.forEach(p => {
                if (!grupos[p.name]) {
                    grupos[p.name] = { items: [], totalKg: 0 };
                }
            });
        } catch(e) {}

        titulosOrdenados.forEach(titulo => {
            const grupo = grupos[titulo];
            const safeTitle = (titulo || '').toString().replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            // Mostrar el tAtulo normalizado en el encabezado; si hay nombres originales diferentes,
            // mostrar aviso de inclusiAn (ej: '36/1 (incluye: 40/1 VI)')
            let headerLabel = titulo;
            try {
                const raws = Array.from(grupo.rawNames || []);
                const uniqueRaw = raws.filter(rn => normalizeTitulo(rn) !== titulo);
                // Omitir el texto "(incluye: ...)" para tAtulos especiales con sufijos romanos (ej. 'VI','IV')
                const isSpecialSuffix = /\b(IV|VI|V|II|III|I)\b/i.test(String(titulo));
                if (uniqueRaw.length > 0 && !isSpecialSuffix) {
                    headerLabel = `${titulo} (incluye: ${uniqueRaw.join(', ')})`;
                }
            } catch(e) {}
            // Permitir soltar sobre el bloque de tAtulo para mover filas a este TATULO (cambia `titulo` del item)
            const isMezGroup = (grupo.items || []).some(it => !!it.isMezcla);
            html += `<div class="table-wrap" ondragover="allowDrop(event)" ondragenter="this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="handleDropToTitle(event, '${safeTitle}', ${isMezGroup ? 'true' : 'false'}, ${tipo === 'htr' ? 'true' : 'false'})"><div class="bal-header-row"><span class="bal-title">TITULO: ${headerLabel}</span>`;
            if (!grupo.items || grupo.items.length === 0) {
                html += `<button style="margin-left:auto; padding:6px 12px; font-size:11px; background:#ef4444; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="confirmDeleteGroup('${safeTitle}', '${tipo === 'crudo' ? 'pure' : 'htr'}')">Eliminar bloque vacAo</button>`;
            }
            html += `</div><table><thead><tr><th class="th-base">ORDEN</th><th class="th-base">CLIENTE</th><th class="th-base">TEMPORADA</th><th class="th-base">RSV</th><th class="th-base">OP</th><th class="th-base">HILADO</th><th class="th-base">NE</th><th class="th-base">TIPO</th><th class="th-base">KG SOLICITADOS</th></tr></thead><tbody>`;
            if (!grupo.items || grupo.items.length === 0) {
                html += `<tr ondragover="allowDrop(event)" ondragenter="this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="handleDropToTitle(event, '${safeTitle}', ${isMezGroup ? 'true' : 'false'}, ${tipo === 'htr' ? 'true' : 'false'})"><td colspan="9" style="text-align:center; padding:30px; color:#999;"><em>Este bloque estA vacAo</em></td></tr>`;
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
        const actCrudoPuro = activeCrudo.filter(r => !r.isMezcla);
        const actCrudoMezcla = activeCrudo.filter(r => r.isMezcla);
        const activeHtr = (GLOBAL_DATA.htr && GLOBAL_DATA.htr.some(r => r.highlight)) ? GLOBAL_DATA.htr.filter(r => r.highlight) : GLOBAL_DATA.htr;
        
        const totalCrudoPuro = round0((actCrudoPuro || []).reduce((sum, r) => sum + r.kg, 0));
        const totalMezcla = round0((actCrudoMezcla || []).reduce((sum, r) => sum + r.kg, 0));
        const totalCrudo = totalCrudoPuro + totalMezcla;
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

        container.innerHTML = `<div class="footer-item"><label>TOTAL KG CRUDO PURO</label><div class="val val-crudo">${fmt(totalCrudoPuro)}</div></div><div class="footer-item"><label>TOTAL KG MEZCLA</label><div class="val val-mezcla">${fmt(totalMezcla)}</div></div><div class="footer-item"><label>TOTAL KG CRUDO (PURO+MIX)</label><div class="val" style="color:#2563eb;">${fmt(totalCrudo)}</div></div><div class="footer-item"><label>TOTAL KG HTR</label><div class="val" style="color:#dc2626;">${fmt(totalHTR)}</div></div><div class="footer-item" style="border-left: 1px solid rgba(255,255,255,0.2); padding-left: 20px;"><label>TOTAL KG (CRUDO + HTR)</label><div class="val val-total">${fmt(totalGlobal)}</div></div>
            <div style="flex-basis:100%; height:8px;"></div>
            <div class="footer-item"><label>Ne CRUDO</label><div class="val">${neCrudoDisplay}</div></div>
            <div class="footer-item"><label>Ne HTR</label><div class="val">${neHtrDisplay}</div></div>
            <div class="footer-item"><label>Ne PROMEDIO</label><div class="val">${neTotalDisplay}</div></div>`;
    }
