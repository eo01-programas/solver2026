function renderBalanceModule() {
        // En balance: Usar highlights si existen (significa que hubo correcciÃ³n), si no, usar todos.
        // Pero OJO: Para arrastrar y soltar necesitamos el array real. 
        // Si filtramos aquÃ­, perdemos la referencia al array principal.
        
        // Mejor enfoque: Siempre trabajar con el array principal filtrado logicamente, 
        // pero para Drag & Drop necesitamos saber el indice en el array PRINCIPAL.
        
        let activeCrudo = GLOBAL_DATA.nuevo;
        if (GLOBAL_DATA.nuevo.some(r => r.highlight)) {
            activeCrudo = GLOBAL_DATA.nuevo.filter(r => r.highlight);
        }

        let activeHtr = GLOBAL_DATA.htr;
        if (GLOBAL_DATA.htr.some(r => r.highlight)) {
            activeHtr = GLOBAL_DATA.htr.filter(r => r.highlight);
        }

        const groupedCrudo = getGroupsFromData(activeCrudo);
        // Insertar placeholders de bloques vacÃ­os (crudo) si el usuario los preservÃ³
        if (GLOBAL_DATA.emptyGroups && Array.isArray(GLOBAL_DATA.emptyGroups.crudo)) {
            GLOBAL_DATA.emptyGroups.crudo.forEach(p => {
                if (!groupedCrudo.some(g => (g.name || '').toString() === (p.name || ''))) {
                    groupedCrudo.push({ name: p.name, items: [], totalKg: 0, isMezcla: !!p.isMezcla });
                }
            });
        }
        // Excluir grupos 'OTROS' de la lista principal para evitar duplicados al renderizar
        const pureGroups = groupedCrudo.filter(g => !g.isMezcla && ((g.name||'').toString().toUpperCase() !== 'OTROS'));
        const mixGroups = groupedCrudo.filter(g => g.isMezcla && ((g.name||'').toString().toUpperCase() !== 'OTROS'));

        // Renderizar OTROS usando el array completo para que no desaparezcan al filtrar por highlights
        document.getElementById('bal-table-pure').innerHTML = generateBalanceGroupTable(pureGroups, 'pure', false) + renderOtrosBlock(false, GLOBAL_DATA.nuevo.filter(r => !r.isMezcla));
        document.getElementById('bal-table-mix').innerHTML = generateBalanceGroupTable(mixGroups, 'mix', false) + renderOtrosBlock(false, activeCrudo.filter(r => r.isMezcla));
        document.getElementById('bal-summary-area').innerHTML = generateSummaryTable(groupedCrudo, false);

        const groupedHtr = getGroupsFromData(activeHtr);
        // Insertar placeholders de bloques vacÃ­os (htr) si existen
        if (GLOBAL_DATA.emptyGroups && Array.isArray(GLOBAL_DATA.emptyGroups.htr)) {
            GLOBAL_DATA.emptyGroups.htr.forEach(p => {
                if (!groupedHtr.some(g => (g.name || '').toString() === (p.name || ''))) {
                    groupedHtr.push({ name: p.name, items: [], totalKg: 0, isMezcla: !!p.isMezcla });
                }
            });
        }
        const groupedHtrForTable = groupedHtr.filter(g => ((g.name||'').toString().toUpperCase() !== 'OTROS'));
        // Para HTR, renderizar OTROS desde el array completo HTR
        document.getElementById('bal-table-htr').innerHTML = generateBalanceGroupTable(groupedHtrForTable, 'htr', true) + renderOtrosBlock(true, GLOBAL_DATA.htr);
        // El resumen HTR debe considerar solo los grupos HTR (excluir OTROS)
        document.getElementById('bal-htr-summary-area').innerHTML = generateSummaryTable(groupedHtrForTable, true);

        updateFooterTotals();
        if (typeof applyTableFilters === 'function') applyTableFilters();
    }

    function getCanonicalGroupName(rawName) {
        const original = rawName;
        let name = rawName.toUpperCase();
        
        // Extraer porcentajes primero (con o sin parÃ©ntesis)
        let pctStr = "";
        const pctMatch1 = name.match(/\((\d{1,3}(?:\/\d{1,3})+)\s*%?\)/);
        const pctMatch2 = name.match(/(\d{1,3}(?:\/\d{1,3})+)\s*%/);
        if (pctMatch1) {
            pctStr = pctMatch1[1];
        } else if (pctMatch2) {
            pctStr = pctMatch2[1];
        }
        
        // Remover todos los porcentajes del nombre
        name = name.replace(/\(?\d{1,3}(?:\/\d{1,3})+\s*%?\)?/g, "");
        
        // Normalizar nombres de componentes
        name = name.replace(/TENCEL/g, "LYOCELL");
        name = name.replace(/\bCOP\b/g, "");
        name = name.replace(/\bNC\b/g, ""); 
        name = name.replace(/\bSTD\b/g, "");
        name = name.replace(/\s+/g, " ").trim();
        name = name.replace(/\s*\/\s*/g, "/");

        // Si quedaron dos materiales juntos separados solo por espacio (ej: "PIMA LYOCELL"),
        // conviÃ©rtelos a formato con "/" para unificar el grupo ("PIMA/LYOCELL").
        name = name.replace(/\b(PIMA|LYOCELL|MODAL|TANGUIS)\s+(PIMA|LYOCELL|MODAL|TANGUIS)\b/g, '$1/$2');
        
        // Agregar porcentajes al final si existen
        if (pctStr) {
            name = name + " (" + pctStr + "%)";
        }
        
        console.log(`getCanonicalGroupName: "${original}" -> "${name}"`);
        return name;
    }

    function getGroupsFromData(rows) {
        const groupsMap = new Map();
        
        rows.forEach(r => {
            let groupKey = r.group;
            let isMezcla = r.isMezcla;
            
            if(!groupsMap.has(groupKey)) {
                groupsMap.set(groupKey, {
                    name: groupKey,
                    items: [],
                    totalKg: 0,
                    isMezcla: isMezcla
                });
            }
            let g = groupsMap.get(groupKey);
            g.items.push(r);
            g.totalKg += r.kg;
        });
        
        return Array.from(groupsMap.values());
    }

    function parseMixComposition(hiladoStr, kgSolicitado) {
        // Detectar porcentajes: acepta "(75/25%)" o "75/25%"
        let pctStr = '';
        const pctMatch1 = hiladoStr.match(/\((\d{1,3}(?:\/\d{1,3})+)\s*%?\)/);
        const pctMatch2 = hiladoStr.match(/(\d{1,3}(?:\/\d{1,3})+)\s*%/);
        if (pctMatch1) {
            pctStr = pctMatch1[1];
        } else if (pctMatch2) {
            pctStr = pctMatch2[1];
        }

        if (!pctStr) {
            // No hay porcentajes explÃ­citos: determinar si el hilado es COP/ORGANICO
            let cleanName = hiladoStr.toUpperCase().replace(/^\d+\/\d+\s+/, "").trim();
            const isPriority = /\bCOP\b/i.test(cleanName) || /(?:ORGANICO|ORG|ORGANIC)/i.test(cleanName);
            const factor = isPriority ? 0.65 : 0.85;
            let k = round0(kgSolicitado / factor);
            const isAlgodon = isAlgodonText(cleanName || "CRUDO");
            return [{ name: cleanName || "CRUDO", kg: k, qq: isAlgodon ? round0(k / 46) : null, kgSol: round0(kgSolicitado) }];
        }

        // Extraer porcentajes y normalizar a decimales (ej: "40/30/30" -> [0.40, 0.30, 0.30])
        const pcts = pctStr.split('/').map(n => parseFloat(n) / 100);

        // Remover el bloque de porcentajes y el tÃ­tulo inicial (ej: "40/1")
        let cleanName = hiladoStr.replace(/\(?\d{1,3}(?:\/\d{1,3})+\s*%?\)?/g, "").trim();
        cleanName = cleanName.replace(/^\d+\/\d+\s+/, "").trim();

        // Parsear nombres: si contiene "/", dividir por eso. Sino, identificar materiales y agrupar
        let names = [];
        if (cleanName.includes('/')) {
            names = cleanName.split('/').map(s => s.trim()).filter(s => s !== "" && s.length > 0);
        } else {
            const materials = ['PIMA', 'LYOCELL', 'TENCEL', 'MODAL', 'TANGUIS', 'ORGANICO', 'ORGANIC', 'BCI', 'USTCP'];
            const modifiers = ['NC', 'STD', 'ORGANICO', 'ORGANIC', 'OCS'];
            let remaining = cleanName;
            while (remaining.length > 0) {
                remaining = remaining.trim();
                if (!remaining) break;
                let matched = false;
                let currentName = '';
                if (remaining.toUpperCase().startsWith('COP ')) {
                    currentName = 'COP ';
                    remaining = remaining.substring(4).trim();
                }
                for (let mat of materials) {
                    if (remaining.toUpperCase().startsWith(mat + ' ') || remaining.toUpperCase() === mat) {
                        currentName += mat;
                        remaining = remaining.substring(mat.length).trim();
                        matched = true;
                        while (remaining.length > 0) {
                            let foundMod = false;
                            for (let mod of modifiers) {
                                if (remaining.toUpperCase().startsWith(mod + ' ') || remaining.toUpperCase() === mod) {
                                    currentName += ' ' + mod;
                                    remaining = remaining.substring(mod.length).trim();
                                    foundMod = true;
                                    break;
                                }
                            }
                            if (!foundMod) break;
                        }
                        break;
                    }
                }
                if (matched && currentName.trim()) names.push(currentName.trim());
                else if (remaining.length > 0) {
                    let nextMaterialIdx = remaining.length;
                    for (let mat of materials) {
                        const idx = remaining.toUpperCase().indexOf(mat);
                        if (idx >= 0 && idx < nextMaterialIdx) nextMaterialIdx = idx;
                    }
                    if (nextMaterialIdx > 0) {
                        names.push(remaining.substring(0, nextMaterialIdx).trim());
                        remaining = remaining.substring(nextMaterialIdx);
                    } else {
                        if (remaining.trim()) names.push(remaining.trim());
                        break;
                    }
                }
            }
        }

        while (names.length < pcts.length) names.push(`COMP ${names.length + 1}`);

        let components = [];
        for (let i = 0; i < pcts.length; i++) {
            components.push({ name: names[i], pct: pcts[i], baseToken: getBaseMaterialToken(names[i]) });
        }

        // Buscar el primer componente que sea 'COP' o contenga 'ORGANICO' y moverlo al inicio
        const firstPriorityIdx = components.findIndex(c => /\bCOP\b/i.test(c.name) || /(?:ORGANICO|ORG|ORGANIC)/i.test(c.name));
        if (firstPriorityIdx > 0) {
            const [prio] = components.splice(firstPriorityIdx, 1);
            components.unshift(prio);
        }

        // Calcular KG por componente: primer componente -> /0.65, siguientes -> /0.85
        let finalComponents = [];
        components.forEach((c, idx) => {
            const factor = (idx === 0) ? 0.65 : 0.85;
            const kgSol = round0(kgSolicitado * c.pct);
            const kg = round0((kgSolicitado * c.pct) / factor);
            const isAlgodon = isAlgodonText(c.name);
            const qq = isAlgodon ? round0(kg / 46) : null;
            finalComponents.push({ name: c.name, kg: kg, qq: qq, kgSol: kgSol });
        });

        return finalComponents;
    }

    // VersiÃ³n HTR de parseMixComposition: usa 0.60 en lugar de 0.65 para primer componente
    function parseMixCompositionHtr(hiladoStr, kgSolicitado) {
        // Detectar porcentajes
        let pctStr = '';
        const pctMatch1 = hiladoStr.match(/\((\d{1,3}(?:\/\d{1,3})+)\s*%?\)/);
        const pctMatch2 = hiladoStr.match(/(\d{1,3}(?:\/\d{1,3})+)\s*%/);
        if (pctMatch1) pctStr = pctMatch1[1];
        else if (pctMatch2) pctStr = pctMatch2[1];

        if (!pctStr) {
            let cleanName = hiladoStr.toUpperCase().replace(/^\d+\/\d+\s+/, "").trim();
            const isPriority = /\bCOP\b/i.test(cleanName) || /(?:ORGANICO|ORG|ORGANIC)/i.test(cleanName);
            const factor = isPriority ? 0.60 : 0.85;
            let k = round0(kgSolicitado / factor);
            const isAlgodon = isAlgodonText(cleanName || "CRUDO");
            return [{ name: cleanName || "CRUDO", kg: k, qq: isAlgodon ? round0(k / 46) : null, kgSol: round0(kgSolicitado) }];
        }

        const pcts = pctStr.split('/').map(n => parseFloat(n) / 100);
        let cleanName = hiladoStr.replace(/\(?\d{1,3}(?:\/\d{1,3})+\s*%?\)?/g, "").trim();
        cleanName = cleanName.replace(/^\d+\/\d+\s+/, "").trim();
        
        // Parsear nombres
        let names = [];
        if (cleanName.includes('/')) {
            names = cleanName.split('/').map(s => s.trim()).filter(s => s !== "" && s.length > 0);
        } else {
            const materials = ['PIMA', 'LYOCELL', 'TENCEL', 'MODAL', 'TANGUIS', 'ORGANICO', 'ORGANIC', 'BCI', 'USTCP'];
            const modifiers = ['NC', 'STD', 'ORGANICO', 'ORGANIC', 'OCS'];
            let remaining = cleanName;
            while (remaining.length > 0) {
                remaining = remaining.trim();
                if (!remaining) break;
                let matched = false;
                let currentName = '';
                if (remaining.toUpperCase().startsWith('COP ')) {
                    currentName = 'COP ';
                    remaining = remaining.substring(4).trim();
                }
                for (let mat of materials) {
                    if (remaining.toUpperCase().startsWith(mat + ' ') || remaining.toUpperCase() === mat) {
                        currentName += mat;
                        remaining = remaining.substring(mat.length).trim();
                        matched = true;
                        while (remaining.length > 0) {
                            let foundMod = false;
                            for (let mod of modifiers) {
                                if (remaining.toUpperCase().startsWith(mod + ' ') || remaining.toUpperCase() === mod) {
                                    currentName += ' ' + mod;
                                    remaining = remaining.substring(mod.length).trim();
                                    foundMod = true;
                                    break;
                                }
                            }
                            if (!foundMod) break;
                        }
                        break;
                    }
                }
                if (matched && currentName.trim()) names.push(currentName.trim());
                else if (remaining.length > 0) {
                    let nextMaterialIdx = remaining.length;
                    for (let mat of materials) {
                        const idx = remaining.toUpperCase().indexOf(mat);
                        if (idx >= 0 && idx < nextMaterialIdx) nextMaterialIdx = idx;
                    }
                    if (nextMaterialIdx > 0) {
                        names.push(remaining.substring(0, nextMaterialIdx).trim());
                        remaining = remaining.substring(nextMaterialIdx);
                    } else {
                        if (remaining.trim()) names.push(remaining.trim());
                        break;
                    }
                }
            }
        }
        
        while (names.length < pcts.length) names.push(`COMP ${names.length + 1}`);
        
        let components = [];
        for (let i = 0; i < pcts.length; i++) {
            components.push({ name: names[i], pct: pcts[i], baseToken: getBaseMaterialToken(names[i]) });
        }

        // Mover primer componente COP/ORGANICO al inicio si existe
        const firstPriorityIdx = components.findIndex(c => /\bCOP\b/i.test(c.name) || /ORGANICO/i.test(c.name));
        if (firstPriorityIdx > 0) {
            const [prio] = components.splice(firstPriorityIdx, 1);
            components.unshift(prio);
        }

        // Calcular: primer componente -> /0.60, siguientes -> /0.85
        let finalComponents = [];
        components.forEach((c, idx) => {
            const factor = (idx === 0) ? 0.60 : 0.85;
            const kgSol = round0(kgSolicitado * c.pct);
            const kg = round0((kgSolicitado * c.pct) / factor);
            const isAlgodon = isAlgodonText(c.name);
            const qq = isAlgodon ? round0(kg / 46) : null;
            finalComponents.push({ name: c.name, kg: kg, qq: qq, kgSol: kgSol });
        });

        return finalComponents;
    }

    function getMixHeaders(sampleStr) {
        const comps = parseMixComposition(sampleStr, 1000);

        // Encabezados: usar directamente el nombre del componente ya procesado por parseMixComposition
        // No aplicar limpiezas adicionales porque ya estÃ¡ limpio
        return comps.map(c => {
            let name = String(c.name).toUpperCase().trim();
            return name || "COMP";
        });
    }

    // Extract base material token detectando material base + modificador
    // COP ORGANICO = TANGUIS ORGANICO (por defecto)
    // COP PIMA ORGANICO = PIMA ORGANICO (diferente tipo de algodÃ³n)
    // NO SE PUEDEN MEZCLAR
    function getBaseMaterialToken(s) {
        if (!s) return '';
        const u = String(s).toUpperCase();
        
        // Detectar material base primero
        let baseMaterial = '';
        if (u.includes('PIMA')) baseMaterial = 'PIMA';
        else if (u.includes('LYOCELL') || u.includes('TENCEL')) baseMaterial = 'LYOCELL';
        else if (u.includes('MODAL')) baseMaterial = 'MODAL';
        else if (u.includes('TANGUIS')) baseMaterial = 'TANGUIS';
        
        // Detectar modificadores (certificaciones)
        let modifier = '';
        if (u.includes('ORGANICO') || u.includes('ORGANIC')) modifier = ' ORGANICO';
        else if (u.includes('BCI')) modifier = ' BCI';
        else if (u.includes('USTCP')) modifier = ' USTCP';
        
        // Si solo tiene modificador sin base explÃ­cita, es TANGUIS por defecto
        if (!baseMaterial && modifier) {
            return 'TANGUIS' + modifier;
        }
        
        // Si tiene base y modificador, combinar
        if (baseMaterial && modifier) {
            return baseMaterial + modifier;
        }
        
        // Si solo tiene base
        if (baseMaterial) return baseMaterial;
        
        return '';
    }

    // Map a mix component name to an existing crudo bucket if possible
    function mapComponentToBase(componentName, crudoKeys) {
        const baseToken = getBaseMaterialToken(componentName);
        if (!baseToken) return null;

        // Try to find matching crudo by base token
        for (let k of crudoKeys) {
            if (getBaseMaterialToken(k) === baseToken) return k;
        }

        // Otherwise return the base token as fallback
        return baseToken;
    }

    function generateBalanceGroupTable(groups, tableType, isHtr) {
        if(groups.length === 0) return '<p class="empty-msg">Sin datos.</p>';
        let html = '';
        groups.forEach(g => {
            let totalKgSol = 0;
            let totalQQReq = 0;
            let totalQQHas = false;
            let mixTotalsKG = [];
            let mixTotalsQQ = [];
            let mixTotalsQQApplicable = [];
            let headerNames = [];

            // Detectar si el grupo contiene al menos un COP ORGANICO del cliente LLL (solo en crudo)
            const groupHasCopOrgLll = (!isHtr) && (g.items || []).some(r => {
                const cli = (r.cliente || '').toString().trim();
                return getClientCert(cli) === 'OCS' && /COP\s*ORGANICO/i.test(r.hilado || '');
            });

            if (tableType === 'mix') {
                if (g.items.length > 0) {
                    const firstYarn = g.items[0].hilado || "";
                    headerNames = getMixHeaders(firstYarn);
                    mixTotalsKG = new Array(headerNames.length).fill(0);
                    mixTotalsQQ = new Array(headerNames.length).fill(0);
                    mixTotalsQQApplicable = headerNames.map(h => isAlgodonText(h));
                }
            }

            const safeGroupName = (g.name || '').toString().replace(/\\/g, "\\\\").replace(/'/g, "\\'");
            html += `<div class="table-wrap"><div class="bal-header-row"><span class="bal-title">MATERIAL: ${g.name}</span>`;

            // Si el grupo estÃ¡ vacÃ­o, mostrar botÃ³n para eliminar
            if (g.items.length === 0) {
                html += `<button style="margin-left:auto; padding:6px 12px; font-size:11px; background:#ef4444; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="confirmDeleteGroup('${safeGroupName}', '${tableType}')">Eliminar bloque vacÃ­o</button>`;
            }
            html += `</div>`;
            
            let colCount = 0;
            if (tableType === 'mix') {
                // Columns: ORDEN, CLIENTE, TEMP, RSV, OP, HILADO, COLOR, NE, KG SOL., then 2 columns per mix component
                html += `<table><thead><tr>
                    <th class="th-base">ORDEN</th><th class="th-base">CLIENTE</th><th class="th-base">TEMP</th><th class="th-base">RSV</th><th class="th-base">OP</th>
                    <th class="th-base">HILADO</th><th class="th-base">COLOR</th><th class="th-base">NE</th><th class="th-base">KG SOL.</th>`;
                headerNames.forEach((hName, idx) => {
                    let thClass = `th-comp-${idx % 3}`;
                    html += `<th class="${thClass}">${hName}</th><th class="${thClass}">${hName} QQ</th>`;
                });
                colCount = 9 + (headerNames.length * 2);
                html += `</tr></thead><tbody>`;
            } else {
                // Columns: ORDEN, CLIENTE, TEMP, RSV, OP, HILADO, COLOR, NE, KG SOL., KG REQ., QQ REQ., (+ optional 2 for ORGANICO/TANGUIS)
                html += `<table><thead><tr>
                    <th class="th-base">ORDEN</th><th class="th-base">CLIENTE</th><th class="th-base">TEMP</th><th class="th-base">RSV</th><th class="th-base">OP</th>
                    <th class="th-base">HILADO</th><th class="th-base">COLOR</th><th class="th-base">NE</th>
                    <th class="th-base">KG SOL.</th><th class="th-base">KG REQ.</th><th class="th-base">QQ REQ.</th>`;
                if(groupHasCopOrgLll) {
                    html += `<th class="th-comp-2">QQ REQ ORGANICO 80%</th><th class="th-comp-1">QQ REQ TANGUIS 20%</th>`;
                }
                colCount = 11 + (groupHasCopOrgLll ? 2 : 0);
                html += `</tr></thead><tbody>`;
            }

            // Si estÃ¡ vacÃ­o, mostrar mensaje
            if (g.items.length === 0) {
                // Hacer que el bloque vacÃ­o acepte drops para permitir regresar items desde OTROS
                html += `<tr ondragover="allowDrop(event)" ondragenter="this.classList.add('drag-over')" ondragleave="this.classList.remove('drag-over')" ondrop="handleDropToGroup(event, '${safeGroupName}', ${g.isMezcla ? 'true' : 'false'}, ${isHtr ? 'true' : 'false'})"><td colspan="${colCount || 10}" style="text-align:center; padding:30px; color:#999;"><em>Este bloque estÃ¡ vacÃ­o</em></td></tr>`;
            }

            // Totales especiales para COP ORGANICO LLL
            let qqOrgTotal = 0;
            let qqTanTotal = 0;

            g.items.forEach(r => {
                const kgSol = round0(r.kg);
                totalKgSol += kgSol;
                const dragAttrs = `draggable="true" ondragstart="handleDragStart(event, '${r._id}')" ondrop="handleDrop(event, '${r._id}')" ondragover="allowDrop(event)"`;
                const colorText = r.colorText || '-';

                if (tableType === 'mix') {
                    const comps = isHtr ? parseMixCompositionHtr(r.hilado || "", kgSol) : parseMixComposition(r.hilado || "", kgSol);
                    const neVal = getNeFromItem(r);
                    html += `<tr ${dragAttrs}>
                        <td>${r.orden}</td><td>${r.cliente}</td><td>${r.temporada}</td><td>${r.rsv||''}</td><td>${r.op}</td>
                        <td class="hilado-cell" style="font-size:11px;">${r.hilado}</td><td style="font-size:10px;">${colorText}</td>
                        <td style="font-weight:600;">${neVal ? fmt(Math.round(neVal)) : '-'}</td>
                        <td style="font-weight:bold;">${fmtDecimal(kgSol)}</td>`;
                    comps.forEach((c, idx) => {
                        html += `<td>${fmtDecimal(c.kg)}</td><td>${fmtDecimal(c.qq)}</td>`;
                        if (idx < mixTotalsKG.length) {
                            mixTotalsKG[idx] += c.kg;
                            if (c.qq !== null && c.qq !== undefined) {
                                mixTotalsQQ[idx] += c.qq;
                                mixTotalsQQApplicable[idx] = true;
                            }
                        }
                    });
                    html += `</tr>`;
                } else {
                    const factor = isHtr ? 0.60 : 0.65;
                    const kgReq = round0(kgSol / factor);
                    const isAlgodonRow = isAlgodonText((r.hilado || r.group || ''));
                    const qqReq = isAlgodonRow ? round0(kgReq / 46) : null;
                    if (isAlgodonRow) { totalQQReq += qqReq; totalQQHas = true; }

                    // Determinar si ESTA fila es COP ORGANICO de LLL en crudo
                    const cli = (r.cliente || '').toString().trim();
                    const isCopOrgLll = (!isHtr) && (getClientCert(cli) === 'OCS') && /COP\s*ORGANICO/i.test(r.hilado || '');
                    let qqOrg = '';
                    let qqTan = '';
                    if (isCopOrgLll) {
                        qqOrg = round0(qqReq * 0.8);
                        qqTan = round0(qqReq * 0.2);
                        qqOrgTotal += qqOrg;
                        qqTanTotal += qqTan;
                    }

                    const neVal = getNeFromItem(r);
                    html += `<tr ${dragAttrs}>
                        <td>${r.orden}</td><td>${r.cliente}</td><td>${r.temporada}</td><td>${r.rsv || ''}</td><td>${r.op}</td>
                        <td class="hilado-cell">${r.hilado}</td><td style="font-size:10px;">${colorText}</td>
                        <td style="font-weight:600;">${neVal ? fmt(Math.round(neVal)) : '-'}</td>
                        <td>${fmtDecimal(kgSol)}</td>
                        <td style="background:#f8fafc; font-weight:bold;">${fmtDecimal(kgReq)}</td>
                        <td style="background:#f0f9ff;">${fmtDecimal(qqReq)}</td>`;
                    if (groupHasCopOrgLll) html += `<td style="color:#15803d;">${qqOrg !== '' ? fmtDecimal(qqOrg) : ''}</td><td style="color:#b45309;">${qqTan !== '' ? fmtDecimal(qqTan) : ''}</td>`;
                    html += `</tr>`;
                }
            });

            if (g.items.length > 0) {
                 if (tableType === 'mix') {
                     const groupNeVal = computeWeightedNe(g.items || []);
                     const groupNeDisplay = (groupNeVal !== null && !isNaN(groupNeVal)) ? fmt(Math.round(groupNeVal)) : '-';
                     // Para mix: poner TOTAL abarcando hasta COLOR, luego NE, luego KG SOL
                     html += `<tr class="bal-subtotal"><td colspan="7" style="text-align:right;">TOTAL ${g.name}:</td><td>${groupNeDisplay}</td><td>${fmt(totalKgSol)}</td>`;
                     mixTotalsKG.forEach((tk, i) => {
                         const qqDisplay = mixTotalsQQApplicable[i] ? fmt(mixTotalsQQ[i]) : '-';
                         html += `<td>${fmt(tk)}</td><td>${qqDisplay}</td>`;
                     });
                     html += `</tr>`;
                  } else {
                      const factor = isHtr ? 0.60 : 0.65;
                      const subKgReq = round0(totalKgSol / factor);
                      const subQQReq = totalQQHas ? round0(totalQQReq) : null;
                      // Calcular Ne del grupo
                      const groupNeVal = computeWeightedNe(g.items || []);
                      const groupNeDisplay = (groupNeVal !== null && !isNaN(groupNeVal)) ? fmt(Math.round(groupNeVal)) : '-';
                      html += `<tr class="bal-subtotal"><td colspan="7" style="text-align:right;">TOTAL ${g.name}:</td><td>${groupNeDisplay}</td><td>${fmt(totalKgSol)}</td><td>${fmt(subKgReq)}</td><td>${fmt(subQQReq)}</td>`;
                      if(groupHasCopOrgLll) html += `<td>${fmt(qqOrgTotal)}</td><td>${fmt(qqTanTotal)}</td>`;
                      html += `</tr>`;
                  }
            }
            html += `</tbody></table></div>`;
        });
        return html;
    }

    // Normaliza claves para comparaciones en resumen
    function normalizeSummaryKey(s) {
        if (!s) return '';
        return String(s)
            .normalize('NFD')
            .replace(/[\u0300-\u036f]/g, '')
            .toUpperCase()
            .replace(/\s+/g, ' ')
            .trim();
    }

    // Orden y alias de resumen (tabla + modal)
    const SUMMARY_ORDER = [
        'ALGODÓN ORGANICO (OCS)',
        'ALGODÓN UPLAND USTCP',
        'ALGODÓN PIMA NC',
        'ALGODÓN UPLAND',
        'ALGODÓN TANGUIS NC BCI',
        'ALGODÓN ORGANICO PIMA (OCS)',
        'LYOCELL STD',
        'PES RECICLADO',
        'MODAL',
        'ALGODÓN ORGANICO PIMA (GOTS)',
        'CAÑAMO'
    ];

    const SUMMARY_ORDER_HTR = [
        'ALGODÓN PIMA NC',
        'ALGODÓN ORGANICO (OCS)',
        'ALGODÓN ORGANICO PIMA (OCS)',
        'LYOCELL STD',
        'PES RECICLADO'
    ];

    const SUMMARY_ORDER_MAP = new Map();
    SUMMARY_ORDER.forEach((name, idx) => {
        const norm = normalizeSummaryKey(name);
        SUMMARY_ORDER_MAP.set(norm, idx);
        const noAlgodon = norm.replace(/^ALGODON\s+/, '');
        if (noAlgodon !== norm) SUMMARY_ORDER_MAP.set(noAlgodon, idx);
    });

    const SUMMARY_ORDER_HTR_MAP = new Map();
    SUMMARY_ORDER_HTR.forEach((name, idx) => {
        const norm = normalizeSummaryKey(name);
        SUMMARY_ORDER_HTR_MAP.set(norm, idx);
        const noAlgodon = norm.replace(/^ALGODON\s+/, '');
        if (noAlgodon !== norm) SUMMARY_ORDER_HTR_MAP.set(noAlgodon, idx);
    });

    const SUMMARY_DISPLAY_BY_NORM = new Map();
    SUMMARY_ORDER.forEach(name => SUMMARY_DISPLAY_BY_NORM.set(normalizeSummaryKey(name), name));

    const SUMMARY_ALIAS_TO_DISPLAY = new Map([
        ['TANGUIS ORGANICO (OCS)', 'ALGODÓN ORGANICO (OCS)'],
        ['PIMA ORGANICO (OCS)', 'ALGODÓN ORGANICO PIMA (OCS)'],
        ['PIMA (OCS)', 'ALGODÓN ORGANICO PIMA (OCS)'],
        ['PIMA OCS', 'ALGODÓN ORGANICO PIMA (OCS)'],
        ['PIMA ORGANICO (GOTS)', 'ALGODÓN ORGANICO PIMA (GOTS)'],
        ['TANGUIS BCI', 'ALGODÓN TANGUIS NC BCI'],
        ['TANGUIS USTCP', 'ALGODÓN UPLAND USTCP'],
        ['LYOCELL', 'LYOCELL STD'],
        ['PES REPREVE', 'PES RECICLADO'],
        ['PES PREPREVE', 'PES RECICLADO'],
        ['PES RECICLADO', 'PES RECICLADO'],
        ['PIMA', 'ALGODÓN PIMA NC'],
        ['TANGUIS', 'ALGODÓN UPLAND']
    ].map(([k, v]) => [normalizeSummaryKey(k), v]));

    function mapSummaryDisplayName(name) {
        const norm = normalizeSummaryKey(name);
        if (SUMMARY_ALIAS_TO_DISPLAY.has(norm)) return SUMMARY_ALIAS_TO_DISPLAY.get(norm);
        if (SUMMARY_DISPLAY_BY_NORM.has(norm)) return SUMMARY_DISPLAY_BY_NORM.get(norm);
        return name;
    }

    function getSummaryOrderIndex(name, isHtr) {
        const display = mapSummaryDisplayName(name);
        const norm = normalizeSummaryKey(display);
        const orderMap = isHtr ? SUMMARY_ORDER_HTR_MAP : SUMMARY_ORDER_MAP;
        if (orderMap.has(norm)) return orderMap.get(norm);
        return Number.POSITIVE_INFINITY;
    }

    function generateSummaryTable(groups, isHtr) {
        if (!groups || groups.length === 0) return '';

        // Flatten items and detect crudo buckets within the provided groups
        const allItems = groups.flatMap(g => g.items || []);
        const crudoKeys = groups.filter(g => !g.isMezcla).map(g => g.name);

        const materialsMap = new Map();

        function addToMat(matName, kg, orden) {
            if(!materialsMap.has(matName)) materialsMap.set(matName, { kg: 0, orders: new Set() });
            const entry = materialsMap.get(matName);
            entry.kg += kg;
            if (orden) entry.orders.add(orden);
        }

        allItems.forEach(r => {
            const isMezcla = Boolean(r.isMezcla);
            if (!isMezcla) {
                const factor = isHtr ? 0.60 : 0.65;
                const kReq = round0(r.kg / factor);
                // Si el hilado contiene (OCS) o (GOTS), incluir en el nombre del grupo
                let groupName = r.group;
                const hilado = (r.hilado || '').toUpperCase();
                if (/(?:ORGANICO|ORG|ORGANIC)/i.test(hilado)) {
                        if (/\(OCS\)/.test(hilado)) {
                        groupName = r.group + ' (OCS)';
                    } else if (/\(GOTS\)/.test(hilado)) {
                        groupName = r.group + ' (GOTS)';
                    }
                }
                addToMat(groupName, kReq, r.orden);
            } else {
                const comps = isHtr ? parseMixCompositionHtr(r.hilado || "", r.kg) : parseMixComposition(r.hilado || "", r.kg);
                comps.forEach(c => {
                    const mapped = mapComponentToBase(c.name, crudoKeys);
                    const target = mapped || c.name;
                    // Si el hilado contiene (OCS) o (GOTS), incluir en el nombre
                    let finalTarget = target;
                    const hilado = (r.hilado || '').toUpperCase();
                    if (/(?:ORGANICO|ORG|ORGANIC)/i.test(c.name)) {
                        if (/\(OCS\)/.test(hilado)) {
                            finalTarget = target + ' (OCS)';
                        } else if (/\(GOTS\)/.test(hilado)) {
                            finalTarget = target + ' (GOTS)';
                        }
                    }
                    addToMat(finalTarget, c.kg, r.orden);
                });
            }
        });

        // Orden preferido para el resumen (si existe en los datos)

        // Reagrupar usando el nombre de display para que también se sumen kilos de alias
        const aggregated = new Map(); // displayName -> { kg, orders:Set }
        materialsMap.forEach((data, rawName) => {
            const display = mapSummaryDisplayName(rawName);
            if (!aggregated.has(display)) aggregated.set(display, { kg: 0, orders: new Set() });
            const agg = aggregated.get(display);
            agg.kg += data.kg;
            data.orders.forEach(o => agg.orders.add(o));
        });

        let rows = Array.from(aggregated.entries()).map(([name, data]) => {
            const isAlgodon = isAlgodonText(name);
            const ingresos = isAlgodon ? (data.kg / 5800) : null;
            const dias = isAlgodon ? (ingresos * 1.5) : null;
            return {
                name,
                kg: data.kg,
                qq: isAlgodon ? round0(data.kg / 46) : null,
                ingresos: ingresos,
                dias: dias
            };
        });

        rows.sort((a, b) => {
            const aIdx = getSummaryOrderIndex(a.name, isHtr);
            const bIdx = getSummaryOrderIndex(b.name, isHtr);
            if (aIdx !== bIdx) return aIdx - bIdx;
            return b.kg - a.kg;
        });

        // Procesar items que estÃ©n en 'OTROS' cuando el resumen no los incluye en grupos
        // (ej: HTR donde se excluye OTROS). Evita duplicar si OTROS ya estÃ¡ presente.
        const hasOtrosGroup = groups.some(g => (g.name || '').toString().toUpperCase() === 'OTROS');
        try {
            if (!hasOtrosGroup) {
                const srcBase = isHtr ? GLOBAL_DATA.htr : GLOBAL_DATA.nuevo;
                const src = (srcBase && srcBase.some(r => r.highlight)) ? srcBase.filter(r => r.highlight) : (srcBase || []);
                const otrosItems = (src || []).filter(it => (it.group || '').toString().toUpperCase() === 'OTROS');
                if (otrosItems && otrosItems.length > 0) {
                    otrosItems.forEach(it => {
                        if (it.isMezcla) {
                            const comps = isHtr ? parseMixCompositionHtr(it.hilado || '', it.kg) : parseMixComposition(it.hilado || '', it.kg);
                            comps.forEach(c => {
                                const mapped = mapComponentToBase(c.name, crudoKeys);
                                const target = mapped || c.name;
                                // Si el hilado original tiene certificaciÃ³n, propagarla al nombre
                                let finalTarget = target;
                                const hiladoStr = (it.hilado || '').toUpperCase();
                                if (/(?:ORGANICO|ORG|ORGANIC)/i.test(c.name)) {
                                    if (/\(OCS\)/.test(hiladoStr)) finalTarget = target + ' (OCS)';
                                    else if (/\(GOTS\)/.test(hiladoStr)) finalTarget = target + ' (GOTS)';
                                }
                                addToMat(finalTarget, c.kg, it.orden);
                            });
                        } else {
                            // No es mezcla: tratar como un material Ãºnico llamado 'OTROS'
                            const factor = isHtr ? 0.60 : 0.65;
                            const kgReq = round0((it.kg || 0) / factor);
                            addToMat('OTROS', kgReq, it.orden);
                        }
                    });
                }
            }
        } catch(e) { console.error('generateSummaryTable OTROS aggregation error', e); }

        // BotÃ³n para agregar a OTROS (usa currentTab para seleccionar candidatos)
        const targetOtros = btoa('OTROS');
        const addBtn = `<div style="margin-bottom:8px; text-align:right;"><button style="padding:6px 10px; background:#10b981; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="openAddComponentModal('${targetOtros}', ${!isHtr})">+ Agregar a OTROS</button></div>`;

        let html = `<div class="summary-title">RESUMEN DE TOTALES (KG REQUERIDO)</div><div class="table-wrap">` + addBtn + `<table class="sum-table"><thead><tr><th>TOTAL</th><th>KG REQ</th><th>QQ REQ</th><th>Nro. Ingresos</th><th>Nro. Dias</th></tr></thead><tbody>`;

        let grandKg = 0, grandQQ = 0, grandIngresos = 0, grandDias = 0;
        let grandQQHas = false;
        let grandIngresosHas = false;
        let grandDiasHas = false;
        rows.forEach(r => {
            const encoded = btoa(r.name);
            // Nota: El botÃ³n eliminar se muestra ahora dentro del modal de detalle.
            html += `<tr><td style="text-align:left; font-weight:600;"> <span style="cursor:pointer;" onclick="showSummaryDetails('${encoded}')">${r.name}</span></td><td>${fmt(round0(r.kg))}</td><td>${fmt(r.qq)}</td><td>${fmt(r.ingresos)}</td><td>${fmt(r.dias)}</td></tr>`;
            grandKg += r.kg;
            if (r.qq !== null && r.qq !== undefined) {
                grandQQ += r.qq;
                grandQQHas = true;
            }
            if (r.ingresos !== null && r.ingresos !== undefined) {
                grandIngresos += r.ingresos;
                grandIngresosHas = true;
            }
            if (r.dias !== null && r.dias !== undefined) {
                grandDias += r.dias;
                grandDiasHas = true;
            }
        });

        const totalQQDisplay = grandQQHas ? fmt(round0(grandQQ)) : '-';
        const totalIngresosDisplay = grandIngresosHas ? fmt(grandIngresos) : '-';
        const totalDiasDisplay = grandDiasHas ? fmt(grandDias) : '-';
        html += `<tr class="total-row"><td style="text-align:right;">TOTAL GENERAL</td><td>${fmt(round0(grandKg))}</td><td>${totalQQDisplay}</td><td>${totalIngresosDisplay}</td><td>${totalDiasDisplay}</td></tr></tbody></table></div>`;

        // NE metrics se muestran en el footer para evitar duplicados en la tabla

        return html;
    }

    function deleteGroupByEncoded(encodedName, isHtr) {
        try {
            const name = atob(encodedName);
            const tableType = isHtr ? 'htr' : 'pure';
            deleteGroup(name, tableType);
        } catch(e) { console.error('deleteGroupByEncoded', e); }
    }

    function showSummaryDetails(encodedGroupName) {
        const groupName = atob(encodedGroupName);
        const isCrudo = (GLOBAL_DATA.currentTab === 'crudo');
        const activeData = isCrudo 
            ? (GLOBAL_DATA.nuevo.some(r=>r.highlight) ? GLOBAL_DATA.nuevo.filter(r=>r.highlight) : GLOBAL_DATA.nuevo)
            : (GLOBAL_DATA.htr.some(r=>r.highlight) ? GLOBAL_DATA.htr.filter(r=>r.highlight) : GLOBAL_DATA.htr);

        // derive crudo keys
        const groups = getGroupsFromData(activeData);
        const crudoKeys = groups.filter(g=>!g.isMezcla).map(g=>g.name);

        const crudoContribs = [];
        const mezclaContribs = [];
        let totalCrudoKgReq = 0;
        let totalMezclaKgReq = 0;
        const groupNorm = normalizeSummaryKey(groupName);
        const isHtr = !isCrudo;
        const baseFactor = isHtr ? 0.60 : 0.65;

        function getNonMixDisplayName(item) {
            let rawGroup = item.group || '';
            const hiladoUpper = (item.hilado || '').toUpperCase();
            if (/(?:ORGANICO|ORG|ORGANIC)/i.test(hiladoUpper)) {
                if (/\(OCS\)/.test(hiladoUpper)) rawGroup = (item.group || '') + ' (OCS)';
                else if (/\(GOTS\)/.test(hiladoUpper)) rawGroup = (item.group || '') + ' (GOTS)';
            }
            return mapSummaryDisplayName(rawGroup);
        }

        function getComponentDisplayName(componentName, hiladoUpper) {
            const mapped = mapComponentToBase(componentName, crudoKeys);
            let target = mapped || componentName;
            if (/(?:ORGANICO|ORG|ORGANIC)/i.test(componentName)) {
                if (/\(OCS\)/.test(hiladoUpper)) target = target + ' (OCS)';
                else if (/\(GOTS\)/.test(hiladoUpper)) target = target + ' (GOTS)';
            }
            return mapSummaryDisplayName(target);
        }

        activeData.forEach(item => {
            const isMezcla = Boolean(item.isMezcla);
            const hiladoUpper = (item.hilado || '').toUpperCase();

            if (!isMezcla) {
                const kReq = round0((item.kg || 0) / baseFactor);
                const displayName = getNonMixDisplayName(item);
                if (normalizeSummaryKey(displayName) === groupNorm) {
                    const isAlgodon = isAlgodonText(displayName);
                    const qqReq = isAlgodon ? round0(kReq/46) : null;
                    crudoContribs.push({ item, kgReq: kReq, qqReq: qqReq });
                    totalCrudoKgReq += kReq;
                }
            } else {
                const comps = isHtr ? parseMixCompositionHtr(item.hilado || '', item.kg) : parseMixComposition(item.hilado || '', item.kg);
                const matching = comps.filter(c => normalizeSummaryKey(getComponentDisplayName(c.name, hiladoUpper)) === groupNorm);
                if (matching.length) {
                    const subtotal = matching.reduce((s,c)=>s+c.kg,0);
                    mezclaContribs.push({ item, components: matching, kgReqSubtotal: subtotal });
                    totalMezclaKgReq += subtotal;
                }
            }
        });

        let modalHTML = `<div style="font-size: 12px;">`;
        // BotÃ³n para agregar/mover hilados hacia este grupo y botÃ³n eliminar dentro del modal
        modalHTML += `<div style="display:flex; justify-content:flex-end; gap:8px; margin-bottom:10px;">
            <button style="padding:6px 10px; background:#10b981; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="openAddComponentModal('${btoa(groupName)}', ${isCrudo})">+ Agregar/Mover</button>
            <button style="padding:6px 10px; background:#ef4444; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="if(confirm('Eliminar grupo ${groupName}?')){ deleteGroupByEncoded('${encodedGroupName}', ${isCrudo ? 'false' : 'true'}); closeDetailModal(); }">&#x1F5D1;&#xFE0F; Eliminar</button>
        </div>`;
        modalHTML += `<h3 style="color: #0f172a; margin-bottom: 20px;">DETALLES DE: ${groupName}</h3>`;

        if (crudoContribs.length) {
            modalHTML += `<div style="margin-bottom:20px;"><h4 style="color:#2563eb;">CRUDOS (${crudoContribs.length})</h4>`;
            crudoContribs.forEach(c => {
                modalHTML += `<div class="modal-item"><div class="modal-item-label">Orden: ${c.item.orden}</div><div class="modal-item-value">Hilado: ${c.item.hilado || ''}</div><div class="modal-item-value">KG Sol: ${fmt(c.item.kg)} | KG Req: ${fmt(c.kgReq)} | QQ: ${fmt(c.qqReq)}</div></div>`;
            });
            modalHTML += `<div style="background:#e0f2fe; border-left:4px solid #2563eb; padding:10px; margin-top:8px;">Subtotal Crudos: ${fmt(totalCrudoKgReq)} kg</div></div>`;
        }

        if (mezclaContribs.length) {
            modalHTML += `<div style="margin-bottom:20px;"><h4 style="color:#ec4899;">MEZCLAS (${mezclaContribs.length})</h4>`;
            mezclaContribs.forEach(m => {
                modalHTML += `<div class="modal-item"><div class="modal-item-label">Orden: ${m.item.orden}</div><div class="modal-item-value">Mezcla: ${m.item.hilado}</div><div class="modal-item-value">KG Sol: ${fmt(m.item.kg)}</div><div style="margin-top:8px; padding-left:10px; border-left:2px solid #ec4899;">`;
                m.components.forEach(c => {
                    modalHTML += `<div style="font-size:11px; color:#64748b; margin-bottom:4px;">- ${c.name}: KG Sol: ${fmt(c.kgSol)} | KG Req: ${fmt(c.kg)} | QQ: ${fmt(c.qq)}</div>`;
                });
                modalHTML += `</div></div>`;
            });
            modalHTML += `<div style="background:#fce7f3; border-left:4px solid #ec4899; padding:10px; margin-top:8px;">Subtotal Mezclas: ${fmt(totalMezclaKgReq)} kg</div></div>`;
        }

        const grand = totalCrudoKgReq + totalMezclaKgReq;
        modalHTML += `<div style="background:#1e293b; color:white; padding:12px; border-radius:4px; text-align:center; margin-top:10px;"><strong>TOTAL KG REQUERIDO: ${fmt(grand)}</strong></div>`;
        modalHTML += `</div>`;

        document.getElementById('modalTitle').textContent = `Detalles: ${groupName}`;
        document.getElementById('modalBody').innerHTML = modalHTML;
        document.getElementById('detailModal').classList.add('show');
    }

    function updateFooterTotals() {
        const container = document.getElementById('dynamic-footer-grid');
        container.innerHTML = ""; 
        let totalKgSol = 0, totalQQReq = 0; 
        const isCrudo = (GLOBAL_DATA.currentTab === 'crudo');
        const activeData = isCrudo 
            ? (GLOBAL_DATA.nuevo.some(r=>r.highlight) ? GLOBAL_DATA.nuevo.filter(r=>r.highlight) : GLOBAL_DATA.nuevo)
            : (GLOBAL_DATA.htr.some(r=>r.highlight) ? GLOBAL_DATA.htr.filter(r=>r.highlight) : GLOBAL_DATA.htr);

        // Calcular mÃ©tricas NE globales siempre (usar arrays completos, no solo highlights)
        const neCrudoVal = computeWeightedNe(GLOBAL_DATA.nuevo || []);
        const neHtrVal = computeWeightedNe(GLOBAL_DATA.htr || []);
        const neTotalVal = computeWeightedNe(((GLOBAL_DATA.nuevo || []).concat(GLOBAL_DATA.htr || [])));
        const neCrudoDisplay = (neCrudoVal !== null && !isNaN(neCrudoVal)) ? fmt(Math.round(neCrudoVal)) : '-';
        const neHtrDisplay = (neHtrVal !== null && !isNaN(neHtrVal)) ? fmt(Math.round(neHtrVal)) : '-';
        const neTotalDisplay = (neTotalVal !== null && !isNaN(neTotalVal)) ? fmt(Math.round(neTotalVal)) : '-';

        if (isCrudo) {
            document.getElementById('footer-context-label').innerText = "CRUDO / MEZCLA";
            let sumPureSol = 0, sumMixSol = 0;
            activeData.forEach(r => {
                // Usar el flag isMezcla stored, no re-detectarlo
                if (r.isMezcla) { 
                    sumMixSol += r.kg; 
                } else { 
                    sumPureSol += r.kg; 
                }
            });
            
            // Calcular totales (sin dividir)
            const totalCrudoGlobal = sumPureSol + sumMixSol;
            
            // Obtener total HTR
            const activeHtr = GLOBAL_DATA.htr.some(r=>r.highlight) ? GLOBAL_DATA.htr.filter(r=>r.highlight) : GLOBAL_DATA.htr;
            const totalHtr = activeHtr.reduce((a,b)=>a+b.kg, 0);
            
            // Total consolidado final
            const totalConsolidado = totalCrudoGlobal + totalHtr;
            
            container.innerHTML = `
                <div class="footer-item"><label>TOTAL CRUDO (SOL)</label><div class="val val-crudo">${fmt(sumPureSol)}</div></div>
                <div class="footer-item"><label>TOTAL MEZCLAS (SOL)</label><div class="val val-mezcla">${fmt(sumMixSol)}</div></div>
                <div class="footer-item"><label>TOTAL CRUDO GLOBAL</label><div class="val" style="color:#2563eb;">${fmt(totalCrudoGlobal)}</div></div>
                <div class="footer-item"><label>TOTAL HTR</label><div class="val" style="color:#dc2626;">${fmt(totalHtr)}</div></div>
                <div class="footer-item" style="border-left: 1px solid rgba(255,255,255,0.2); padding-left: 20px;"><label>TOTAL CONSOLIDADO</label><div class="val val-total" id="res-total-kg">${fmt(totalConsolidado)}</div></div>
                <div style="flex-basis:100%; height:8px;"></div>
                <div class="footer-item"><label>Ne CRUDO</label><div class="val">${neCrudoDisplay}</div></div>
                <div class="footer-item"><label>Ne HTR</label><div class="val">${neHtrDisplay}</div></div>
                <div class="footer-item"><label>Ne PROMEDIO</label><div class="val">${neTotalDisplay}</div></div>`;
        } else {
            document.getElementById('footer-context-label').innerText = "HTR";
            totalKgSol = round0(activeData.reduce((a,b)=>a+b.kg, 0));
            let totalKgSolAlgodon = 0;
            activeData.forEach(r => {
                if (isAlgodonText((r.hilado || r.group || ''))) totalKgSolAlgodon += (r.kg || 0);
            });
            const totalQQDirect = totalKgSolAlgodon > 0 ? round0(totalKgSolAlgodon / 46) : null;
            container.innerHTML = `
                <div class="footer-item"><label>TOTAL KG HTR</label><div class="val" id="res-total-kg">${fmt(totalKgSol)}</div></div>
                <div class="footer-item"><label>TOTAL QQ (DIRECTO)</label><div class="val" id="res-total-qq">${fmt(totalQQDirect)}</div></div>
                <div style="flex-basis:100%; height:8px;"></div>
                <div class="footer-item"><label>Ne CRUDO</label><div class="val">${neCrudoDisplay}</div></div>
                <div class="footer-item"><label>Ne HTR</label><div class="val">${neHtrDisplay}</div></div>
                <div class="footer-item"><label>Ne PROMEDIO</label><div class="val">${neTotalDisplay}</div></div>`;
        }
    }

    // --- Add/Move Components modal functions ---
    function openAddComponentModal(encodedTargetGroup, isCrudo) {
        const targetGroup = atob(encodedTargetGroup);
        // Decide source based on provided flag; fallback to currentTab
        const useCrudo = (typeof isCrudo !== 'undefined') ? Boolean(isCrudo) : (GLOBAL_DATA.currentTab === 'crudo');
        // Build list of candidate items from active data (crudo or htr)
        const srcCrudo = (GLOBAL_DATA.nuevo || []).concat(GLOBAL_DATA.nuevoOriginal || []);
        const srcHtr = (GLOBAL_DATA.htr || []).concat(GLOBAL_DATA.htrOriginal || []);
        const candidates = useCrudo ? srcCrudo : srcHtr;

        // Remove duplicates by _id and exclude items already in OTROS
        const seen = new Set();
        const unique = [];
        candidates.forEach(c => {
            if (!c || !c._id) return;
            if ((c.group || '').toString().toUpperCase() === 'OTROS') return; // skip already OTROS
            if (!seen.has(c._id)) { seen.add(c._id); unique.push(c); }
        });

        let html = `<div style="max-height:60vh; overflow:auto;">`;
        html += `<p>Selecciona el hilado que deseas mover al grupo <strong>${targetGroup}</strong>:</p>`;
        if (unique.length === 0) {
            html += `<p class="empty-msg">Sin filas disponibles.</p>`;
        } else {
            unique.forEach(u => {
                const safeId = (u._id || '');
                html += `<div style="padding:10px; border-left:4px solid #e2e8f0; margin-bottom:8px; display:flex; justify-content:space-between; align-items:center; background:#f8fafc;">
                            <div style="font-size:12px;"><strong>Orden:</strong> ${u.orden || '-'}<br><strong>Hilado:</strong> ${u.hilado || '-'}<br><strong>Grupo:</strong> ${u.group || '-'}</div>
                            <div style="display:flex; gap:8px;">
                                <button style="padding:6px 10px; background:#3b82f6; color:white; border:none; border-radius:4px; cursor:pointer;" onclick="confirmMoveItem('${safeId}', '${btoa(targetGroup)}')">Mover</button>
                            </div>
                         </div>`;
            });
        }
        html += `</div>`;
        document.getElementById('addComponentBody').innerHTML = html;
        document.getElementById('addComponentModal').classList.add('show');
    }

    function closeAddComponentModal() { document.getElementById('addComponentModal').classList.remove('show'); }

    function confirmMoveItem(itemId, encodedTargetGroup) {
        const targetGroup = atob(encodedTargetGroup);
        if (!confirm(`Â¿Desea mover este hilado al grupo ${targetGroup}?`)) return;
        moveItemToGroup(itemId, targetGroup);
    }

    function moveItemToGroup(itemId, targetGroup) {
        // Find item in any array and move
        function updateInArray(arr) {
            for (let i=0;i<arr.length;i++) {
                if (arr[i]._id === itemId) {
                    // Recalculate derived fields from hilado first
                    recalcItemFields(arr[i]);
                    // Then force the new group (and mark as non-mezcla if moving to OTROS)
                    arr[i].group = targetGroup;
                    if ((targetGroup || '').toString().toUpperCase() === 'OTROS') {
                        arr[i].isMezcla = false;
                        // Assign title to OTROS for Title module
                        arr[i].titulo = 'OTROS';
                    } else {
                        // update titulo if needed from hilado
                        const m = (arr[i].hilado || '').toString().match(/^\s*(\d+\/\d+)/);
                        arr[i].titulo = m ? m[1] : arr[i].titulo || 'SIN TITULO';
                    }
                    return true;
                }
            }
            return false;
        }
        let found = updateInArray(GLOBAL_DATA.nuevo) || updateInArray(GLOBAL_DATA.nuevoOriginal) || updateInArray(GLOBAL_DATA.htr) || updateInArray(GLOBAL_DATA.htrOriginal);
        console.log('moveItemToGroup: moved', itemId, '->', targetGroup, 'found=', found);
        if (!found) { alert('No se encontrÃ³ la fila.'); closeAddComponentModal(); return; }

        // Recompute highlights and rerender
        identifyIncludedRows(GLOBAL_DATA.nuevo, GLOBAL_DATA.excelTotals.crudo);
        identifyIncludedRows(GLOBAL_DATA.htr, GLOBAL_DATA.excelTotals.htr);
        sortDataArray(GLOBAL_DATA.nuevo);
        sortDataArray(GLOBAL_DATA.htr);
        renderPCPModule();
        renderBalanceModule();
        renderTituloModule();
        updateFooterTotals();
        closeAddComponentModal();
        alert('Hilado movido a ' + targetGroup);

        // Debug: asegurar que el grupo OTROS aparece en la vista cuando corresponde
        try {
            const grouped = getGroupsFromData(GLOBAL_DATA.nuevo.concat(GLOBAL_DATA.htr));
            const hasOtros = grouped.some(g => (g.name || '').toString().toUpperCase() === 'OTROS');
            console.log('moveItemToGroup: grouped contains OTROS?', hasOtros);
            // Si encontramos items con group OTROS pero no hay grupo en grouped (caso raro), forzar placeholder
            const anyItemOtros = GLOBAL_DATA.nuevo.concat(GLOBAL_DATA.htr).some(it => (it.group || '').toString().toUpperCase() === 'OTROS');
            if (anyItemOtros && !hasOtros) {
                // Forzar placeholder en crudo por defecto
                addEmptyGroup('OTROS', 'pure');
                renderBalanceModule();
                console.warn('moveItemToGroup: Forzado placeholder OTROS porque se detectÃ³ inconsistencia');
            }
        } catch(e) { console.error('moveItemToGroup debug error', e); }
    }

    // --- DRAG AND DROP HANDLERS ---
    let draggedId = null;
    let draggedItemRef = null;

    function handleDragStart(e, id) {
        draggedId = id;
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', id);
        const tr = e.target.closest('tr');
        if (tr) {
            tr.classList.add('dragging');
        }
        // Guardar referencia al objeto arrastrado (buscar en arrays principales u originales)
        draggedItemRef = null;
        try {
            const searchArrays = [GLOBAL_DATA.nuevo, GLOBAL_DATA.nuevoOriginal, GLOBAL_DATA.htr, GLOBAL_DATA.htrOriginal];
            for (let arr of searchArrays) {
                if (!arr) continue;
                const found = arr.find(x => x && x._id === id);
                if (found) { draggedItemRef = { item: found, fromArray: arr }; break; }
            }
        } catch(ex) { console.warn('handleDragStart ref search error', ex); }
        console.log('Arrastrando hilado:', id, 'refFound=', !!draggedItemRef);
    }

    function allowDrop(e) {
        e.preventDefault();
        e.stopPropagation();
        e.dataTransfer.dropEffect = 'move';
        const tr = e.target.closest('tr');
        if (tr && !tr.classList.contains('dragging')) {
            tr.classList.add('drag-over');
        }
    }

    document.addEventListener('dragleave', function(e) {
        if(e.target.tagName === 'TD') {
            const tr = e.target.closest('tr');
            if(tr) tr.classList.remove('drag-over');
        }
    });

    function handleDrop(e, targetId) {
        e.preventDefault();
        const tr = e.target.closest('tr');
        if (tr) tr.classList.remove('drag-over');
        
        if (!draggedId || draggedId === targetId) {
            draggedId = null;
            return;
        }

        try {
            const isCrudo = (GLOBAL_DATA.currentTab === 'crudo');
            const dataArr = isCrudo ? GLOBAL_DATA.nuevo : GLOBAL_DATA.htr;

            const draggedIndex = dataArr.findIndex(i => i._id === draggedId);
            const targetIndex = dataArr.findIndex(i => i._id === targetId);

            if (draggedIndex === -1 || targetIndex === -1) {
                console.warn('No se encontrÃ³ hilado a mover');
                draggedId = null;
                return;
            }

            const draggedItem = dataArr[draggedIndex];
            const targetItem = dataArr[targetIndex];
            const originalGroup = draggedItem.group;
            const originalTitulo = draggedItem.titulo;
            const originalIsMezcla = !!draggedItem.isMezcla;
            const originalIndex = draggedIndex; // Guardar Ã­ndice original

            // Detectar si estamos en el mÃ³dulo TÃ­tulo
            let tituloModuleActive = false;
            try {
                tituloModuleActive = document.getElementById('mod-titulo').classList.contains('active');
            } catch(e) {}

            // Si estamos en TÃ­tulo: cambiar `titulo` (no `group`)
            // Si estamos en Material: cambiar `group` (como siempre)
            if (tituloModuleActive) {
                draggedItem.titulo = targetItem.titulo;
                console.log('Hilado movido a tÃ­tulo:', draggedItem.titulo);
            } else {
                draggedItem.group = targetItem.group;
                draggedItem.isMezcla = targetItem.isMezcla;
                console.log('Hilado movido a grupo:', draggedItem.group);
            }

            // LOGICA DE REORDENAMIENTO: mover el item a despuÃ©s del destino
            dataArr.splice(draggedIndex, 1);
            let newTargetIndex = dataArr.findIndex(i => i._id === targetId);
            if (newTargetIndex >= 0) {
                dataArr.splice(newTargetIndex + 1, 0, draggedItem);
            } else {
                dataArr.push(draggedItem);
            }
            
            // Detectar si el bloque original quedÃ³ vacÃ­o (grupo o tÃ­tulo segÃºn mÃ³dulo)
            let remainingInBlock = [];
            let blockName = '';
            if (tituloModuleActive) {
                remainingInBlock = dataArr.filter(r => (r.titulo || '').toString() === originalTitulo);
                blockName = originalTitulo;
            } else {
                remainingInBlock = dataArr.filter(r => (r.group || '').toString() === originalGroup);
                blockName = originalGroup;
            }

            if (remainingInBlock.length === 0 && blockName && blockName.toUpperCase() !== 'OTROS') {
                // Bloque quedÃ³ vacÃ­o, mostrar diÃ¡logo SIN renderizar aÃºn
                setTimeout(() => {
                    if (confirm(`Â¿Deseas eliminar el bloque vacÃ­o "${blockName}" o mantenerlo vacÃ­o?\n\nCancelar = Mantener vacÃ­o\nAceptar = Eliminar`)) {
                        // Usuario aceptÃ³: eliminar el bloque
                        if (tituloModuleActive) {
                            // En TÃ­tulo, solo re-renderizar (el bloque no aparecerÃ¡ porque no hay items con ese titulo)
                            renderTituloModule();
                        } else {
                            deleteGroup(blockName, isCrudo ? (originalIsMezcla ? 'mix' : 'pure') : 'htr');
                        }
                    } else {
                        // Usuario cancelÃ³: revertir cambios (titulo Y posiciÃ³n en array)
                        if (tituloModuleActive) {
                            // Revertir el cambio de tÃ­tulo al original
                            draggedItem.titulo = originalTitulo;
                            // Revertir posiciÃ³n en array: remover de posiciÃ³n actual e insertar en posiciÃ³n original
                            const currentIdx = dataArr.findIndex(i => i._id === draggedId);
                            if (currentIdx !== -1) {
                                dataArr.splice(currentIdx, 1);
                                // Asegurar que no sobrepasamos el rango
                                if (originalIndex <= dataArr.length) {
                                    dataArr.splice(originalIndex, 0, draggedItem);
                                } else {
                                    dataArr.push(draggedItem);
                                }
                            }
                            renderTituloModule();
                        } else {
                            addEmptyGroup(blockName, isCrudo ? (originalIsMezcla ? 'mix' : 'pure') : 'htr');
                            renderBalanceModule();
                        }
                    }
                }, 100);
            } else {
                // Forzar recÃ¡lculo del mÃ³dulo
                if (tituloModuleActive) {
                    renderTituloModule();
                } else {
                    renderBalanceModule();
                }
            }
        } catch(err) {
            console.error('Error en handleDrop:', err);
        } finally {
            draggedId = null;
        }
    }

    function handleDropToOtros(e) {
        e.preventDefault();
        const tbody = e.target.closest('tbody');
        if (tbody) {
            const tr = tbody.closest('div.table-wrap');
            // Remove any visual hints
            const possible = document.querySelectorAll('tr.drag-over');
            possible.forEach(x => x.classList.remove('drag-over'));
        }

        if (!draggedId) return;
        // Try usar draggedItemRef primero (mÃ¡s robusto)
        let item = null;
        if (draggedItemRef && draggedItemRef.item) {
            item = draggedItemRef.item;
        }
        try {
            const isCrudo = (GLOBAL_DATA.currentTab === 'crudo');
            const dataArr = isCrudo ? GLOBAL_DATA.nuevo : GLOBAL_DATA.htr;
            if (!item) {
                // Fallback: buscar en arrays
                const draggedIndex = dataArr.findIndex(i => i._id === draggedId);
                if (draggedIndex === -1) { draggedId = null; draggedItemRef = null; return; }
                item = dataArr[draggedIndex];
            }
            // Guardar el grupo original para detectar si quedarÃ¡ vacÃ­o
            const originalGroup = item.group;
            const originalIsMezcla = !!item.isMezcla;

            // Reasignar grupo a OTROS y marcar como no mezcla
            item.group = 'OTROS';
            item.isMezcla = false;
            // Si estamos en el mÃ³dulo TÃ­tulo, asignar tambiÃ©n al bloque de tÃ­tulo OTROS
            try {
                const tituloModuleActive = document.getElementById('mod-titulo').classList.contains('active');
                if (tituloModuleActive) item.titulo = 'OTROS';
            } catch(e) {}

            // Mover al final del array
            // Solo hacer splice si encontramos el Ã­ndice
            const idxInArr = dataArr.findIndex(i => i._id === draggedId);
            if (idxInArr !== -1) dataArr.splice(idxInArr, 1);
            dataArr.push(item);

            // Detectar si el grupo original quedÃ³ vacÃ­o
            const remainingInGroup = dataArr.filter(r => (r.group || '').toString() === originalGroup);
            if (remainingInGroup.length === 0 && originalGroup && originalGroup.toUpperCase() !== 'OTROS') {
                // Grupo quedÃ³ vacÃ­o, mostrar diÃ¡logo SIN renderizar aÃºn
                setTimeout(() => {
                    if (confirm(`Â¿Deseas eliminar el bloque vacÃ­o "${originalGroup}" o mantenerlo vacÃ­o?\n\nCancelar = Mantener vacÃ­o\nAceptar = Eliminar`)) {
                        // Usuario aceptÃ³: eliminar el grupo
                        deleteGroup(originalGroup, isCrudo ? (originalIsMezcla ? 'mix' : 'pure') : 'htr');
                    } else {
                        // Usuario cancelÃ³: preservar el bloque vacÃ­o y renderizar
                        addEmptyGroup(originalGroup, isCrudo ? (originalIsMezcla ? 'mix' : 'pure') : 'htr');
                        renderBalanceModule();
                    }
                }, 100);
            } else {
                renderBalanceModule();
            }
        } catch(err) {
            console.error('Error en handleDropToOtros:', err);
        } finally {
            draggedId = null;
            draggedItemRef = null;
        }
    }

    // Permitir soltar sobre un bloque vacÃ­o para mover el item al grupo target
    function handleDropToGroup(e, targetGroupName, isMez, isHtr) {
        e.preventDefault();
        try {
            const isCrudo = (GLOBAL_DATA.currentTab === 'crudo');
            // Preferir isHtr flag si se pasa
            const useHtr = (typeof isHtr !== 'undefined') ? Boolean(isHtr) : !isCrudo;
            const dataArr = useHtr ? GLOBAL_DATA.htr : GLOBAL_DATA.nuevo;
            if (!draggedId) return;
            const idx = dataArr.findIndex(i => i._id === draggedId);
            if (idx === -1) {
                // If not found in the active array, search the other array
                const other = useHtr ? GLOBAL_DATA.nuevo : GLOBAL_DATA.htr;
                const idx2 = other.findIndex(i => i._id === draggedId);
                if (idx2 === -1) { draggedId = null; return; }
                // Move item between arrays
                const item = other.splice(idx2,1)[0];
                item.group = targetGroupName;
                item.isMezcla = Boolean(isMez);
                if (!item.titulo || item.titulo === 'OTROS') {
                    const m = (item.hilado || '').toString().match(/^\s*(\d+\/\d+)/);
                    item.titulo = m ? m[1] : item.titulo || 'SIN TITULO';
                }
                dataArr.push(item);
            } else {
                const item = dataArr[idx];
                item.group = targetGroupName;
                item.isMezcla = Boolean(isMez);
                if (!item.titulo || item.titulo === 'OTROS') {
                    const m = (item.hilado || '').toString().match(/^\s*(\d+\/\d+)/);
                    item.titulo = m ? m[1] : item.titulo || 'SIN TITULO';
                }
            }

            // Re-render and recalc
            identifyIncludedRows(GLOBAL_DATA.nuevo, GLOBAL_DATA.excelTotals.crudo);
            identifyIncludedRows(GLOBAL_DATA.htr, GLOBAL_DATA.excelTotals.htr);
            sortDataArray(GLOBAL_DATA.nuevo);
            sortDataArray(GLOBAL_DATA.htr);
            renderPCPModule();
            renderBalanceModule();
            renderTituloModule();
            updateFooterTotals();
        } catch(e) { console.error('handleDropToGroup error', e); }
        finally { draggedId = null; }
    }

    // Permitir soltar sobre un BLOQUE de TÃTULO para cambiar el `titulo` del Ã­tem
    function handleDropToTitle(e, targetTitle, isMez, isHtr) {
        e.preventDefault();
        try {
            console.log('handleDropToTitle invoked', { draggedId, targetTitle, isMez, isHtr });
            const isCrudo = (GLOBAL_DATA.currentTab === 'crudo');
            // Preferir isHtr flag si se pasa
            const useHtr = (typeof isHtr !== 'undefined') ? Boolean(isHtr) : !isCrudo;
            const dataArr = useHtr ? GLOBAL_DATA.htr : GLOBAL_DATA.nuevo;
            if (!draggedId) { console.warn('handleDropToTitle: no draggedId'); return; }
            // Buscar el Ã­tem en todos los arrays posibles (incluyendo originales)
            const searchArrays = [GLOBAL_DATA.nuevo, GLOBAL_DATA.nuevoOriginal, GLOBAL_DATA.htr, GLOBAL_DATA.htrOriginal];
            let found = false;
            for (let arr of searchArrays) {
                if (!arr) continue;
                const i = arr.findIndex(x => x && x._id === draggedId);
                if (i !== -1) {
                    const item = arr.splice(i,1)[0];
                    // Asignar nuevo tÃ­tulo
                    item.titulo = targetTitle;
                    // AÃ±adir al array objetivo (dataArr puede ser el mismo u otro)
                    dataArr.push(item);
                    found = true;
                    console.log('handleDropToTitle: moved item', item._id, 'to title', targetTitle);
                    break;
                }
            }
            // Si no se encontrÃ³ en los originales, buscar dentro del dataArr y actualizar titulo en sitio
            if (!found) {
                const idx = dataArr.findIndex(i => i._id === draggedId);
                if (idx !== -1) {
                    dataArr[idx].titulo = targetTitle;
                    found = true;
                }
            }
            if (!found) { console.warn('handleDropToTitle: item not found in any array', draggedId); }
            // If we have a direct reference from dragstart, prefer moving that object
            if (draggedItemRef && draggedItemRef.item) {
                try {
                    // remove from its original array if present
                    const origArr = draggedItemRef.fromArray;
                    const idxOrig = origArr.findIndex(x => x && x._id === draggedItemRef.item._id);
                    if (idxOrig !== -1) origArr.splice(idxOrig,1);
                    // assign titulo and push to target
                    draggedItemRef.item.titulo = targetTitle;
                    dataArr.push(draggedItemRef.item);
                    console.log('handleDropToTitle: moved via ref', draggedItemRef.item._id);
                    found = true;
                } catch(e) { console.error('handleDropToTitle ref-move error', e); }
            }

            // Re-render and recalc
            identifyIncludedRows(GLOBAL_DATA.nuevo, GLOBAL_DATA.excelTotals.crudo);
            identifyIncludedRows(GLOBAL_DATA.htr, GLOBAL_DATA.excelTotals.htr);
            sortDataArray(GLOBAL_DATA.nuevo);
            sortDataArray(GLOBAL_DATA.htr);
            renderPCPModule();
            renderBalanceModule();
            renderTituloModule();
            updateFooterTotals();
        } catch(e) { console.error('handleDropToTitle error', e); }
        finally { draggedId = null; }
    }

    // Confirmar eliminaciÃ³n de grupo vacÃ­o
    function confirmDeleteGroup(groupName, tableType) {
        if (confirm(`Â¿Deseas eliminar el bloque vacÃ­o "${groupName}"?`)) {
            deleteGroup(groupName, tableType);
        }
    }

    // Eliminar grupo vacÃ­o
    function deleteGroup(groupName, tableType) {
        // Buscar y eliminar del array correspondiente
        if (tableType === 'pure') {
            GLOBAL_DATA.nuevo = GLOBAL_DATA.nuevo.filter(r => (r.group || '').toString() !== groupName);
        } else if (tableType === 'mix') {
            GLOBAL_DATA.nuevo = GLOBAL_DATA.nuevo.filter(r => (r.group || '').toString() !== groupName);
        } else if (tableType === 'htr') {
            GLOBAL_DATA.htr = GLOBAL_DATA.htr.filter(r => (r.group || '').toString() !== groupName);
        }
        // Remover cualquier placeholder asociado
        if (GLOBAL_DATA.emptyGroups) {
            if (tableType === 'htr') {
                GLOBAL_DATA.emptyGroups.htr = (GLOBAL_DATA.emptyGroups.htr || []).filter(p => (p.name || '') !== groupName);
            } else {
                GLOBAL_DATA.emptyGroups.crudo = (GLOBAL_DATA.emptyGroups.crudo || []).filter(p => (p.name || '') !== groupName);
            }
        }
        // Re-renderizar
        renderBalanceModule();
    }

    // AÃ±adir placeholder para un bloque vacÃ­o (evita duplicados)
    function addEmptyGroup(groupName, tableType) {
        if (!groupName) return;
        if (!GLOBAL_DATA.emptyGroups) GLOBAL_DATA.emptyGroups = { crudo: [], htr: [] };
        if (tableType === 'htr') {
            if (!GLOBAL_DATA.emptyGroups.htr.some(p => p.name === groupName)) {
                GLOBAL_DATA.emptyGroups.htr.push({ name: groupName, isMezcla: false });
            }
        } else {
            const isMez = (tableType === 'mix');
            if (!GLOBAL_DATA.emptyGroups.crudo.some(p => p.name === groupName)) {
                GLOBAL_DATA.emptyGroups.crudo.push({ name: groupName, isMezcla: isMez });
            }
        }
    }
