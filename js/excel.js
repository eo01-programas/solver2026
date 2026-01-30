// excel.js - Exportacion de datos a Excel con formato profesional y FORMULAS
// Genera un archivo Excel con 4 hojas: DATA_CRUDO_HTR, MAT_CRUDO_MEZCLA_HTR, TITULO_CRUDO_MEZCLA_HTR, RESUMEN_AGRUP
// VERSION 2: Corrige la extraccion de mezclas (componentes con %) y QQ por tipo (ORGANICO/TANGUIS)

async function exportToExcel() {
    if (!GLOBAL_DATA || (!GLOBAL_DATA.nuevo.length && !GLOBAL_DATA.htr.length)) {
        alert('No hay datos para exportar. Cargue un archivo primero.');
        return;
    }

    try {
        const workbook = new ExcelJS.Workbook();
        workbook.creator = 'Hilanderia System';
        workbook.created = new Date();

        // Estilos comunes
        const styles = getExcelStyles();

        // =====================================================
        // HOJA 1: DATA_CRUDO_HTR - Datos de PCP Fuente
        // =====================================================
        const sheet1 = workbook.addWorksheet('DATA_CRUDO_HTR');
        createSheet1_DataCrudoHtr(sheet1, styles);

        // =====================================================
        // HOJA 2: MAT_CRUDO_MEZCLA_HTR - Agrup Material
        // =====================================================
        const sheet2 = workbook.addWorksheet('MAT_CRUDO_MEZCLA_HTR');
        createSheet2_MatCrudoMezclaHtr(sheet2, styles);

        // =====================================================
        // HOJA 3: TITULO_CRUDO_MEZCLA_HTR - Agrup Titulo
        // =====================================================
        const sheet3 = workbook.addWorksheet('TITULO_CRUDO_MEZCLA_HTR');
        createSheet3_TituloCrudoMezclaHtr(sheet3, styles);

        // =====================================================
        // HOJA 4: RESUMEN_AGRUP - Resumen Totales
        // =====================================================
        const sheet4 = workbook.addWorksheet('RESUMEN_AGRUP');
        createSheet4_ResumenAgrup(sheet4, styles);

        // Generar y descargar el archivo
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const url = window.URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        const fecha = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        a.download = `Balance_Hilanderia_${fecha}.xlsx`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);

        console.log('Excel exportado exitosamente');
    } catch (error) {
        console.error('Error al exportar Excel:', error);
        alert('Error al generar el archivo Excel: ' + error.message);
    }
}

// =====================================================
// ESTILOS COMUNES
// =====================================================
function getExcelStyles() {
    return {
        header: {
            font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1e3a5f' } },
            alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
            border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            }
        },
        subHeader: {
            font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF3b82f6' } },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            }
        },
        title: {
            font: { bold: true, size: 14, color: { argb: 'FF1e3a5f' } },
            alignment: { horizontal: 'left', vertical: 'middle' }
        },
        cell: {
            border: {
                top: { style: 'thin', color: { argb: 'FFd1d5db' } },
                left: { style: 'thin', color: { argb: 'FFd1d5db' } },
                bottom: { style: 'thin', color: { argb: 'FFd1d5db' } },
                right: { style: 'thin', color: { argb: 'FFd1d5db' } }
            },
            alignment: { vertical: 'middle' }
        },
        totalRow: {
            font: { bold: true },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf1f5f9' } },
            border: {
                top: { style: 'medium', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'medium', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            }
        },
        summaryLabel: {
            font: { bold: true, size: 11 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFe2e8f0' } },
            alignment: { horizontal: 'right', vertical: 'middle' },
            border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            }
        },
        summaryValue: {
            font: { bold: true, size: 11, color: { argb: 'FF2563eb' } },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf0f9ff' } },
            alignment: { horizontal: 'center', vertical: 'middle' },
            border: {
                top: { style: 'thin', color: { argb: 'FF000000' } },
                left: { style: 'thin', color: { argb: 'FF000000' } },
                bottom: { style: 'thin', color: { argb: 'FF000000' } },
                right: { style: 'thin', color: { argb: 'FF000000' } }
            }
        },
        compHeader0: { font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF059669' } }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }, border: { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } } },
        compHeader1: { font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdc2626' } }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }, border: { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } } },
        compHeader2: { font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 10 }, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFca8a04' } }, alignment: { horizontal: 'center', vertical: 'middle', wrapText: true }, border: { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } } }
    };
}

// =====================================================
// HOJA 1: DATA_CRUDO_HTR
// =====================================================
function createSheet1_DataCrudoHtr(sheet, styles) {
    sheet.columns = [
        { width: 10 }, { width: 15 }, { width: 10 }, { width: 8 },
        { width: 45 }, { width: 15 }, { width: 12 }, { width: 25 }
    ];

    let row = 1;
    const crudoData = GLOBAL_DATA.nuevoOriginal && GLOBAL_DATA.nuevoOriginal.length ? GLOBAL_DATA.nuevoOriginal : GLOBAL_DATA.nuevo;
    
    // CRUDO / NUEVO
    sheet.getCell(`A${row}`).value = 'CRUDO / NUEVO';
    sheet.getCell(`A${row}`).style = styles.title;
    sheet.mergeCells(`A${row}:H${row}`);
    row++;

    const headers = ['ORDEN', 'CLIENTE', 'TEMP', 'OP', 'HILADO', 'COLOR', 'KG SOL.', 'OBSERVACIONES'];
    const headerRow1 = sheet.getRow(row);
    headers.forEach((h, i) => { headerRow1.getCell(i + 1).value = h; headerRow1.getCell(i + 1).style = styles.header; });
    headerRow1.height = 25;
    row++;

    const crudoStartRow = row;
    crudoData.forEach(r => {
        const dataRow = sheet.getRow(row);
        dataRow.values = [r.orden || '-', r.cliente || '-', r.temporada || '-', r.op || '-', r.hilado || '-', r.colorText || '-', r.kg || 0, r.obs || '-'];
        for (let i = 1; i <= 8; i++) { dataRow.getCell(i).style = { ...styles.cell }; if (i === 7) dataRow.getCell(i).numFmt = '#,##0.00'; }
        row++;
    });
    const crudoEndRow = row - 1;

    const totalCrudoRow = sheet.getRow(row);
    totalCrudoRow.getCell(1).value = 'TOTAL KG CRUDO:';
    sheet.mergeCells(`A${row}:F${row}`);
    totalCrudoRow.getCell(1).style = { ...styles.totalRow, alignment: { horizontal: 'right' } };
    totalCrudoRow.getCell(7).value = { formula: `SUM(G${crudoStartRow}:G${crudoEndRow})` };
    totalCrudoRow.getCell(7).style = { ...styles.totalRow, numFmt: '#,##0.00' };
    const crudoTotalRowNum = row;
    row += 2;

    // HTR
    const htrData = GLOBAL_DATA.htrOriginal && GLOBAL_DATA.htrOriginal.length ? GLOBAL_DATA.htrOriginal : GLOBAL_DATA.htr;
    
    sheet.getCell(`A${row}`).value = 'HTR';
    sheet.getCell(`A${row}`).style = styles.title;
    sheet.mergeCells(`A${row}:H${row}`);
    row++;

    const headerRow2 = sheet.getRow(row);
    headers.forEach((h, i) => { headerRow2.getCell(i + 1).value = h; headerRow2.getCell(i + 1).style = styles.header; });
    headerRow2.height = 25;
    row++;

    const htrStartRow = row;
    htrData.forEach(r => {
        const dataRow = sheet.getRow(row);
        dataRow.values = [r.orden || '-', r.cliente || '-', r.temporada || '-', r.op || '-', r.hilado || '-', r.colorText || '-', r.kg || 0, r.obs || '-'];
        for (let i = 1; i <= 8; i++) { dataRow.getCell(i).style = { ...styles.cell }; if (i === 7) dataRow.getCell(i).numFmt = '#,##0.00'; }
        row++;
    });
    const htrEndRow = row - 1;

    const totalHtrRow = sheet.getRow(row);
    totalHtrRow.getCell(1).value = 'TOTAL KG HTR:';
    sheet.mergeCells(`A${row}:F${row}`);
    totalHtrRow.getCell(1).style = { ...styles.totalRow, alignment: { horizontal: 'right' } };
    totalHtrRow.getCell(7).value = { formula: `SUM(G${htrStartRow}:G${htrEndRow})` };
    totalHtrRow.getCell(7).style = { ...styles.totalRow, numFmt: '#,##0.00' };
    const htrTotalRowNum = row;
    row += 3;

    // RESUMEN
    sheet.getCell(`A${row}`).value = 'RESUMEN GENERAL';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 12 } };
    sheet.mergeCells(`A${row}:C${row}`);
    row++;

    sheet.getCell(`A${row}`).value = 'TOTAL KG CRUDO';
    sheet.getCell(`A${row}`).style = styles.summaryLabel;
    sheet.mergeCells(`A${row}:B${row}`);
    sheet.getCell(`C${row}`).value = { formula: `G${crudoTotalRowNum}` };
    sheet.getCell(`C${row}`).style = { ...styles.summaryValue, numFmt: '#,##0.00' };
    row++;

    sheet.getCell(`A${row}`).value = 'TOTAL KG HTR';
    sheet.getCell(`A${row}`).style = styles.summaryLabel;
    sheet.mergeCells(`A${row}:B${row}`);
    sheet.getCell(`C${row}`).value = { formula: `G${htrTotalRowNum}` };
    sheet.getCell(`C${row}`).style = { ...styles.summaryValue, numFmt: '#,##0.00' };
    row++;

    sheet.getCell(`A${row}`).value = 'TOTAL GENERAL';
    sheet.getCell(`A${row}`).style = { ...styles.summaryLabel, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1e3a5f' } }, font: { bold: true, color: { argb: 'FFFFFFFF' } } };
    sheet.mergeCells(`A${row}:B${row}`);
    sheet.getCell(`C${row}`).value = { formula: `G${crudoTotalRowNum}+G${htrTotalRowNum}` };
    sheet.getCell(`C${row}`).style = { ...styles.summaryValue, numFmt: '#,##0.00', font: { bold: true, size: 12, color: { argb: 'FF16a34a' } } };
    row++;

    const neVal = GLOBAL_DATA.excelTotals.ne || computeWeightedNeGlobal();
    sheet.getCell(`A${row}`).value = 'NE PROMEDIO';
    sheet.getCell(`A${row}`).style = styles.summaryLabel;
    sheet.mergeCells(`A${row}:B${row}`);
    sheet.getCell(`C${row}`).value = neVal || '-';
    sheet.getCell(`C${row}`).style = styles.summaryValue;
}

// =====================================================
// HOJA 2: MAT_CRUDO_MEZCLA_HTR - CON COMPONENTES DE MEZCLA
// =====================================================
function createSheet2_MatCrudoMezclaHtr(sheet, styles) {
    let row = 1;

    const activeCrudo = GLOBAL_DATA.nuevo.some(r => r.highlight) ? GLOBAL_DATA.nuevo.filter(r => r.highlight) : GLOBAL_DATA.nuevo;
    const activeHtr = GLOBAL_DATA.htr.some(r => r.highlight) ? GLOBAL_DATA.htr.filter(r => r.highlight) : GLOBAL_DATA.htr;

    // SECCION CRUDOS (no mezcla)
    const pureItems = activeCrudo.filter(r => !r.isMezcla);
    const pureGroups = groupByMaterialExcel(pureItems);

    sheet.getCell(`A${row}`).value = 'CRUDOS';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 14, color: { argb: 'FF16a34a' } } };
    row++;

    row = writePureGroupTable(sheet, pureGroups, row, styles, false);
    row += 2;

    // SECCION MEZCLAS - CON COMPONENTES
    const mixItems = activeCrudo.filter(r => r.isMezcla);
    const mixGroups = groupByMaterialExcel(mixItems);

    sheet.getCell(`A${row}`).value = 'MEZCLAS';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 14, color: { argb: 'FFca8a04' } } };
    row++;

    row = writeMixGroupTableWithComponents(sheet, mixGroups, row, styles, false);
    row += 2;

    // SECCION HTR - PUEDE TENER MEZCLAS
    const htrPure = activeHtr.filter(r => !r.isMezcla);
    const htrMix = activeHtr.filter(r => r.isMezcla);
    const htrPureGroups = groupByMaterialExcel(htrPure);
    const htrMixGroups = groupByMaterialExcel(htrMix);

    sheet.getCell(`A${row}`).value = 'HTR';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 14, color: { argb: 'FFdc2626' } } };
    row++;

    if (htrPureGroups.length > 0) {
        row = writePureGroupTable(sheet, htrPureGroups, row, styles, true);
        row++;
    }
    
    if (htrMixGroups.length > 0) {
        sheet.getCell(`A${row}`).value = 'HTR - MEZCLAS';
        sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 12, color: { argb: 'FFca8a04' } } };
        row++;
        row = writeMixGroupTableWithComponents(sheet, htrMixGroups, row, styles, true);
    }
}

// =====================================================
// TABLA PARA GRUPOS PUROS (NO MEZCLA)
// =====================================================
function writePureGroupTable(sheet, groups, startRow, styles, isHtr) {
    let row = startRow;

    // Configurar columnas para puros
    sheet.getColumn(1).width = 10;  // ORDEN
    sheet.getColumn(2).width = 12;  // CLIENTE
    sheet.getColumn(3).width = 8;   // TEMP
    sheet.getColumn(4).width = 8;   // RSV
    sheet.getColumn(5).width = 8;   // OP
    sheet.getColumn(6).width = 40;  // HILADO
    sheet.getColumn(7).width = 12;  // COLOR
    sheet.getColumn(8).width = 8;   // NE
    sheet.getColumn(9).width = 12;  // KG SOL
    sheet.getColumn(10).width = 8;  // FACTOR
    sheet.getColumn(11).width = 12; // KG REQ
    sheet.getColumn(12).width = 10; // QQ REQ
    sheet.getColumn(13).width = 12; // QQ ORG 80%
    sheet.getColumn(14).width = 12; // QQ TAN 20%

    groups.forEach(g => {
        if (g.name.toUpperCase() === 'OTROS' && g.items.length === 0) return;

        // Detectar si el grupo tiene COP ORGANICO de cliente LLL (OCS)
        const groupHasCopOrgOcs = g.items.some(r => {
            const cert = getClientCertExcel(r.cliente);
            return cert === 'OCS' && /COP\s*(?:ORGANICO|ORG|ORGANIC)/i.test(r.hilado || '');
        });

        const factor = getGroupFactorExcel(isHtr, g.name);
        const groupColor = isHtr ? 'FFdc2626' : 'FF059669';

        // Header del grupo
        sheet.getCell(`A${row}`).value = `MATERIAL: ${g.name} (Factor: ${factor.toFixed(2)})`;
        sheet.getCell(`A${row}`).style = {
            font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: groupColor } },
            alignment: { horizontal: 'left', vertical: 'middle' }
        };
        const mergeEnd = groupHasCopOrgOcs ? 'N' : 'L';
        sheet.mergeCells(`A${row}:${mergeEnd}${row}`);
        row++;

        // Headers de columnas
        const headers = ['ORDEN', 'CLIENTE', 'TEMP', 'RSV', 'OP', 'HILADO', 'COLOR', 'NE', 'KG SOL', 'FACTOR', 'KG REQ', 'QQ REQ'];
        if (groupHasCopOrgOcs) {
            headers.push('QQ ORG 80%', 'QQ TAN 20%');
        }
        const headerRow = sheet.getRow(row);
        headers.forEach((h, i) => {
            headerRow.getCell(i + 1).value = h;
            headerRow.getCell(i + 1).style = styles.header;
        });
        headerRow.height = 22;
        row++;

        const dataStartRow = row;

        // Datos
        g.items.forEach(r => {
            const kgSol = r.kg || 0;
            const neVal = getNeFromItemExcel(r);
            const isAlgodon = isAlgodonTextExcel(r.hilado || r.group || '');
            
            // Detectar si es COP ORGANICO OCS
            const cert = getClientCertExcel(r.cliente);
            const isCopOrgOcs = cert === 'OCS' && /COP\s*(?:ORGANICO|ORG|ORGANIC)/i.test(r.hilado || '');

            const dataRow = sheet.getRow(row);
            dataRow.getCell(1).value = r.orden || '-';
            dataRow.getCell(2).value = r.cliente || '-';
            dataRow.getCell(3).value = r.temporada || '-';
            dataRow.getCell(4).value = r.rsv || '';
            dataRow.getCell(5).value = r.op || '-';
            dataRow.getCell(6).value = r.hilado || '-';
            dataRow.getCell(7).value = r.colorText || '-';
            dataRow.getCell(8).value = neVal ? Math.round(neVal) : '-';
            dataRow.getCell(9).value = kgSol;
            dataRow.getCell(10).value = factor;

            // KG REQ = KG SOL / FACTOR
            dataRow.getCell(11).value = { formula: `ROUND(I${row}/J${row},0)` };

            // QQ REQ = KG REQ / 46 (solo si es algodon)
            if (isAlgodon) {
                dataRow.getCell(12).value = { formula: `ROUND(K${row}/46,0)` };
            } else {
                dataRow.getCell(12).value = '-';
            }

            // QQ ORG 80% y QQ TAN 20% (solo si groupHasCopOrgOcs)
            if (groupHasCopOrgOcs) {
                if (isCopOrgOcs) {
                    dataRow.getCell(13).value = { formula: `ROUND(L${row}*0.8,0)` };
                    dataRow.getCell(14).value = { formula: `ROUND(L${row}*0.2,0)` };
                } else {
                    dataRow.getCell(13).value = '';
                    dataRow.getCell(14).value = '';
                }
            }

            // Estilos
            const colCount = groupHasCopOrgOcs ? 14 : 12;
            for (let i = 1; i <= colCount; i++) {
                dataRow.getCell(i).style = { ...styles.cell };
                if (i === 9) dataRow.getCell(i).numFmt = '#,##0.00';
                if (i === 10) dataRow.getCell(i).numFmt = '0.00';
                if (i >= 11) dataRow.getCell(i).numFmt = '#,##0';
            }
            dataRow.getCell(11).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf8fafc' } }, numFmt: '#,##0' };
            dataRow.getCell(12).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf0f9ff' } }, numFmt: '#,##0' };
            if (groupHasCopOrgOcs) {
                dataRow.getCell(13).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFdcfce7' } }, numFmt: '#,##0' };
                dataRow.getCell(14).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } }, numFmt: '#,##0' };
            }
            row++;
        });

        const dataEndRow = row - 1;

        // Fila de totales
        if (g.items.length > 0) {
            const totalRow = sheet.getRow(row);
            totalRow.getCell(1).value = `TOTAL ${g.name}:`;
            sheet.mergeCells(`A${row}:H${row}`);
            totalRow.getCell(1).style = { ...styles.totalRow, alignment: { horizontal: 'right' } };

            totalRow.getCell(9).value = { formula: `SUM(I${dataStartRow}:I${dataEndRow})` };
            totalRow.getCell(9).style = { ...styles.totalRow, numFmt: '#,##0.00' };
            totalRow.getCell(10).style = styles.totalRow;
            totalRow.getCell(11).value = { formula: `SUM(K${dataStartRow}:K${dataEndRow})` };
            totalRow.getCell(11).style = { ...styles.totalRow, numFmt: '#,##0' };
            totalRow.getCell(12).value = { formula: `SUM(L${dataStartRow}:L${dataEndRow})` };
            totalRow.getCell(12).style = { ...styles.totalRow, numFmt: '#,##0' };

            if (groupHasCopOrgOcs) {
                totalRow.getCell(13).value = { formula: `SUM(M${dataStartRow}:M${dataEndRow})` };
                totalRow.getCell(13).style = { ...styles.totalRow, numFmt: '#,##0' };
                totalRow.getCell(14).value = { formula: `SUM(N${dataStartRow}:N${dataEndRow})` };
                totalRow.getCell(14).style = { ...styles.totalRow, numFmt: '#,##0' };
            }
            row++;
        }
        row++;
    });

    return row;
}

// =====================================================
// TABLA PARA GRUPOS DE MEZCLA - CON COMPONENTES
// Cada componente tiene: P% (participacion), MERMA, KG REQ (formula), QQ (formula si algodon)
// =====================================================
function writeMixGroupTableWithComponents(sheet, groups, startRow, styles, isHtr) {
    let row = startRow;

    groups.forEach(g => {
        if (g.name.toUpperCase() === 'OTROS' && g.items.length === 0) return;
        if (g.items.length === 0) return;

        // Obtener componentes del primer hilado para determinar headers
        const firstHilado = g.items[0].hilado || '';
        const sampleComps = isHtr 
            ? parseMixCompositionHtr(firstHilado, 1000, { groupName: g.name })
            : parseMixComposition(firstHilado, 1000, { groupName: g.name });

        const groupColor = isHtr ? 'FFdc2626' : 'FFca8a04';

        // Header del grupo
        sheet.getCell(`A${row}`).value = `MEZCLA: ${g.name}`;
        sheet.getCell(`A${row}`).style = {
            font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: groupColor } },
            alignment: { horizontal: 'left', vertical: 'middle' }
        };
        row++;

        // Headers base
        const baseHeaders = ['ORDEN', 'CLIENTE', 'TEMP', 'RSV', 'OP', 'HILADO', 'COLOR', 'NE', 'KG SOL'];
        let colIdx = 1;

        // Configurar anchos base
        sheet.getColumn(1).width = 10;
        sheet.getColumn(2).width = 12;
        sheet.getColumn(3).width = 8;
        sheet.getColumn(4).width = 8;
        sheet.getColumn(5).width = 8;
        sheet.getColumn(6).width = 40;
        sheet.getColumn(7).width = 12;
        sheet.getColumn(8).width = 8;
        sheet.getColumn(9).width = 12;

        const headerRow = sheet.getRow(row);
        baseHeaders.forEach((h, i) => {
            headerRow.getCell(i + 1).value = h;
            headerRow.getCell(i + 1).style = styles.header;
        });
        colIdx = baseHeaders.length + 1;

        // Headers de componentes: cada uno tiene P%, MERMA, KG REQ, y QQ (si algodon)
        const compColMap = []; // [{name, pctCol, mermaCol, kgCol, qqCol}]

        sampleComps.forEach((c, idx) => {
            const compName = c.name;
            const isAlgodonComp = isAlgodonMixComponentExcel(compName);
            const compStyle = styles[`compHeader${idx % 3}`] || styles.compHeader0;
            
            // Columna nombre componente (header spanning)
            const compNameHeader = compName;
            
            // Columna P% (participacion)
            headerRow.getCell(colIdx).value = `${compName}\nP%`;
            headerRow.getCell(colIdx).style = compStyle;
            sheet.getColumn(colIdx).width = 8;
            const pctCol = colIdx;
            colIdx++;

            // Columna MERMA
            headerRow.getCell(colIdx).value = `${compName}\nMERMA`;
            headerRow.getCell(colIdx).style = compStyle;
            sheet.getColumn(colIdx).width = 8;
            const mermaCol = colIdx;
            colIdx++;

            // Columna KG REQ (formula: KG SOL * P% / (1 - MERMA))
            headerRow.getCell(colIdx).value = `${compName}\nKG REQ`;
            headerRow.getCell(colIdx).style = compStyle;
            sheet.getColumn(colIdx).width = 12;
            const kgCol = colIdx;
            colIdx++;

            // Columna QQ del componente (solo si es algodon)
            let qqCol = null;
            if (isAlgodonComp) {
                headerRow.getCell(colIdx).value = `${compName}\nQQ`;
                headerRow.getCell(colIdx).style = compStyle;
                sheet.getColumn(colIdx).width = 10;
                qqCol = colIdx;
                colIdx++;
            }

            compColMap.push({ name: compName, pctCol, mermaCol, kgCol, qqCol, isAlgodon: isAlgodonComp });
        });

        headerRow.height = 35;
        row++;

        const dataStartRow = row;

        // Datos
        g.items.forEach(r => {
            const kgSol = r.kg || 0;
            const neVal = getNeFromItemExcel(r);
            const comps = isHtr 
                ? parseMixCompositionHtr(r.hilado || '', kgSol, { groupName: g.name })
                : parseMixComposition(r.hilado || '', kgSol, { groupName: g.name });

            const dataRow = sheet.getRow(row);
            dataRow.getCell(1).value = r.orden || '-';
            dataRow.getCell(2).value = r.cliente || '-';
            dataRow.getCell(3).value = r.temporada || '-';
            dataRow.getCell(4).value = r.rsv || '';
            dataRow.getCell(5).value = r.op || '-';
            dataRow.getCell(6).value = r.hilado || '-';
            dataRow.getCell(7).value = r.colorText || '-';
            dataRow.getCell(8).value = neVal ? Math.round(neVal) : '-';
            dataRow.getCell(9).value = kgSol;

            // Estilos base
            for (let i = 1; i <= 9; i++) {
                dataRow.getCell(i).style = { ...styles.cell };
                if (i === 9) dataRow.getCell(i).numFmt = '#,##0.00';
            }

            // Datos de componentes con formulas
            comps.forEach((c, idx) => {
                if (idx < compColMap.length) {
                    const colInfo = compColMap[idx];
                    const pctColLetter = getColumnLetter(colInfo.pctCol);
                    const mermaColLetter = getColumnLetter(colInfo.mermaCol);
                    const kgColLetter = getColumnLetter(colInfo.kgCol);
                    
                    // P% como valor editable (decimal, ej: 0.75 para 75%)
                    const pctValue = c.pct || 0;
                    dataRow.getCell(colInfo.pctCol).value = pctValue;
                    dataRow.getCell(colInfo.pctCol).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFe0f2fe' } } };
                    dataRow.getCell(colInfo.pctCol).numFmt = '0%';

                    // MERMA como valor editable (decimal, ej: 0.40 para 40%)
                    const mermaValue = c.merma || 0;
                    dataRow.getCell(colInfo.mermaCol).value = mermaValue;
                    dataRow.getCell(colInfo.mermaCol).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFfef3c7' } } };
                    dataRow.getCell(colInfo.mermaCol).numFmt = '0%';

                    // KG REQ = ROUND(KG SOL * P% / (1 - MERMA), 0)
                    // Formula: =ROUND(I{row}*{pctCol}{row}/(1-{mermaCol}{row}),0)
                    dataRow.getCell(colInfo.kgCol).value = { 
                        formula: `ROUND(I${row}*${pctColLetter}${row}/(1-${mermaColLetter}${row}),0)` 
                    };
                    dataRow.getCell(colInfo.kgCol).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf8fafc' } }, numFmt: '#,##0' };
                    
                    // QQ = ROUND(KG REQ / 46, 0) - solo si es algodon
                    if (colInfo.qqCol && colInfo.isAlgodon) {
                        dataRow.getCell(colInfo.qqCol).value = { 
                            formula: `ROUND(${kgColLetter}${row}/46,0)` 
                        };
                        dataRow.getCell(colInfo.qqCol).style = { ...styles.cell, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFf0f9ff' } }, numFmt: '#,##0' };
                    }
                }
            });

            row++;
        });

        const dataEndRow = row - 1;

        // Fila de totales
        if (g.items.length > 0) {
            const totalRow = sheet.getRow(row);
            totalRow.getCell(1).value = `TOTAL ${g.name}:`;
            sheet.mergeCells(`A${row}:H${row}`);
            totalRow.getCell(1).style = { ...styles.totalRow, alignment: { horizontal: 'right' } };

            totalRow.getCell(9).value = { formula: `SUM(I${dataStartRow}:I${dataEndRow})` };
            totalRow.getCell(9).style = { ...styles.totalRow, numFmt: '#,##0.00' };

            // Totales de componentes (solo KG REQ y QQ)
            compColMap.forEach(colInfo => {
                // P% y MERMA no se suman en totales
                totalRow.getCell(colInfo.pctCol).value = '';
                totalRow.getCell(colInfo.pctCol).style = styles.totalRow;
                totalRow.getCell(colInfo.mermaCol).value = '';
                totalRow.getCell(colInfo.mermaCol).style = styles.totalRow;

                // Total KG REQ
                const kgColLetter = getColumnLetter(colInfo.kgCol);
                totalRow.getCell(colInfo.kgCol).value = { formula: `SUM(${kgColLetter}${dataStartRow}:${kgColLetter}${dataEndRow})` };
                totalRow.getCell(colInfo.kgCol).style = { ...styles.totalRow, numFmt: '#,##0' };

                // Total QQ
                if (colInfo.qqCol) {
                    const qqColLetter = getColumnLetter(colInfo.qqCol);
                    totalRow.getCell(colInfo.qqCol).value = { formula: `SUM(${qqColLetter}${dataStartRow}:${qqColLetter}${dataEndRow})` };
                    totalRow.getCell(colInfo.qqCol).style = { ...styles.totalRow, numFmt: '#,##0' };
                }
            });

            row++;
        }
        row += 2;
    });

    return row;
}

// =====================================================
// HOJA 3: TITULO_CRUDO_MEZCLA_HTR
// =====================================================
function createSheet3_TituloCrudoMezclaHtr(sheet, styles) {
    sheet.columns = [
        { width: 10 }, { width: 12 }, { width: 10 }, { width: 8 },
        { width: 8 }, { width: 40 }, { width: 8 }, { width: 10 }, { width: 12 }
    ];

    let row = 1;

    const activeCrudo = GLOBAL_DATA.nuevo.some(r => r.highlight) ? GLOBAL_DATA.nuevo.filter(r => r.highlight) : GLOBAL_DATA.nuevo;
    const activeHtr = GLOBAL_DATA.htr.some(r => r.highlight) ? GLOBAL_DATA.htr.filter(r => r.highlight) : GLOBAL_DATA.htr;

    // CRUDO
    const crudoPure = activeCrudo.filter(r => !r.isMezcla);
    const crudoTituloGroups = groupByTituloExcel(crudoPure);

    sheet.getCell(`A${row}`).value = 'CRUDO';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 14, color: { argb: 'FF16a34a' } } };
    row++;

    row = writeTituloGroupedTable(sheet, crudoTituloGroups, row, styles);
    row += 2;

    // MEZCLA
    const crudoMix = activeCrudo.filter(r => r.isMezcla);
    const mezclaGroups = groupByTituloExcel(crudoMix);

    sheet.getCell(`A${row}`).value = 'MEZCLA';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 14, color: { argb: 'FFca8a04' } } };
    row++;

    row = writeTituloGroupedTable(sheet, mezclaGroups, row, styles);
    row += 2;

    // HTR
    const htrGroups = groupByTituloExcel(activeHtr);

    sheet.getCell(`A${row}`).value = 'HTR';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 14, color: { argb: 'FFdc2626' } } };
    row++;

    row = writeTituloGroupedTable(sheet, htrGroups, row, styles);
}

function writeTituloGroupedTable(sheet, groups, startRow, styles) {
    let row = startRow;

    groups.forEach(group => {
        if (group.titulo.toUpperCase() === 'OTROS' && group.items.length === 0) return;

        sheet.getCell(`A${row}`).value = `TITULO: ${group.titulo}`;
        sheet.getCell(`A${row}`).style = {
            font: { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 },
            fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF6366f1' } },
            alignment: { horizontal: 'left', vertical: 'middle' }
        };
        sheet.mergeCells(`A${row}:I${row}`);
        row++;

        const headers = ['ORDEN', 'CLIENTE', 'TEMPORADA', 'RSV', 'OP', 'HILADO', 'NE', 'TIPO', 'KG SOL'];
        const headerRow = sheet.getRow(row);
        headers.forEach((h, i) => {
            headerRow.getCell(i + 1).value = h;
            headerRow.getCell(i + 1).style = styles.header;
        });
        headerRow.height = 22;
        row++;

        const dataStartRow = row;

        group.items.forEach(r => {
            const dataRow = sheet.getRow(row);
            dataRow.values = [
                r.orden || '-', r.cliente || '-', r.temporada || '-', r.rsv || '',
                r.op || '-', r.hilado || '-', getNeFromItemExcel(r) ? Math.round(getNeFromItemExcel(r)) : '-',
                r.tipo || '-', r.kg || 0
            ];
            for (let i = 1; i <= 9; i++) {
                dataRow.getCell(i).style = { ...styles.cell };
                if (i === 9) dataRow.getCell(i).numFmt = '#,##0.00';
            }
            row++;
        });

        const dataEndRow = row - 1;

        if (group.items.length > 0) {
            const totalRow = sheet.getRow(row);
            totalRow.getCell(1).value = `TOTAL ${group.titulo}:`;
            sheet.mergeCells(`A${row}:H${row}`);
            totalRow.getCell(1).style = { ...styles.totalRow, alignment: { horizontal: 'right' } };
            totalRow.getCell(9).value = { formula: `SUM(I${dataStartRow}:I${dataEndRow})` };
            totalRow.getCell(9).style = { ...styles.totalRow, numFmt: '#,##0.00' };
            row++;
        }
        row++;
    });

    return row;
}

// =====================================================
// HOJA 4: RESUMEN_AGRUP
// =====================================================
function createSheet4_ResumenAgrup(sheet, styles) {
    sheet.columns = [
        { width: 35 }, { width: 15 }, { width: 12 }, { width: 12 }, { width: 12 }
    ];

    let row = 1;

    const activeCrudo = GLOBAL_DATA.nuevo.some(r => r.highlight) ? GLOBAL_DATA.nuevo.filter(r => r.highlight) : GLOBAL_DATA.nuevo;
    const activeHtr = GLOBAL_DATA.htr.some(r => r.highlight) ? GLOBAL_DATA.htr.filter(r => r.highlight) : GLOBAL_DATA.htr;

    // Titulo
    sheet.getCell(`A${row}`).value = 'RESUMEN DE TOTALES - KG REQUERIDO';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 16 } };
    sheet.mergeCells(`A${row}:E${row}`);
    row += 2;

    // CRUDOS + MEZCLAS
    sheet.getCell(`A${row}`).value = 'CRUDOS + MEZCLAS';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 12, color: { argb: 'FF16a34a' } } };
    row++;

    const crudoResult = writeSummaryTableWithFormulas(sheet, activeCrudo, row, styles, false);
    row = crudoResult.endRow;
    const crudoTotalRef = crudoResult.totalCellRef;
    row += 2;

    // HTR
    sheet.getCell(`A${row}`).value = 'HTR';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 12, color: { argb: 'FFdc2626' } } };
    row++;

    const htrResult = writeSummaryTableWithFormulas(sheet, activeHtr, row, styles, true);
    row = htrResult.endRow;
    const htrTotalRef = htrResult.totalCellRef;
    row += 2;

    // COMBINADO
    sheet.getCell(`A${row}`).value = 'TOTAL COMBINADO (CRUDOS + MEZCLAS + HTR)';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 12, color: { argb: 'FF2563eb' } } };
    row++;

    const combinedResult = writeCombinedSummaryTable(sheet, activeCrudo, activeHtr, row, styles);
    row = combinedResult.endRow;
    row += 2;

    // RESUMEN FINAL
    sheet.getCell(`A${row}`).value = 'RESUMEN GENERAL';
    sheet.getCell(`A${row}`).style = { ...styles.title, font: { ...styles.title.font, size: 14 } };
    row++;

    sheet.getCell(`A${row}`).value = 'TOTAL KG CRUDO';
    sheet.getCell(`A${row}`).style = styles.summaryLabel;
    sheet.getCell(`B${row}`).value = { formula: crudoTotalRef || '0' };
    sheet.getCell(`B${row}`).style = { ...styles.summaryValue, numFmt: '#,##0' };
    const resCrudoRow = row;
    row++;

    sheet.getCell(`A${row}`).value = 'TOTAL KG HTR';
    sheet.getCell(`A${row}`).style = styles.summaryLabel;
    sheet.getCell(`B${row}`).value = { formula: htrTotalRef || '0' };
    sheet.getCell(`B${row}`).style = { ...styles.summaryValue, numFmt: '#,##0' };
    const resHtrRow = row;
    row++;

    sheet.getCell(`A${row}`).value = 'TOTAL GENERAL';
    sheet.getCell(`A${row}`).style = { ...styles.summaryLabel, fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1e3a5f' } }, font: { bold: true, color: { argb: 'FFFFFFFF' } } };
    sheet.getCell(`B${row}`).value = { formula: `B${resCrudoRow}+B${resHtrRow}` };
    sheet.getCell(`B${row}`).style = { ...styles.summaryValue, numFmt: '#,##0', font: { bold: true, size: 12, color: { argb: 'FF16a34a' } } };
    row++;

    const neVal = computeWeightedNeGlobal();
    sheet.getCell(`A${row}`).value = 'NE PROMEDIO';
    sheet.getCell(`A${row}`).style = styles.summaryLabel;
    sheet.getCell(`B${row}`).value = neVal ? Math.round(neVal) : '-';
    sheet.getCell(`B${row}`).style = styles.summaryValue;
}

// =====================================================
// RESUMEN CON FORMULA - SEPARA COMPONENTES DE MEZCLA
// =====================================================
function writeSummaryTableWithFormulas(sheet, data, startRow, styles, isHtr) {
    let row = startRow;

    const headers = ['MATERIAL', 'KG REQ', 'QQ (=B/46)', 'INGRESOS (=B/5800)', 'DIAS (=D*1.5)'];
    const headerRow = sheet.getRow(row);
    headers.forEach((h, i) => {
        headerRow.getCell(i + 1).value = h;
        headerRow.getCell(i + 1).style = styles.header;
    });
    headerRow.height = 22;
    row++;

    const dataStartRow = row;
    const materialsMap = new Map();

    // Procesar todos los items, separando componentes de mezcla
    data.forEach(r => {
        if (!r.isMezcla) {
            // No mezcla: usar factor del grupo
            const factor = getGroupFactorExcel(isHtr, r.group || '');
            const kgReq = Math.round((r.kg || 0) / factor);
            let matName = r.group || 'SIN GRUPO';
            
            // Agregar certificacion si tiene
            const hilado = (r.hilado || '').toUpperCase();
            if (/(?:ORGANICO|ORG|ORGANIC)/i.test(hilado)) {
                if (/\(OCS\)/.test(hilado)) matName += ' (OCS)';
                else if (/\(GOTS\)/.test(hilado)) matName += ' (GOTS)';
            }

            if (!materialsMap.has(matName)) materialsMap.set(matName, { kg: 0, isAlgodon: isAlgodonTextExcel(matName) });
            materialsMap.get(matName).kg += kgReq;
        } else {
            // Mezcla: dividir en componentes
            const comps = isHtr 
                ? parseMixCompositionHtr(r.hilado || '', r.kg || 0, { groupName: r.group })
                : parseMixComposition(r.hilado || '', r.kg || 0, { groupName: r.group });

            comps.forEach(c => {
                let matName = getBaseMaterialTokenExcel(c.name) || c.name;
                
                // Agregar certificacion si tiene
                const hilado = (r.hilado || '').toUpperCase();
                if (/(?:ORGANICO|ORG|ORGANIC)/i.test(c.name)) {
                    if (/\(OCS\)/.test(hilado)) matName += ' (OCS)';
                    else if (/\(GOTS\)/.test(hilado)) matName += ' (GOTS)';
                }

                if (!materialsMap.has(matName)) materialsMap.set(matName, { kg: 0, isAlgodon: isAlgodonMixComponentExcel(c.name) });
                materialsMap.get(matName).kg += c.kg || 0;
            });
        }
    });

    // Escribir filas
    materialsMap.forEach((matData, matName) => {
        const dataRow = sheet.getRow(row);
        dataRow.getCell(1).value = matName;
        dataRow.getCell(1).style = { ...styles.cell };

        dataRow.getCell(2).value = matData.kg;
        dataRow.getCell(2).style = { ...styles.cell, numFmt: '#,##0' };

        if (matData.isAlgodon) {
            dataRow.getCell(3).value = { formula: `ROUND(B${row}/46,0)` };
        } else {
            dataRow.getCell(3).value = '-';
        }
        dataRow.getCell(3).style = { ...styles.cell, numFmt: '#,##0' };

        dataRow.getCell(4).value = { formula: `B${row}/5800` };
        dataRow.getCell(4).style = { ...styles.cell, numFmt: '#,##0.00' };

        dataRow.getCell(5).value = { formula: `D${row}*1.5` };
        dataRow.getCell(5).style = { ...styles.cell, numFmt: '#,##0.00' };

        row++;
    });

    const dataEndRow = row - 1;
    let totalCellRef = '';

    if (materialsMap.size > 0) {
        const totalRow = sheet.getRow(row);
        totalRow.getCell(1).value = 'TOTAL:';
        totalRow.getCell(1).style = { ...styles.totalRow, alignment: { horizontal: 'right' } };

        totalRow.getCell(2).value = { formula: `SUM(B${dataStartRow}:B${dataEndRow})` };
        totalRow.getCell(2).style = { ...styles.totalRow, numFmt: '#,##0' };
        totalCellRef = `B${row}`;

        totalRow.getCell(3).value = { formula: `SUM(C${dataStartRow}:C${dataEndRow})` };
        totalRow.getCell(3).style = { ...styles.totalRow, numFmt: '#,##0' };

        totalRow.getCell(4).value = { formula: `SUM(D${dataStartRow}:D${dataEndRow})` };
        totalRow.getCell(4).style = { ...styles.totalRow, numFmt: '#,##0.00' };

        totalRow.getCell(5).value = { formula: `SUM(E${dataStartRow}:E${dataEndRow})` };
        totalRow.getCell(5).style = { ...styles.totalRow, numFmt: '#,##0.00' };

        row++;
    }

    return { endRow: row, totalCellRef };
}

function writeCombinedSummaryTable(sheet, crudoData, htrData, startRow, styles) {
    let row = startRow;

    const headers = ['MATERIAL', 'KG REQ', 'QQ (=B/46)', 'INGRESOS (=B/5800)', 'DIAS (=D*1.5)'];
    const headerRow = sheet.getRow(row);
    headers.forEach((h, i) => {
        headerRow.getCell(i + 1).value = h;
        headerRow.getCell(i + 1).style = styles.header;
    });
    headerRow.height = 22;
    row++;

    const dataStartRow = row;
    const materialsMap = new Map();

    const processData = (data, isHtr) => {
        data.forEach(r => {
            if (!r.isMezcla) {
                const factor = getGroupFactorExcel(isHtr, r.group || '');
                const kgReq = Math.round((r.kg || 0) / factor);
                let matName = r.group || 'SIN GRUPO';
                const hilado = (r.hilado || '').toUpperCase();
                if (/(?:ORGANICO|ORG|ORGANIC)/i.test(hilado)) {
                    if (/\(OCS\)/.test(hilado)) matName += ' (OCS)';
                    else if (/\(GOTS\)/.test(hilado)) matName += ' (GOTS)';
                }
                if (!materialsMap.has(matName)) materialsMap.set(matName, { kg: 0, isAlgodon: isAlgodonTextExcel(matName) });
                materialsMap.get(matName).kg += kgReq;
            } else {
                const comps = isHtr 
                    ? parseMixCompositionHtr(r.hilado || '', r.kg || 0, { groupName: r.group })
                    : parseMixComposition(r.hilado || '', r.kg || 0, { groupName: r.group });
                comps.forEach(c => {
                    let matName = getBaseMaterialTokenExcel(c.name) || c.name;
                    const hilado = (r.hilado || '').toUpperCase();
                    if (/(?:ORGANICO|ORG|ORGANIC)/i.test(c.name)) {
                        if (/\(OCS\)/.test(hilado)) matName += ' (OCS)';
                        else if (/\(GOTS\)/.test(hilado)) matName += ' (GOTS)';
                    }
                    if (!materialsMap.has(matName)) materialsMap.set(matName, { kg: 0, isAlgodon: isAlgodonMixComponentExcel(c.name) });
                    materialsMap.get(matName).kg += c.kg || 0;
                });
            }
        });
    };

    processData(crudoData, false);
    processData(htrData, true);

    materialsMap.forEach((matData, matName) => {
        const dataRow = sheet.getRow(row);
        dataRow.getCell(1).value = matName;
        dataRow.getCell(1).style = { ...styles.cell };
        dataRow.getCell(2).value = matData.kg;
        dataRow.getCell(2).style = { ...styles.cell, numFmt: '#,##0' };
        if (matData.isAlgodon) {
            dataRow.getCell(3).value = { formula: `ROUND(B${row}/46,0)` };
        } else {
            dataRow.getCell(3).value = '-';
        }
        dataRow.getCell(3).style = { ...styles.cell, numFmt: '#,##0' };
        dataRow.getCell(4).value = { formula: `B${row}/5800` };
        dataRow.getCell(4).style = { ...styles.cell, numFmt: '#,##0.00' };
        dataRow.getCell(5).value = { formula: `D${row}*1.5` };
        dataRow.getCell(5).style = { ...styles.cell, numFmt: '#,##0.00' };
        row++;
    });

    const dataEndRow = row - 1;

    if (materialsMap.size > 0) {
        const totalRow = sheet.getRow(row);
        totalRow.getCell(1).value = 'TOTAL COMBINADO:';
        totalRow.getCell(1).style = { ...styles.totalRow, alignment: { horizontal: 'right' }, font: { bold: true, color: { argb: 'FF2563eb' } } };
        totalRow.getCell(2).value = { formula: `SUM(B${dataStartRow}:B${dataEndRow})` };
        totalRow.getCell(2).style = { ...styles.totalRow, numFmt: '#,##0', font: { bold: true, color: { argb: 'FF2563eb' } } };
        totalRow.getCell(3).value = { formula: `SUM(C${dataStartRow}:C${dataEndRow})` };
        totalRow.getCell(3).style = { ...styles.totalRow, numFmt: '#,##0' };
        totalRow.getCell(4).value = { formula: `SUM(D${dataStartRow}:D${dataEndRow})` };
        totalRow.getCell(4).style = { ...styles.totalRow, numFmt: '#,##0.00' };
        totalRow.getCell(5).value = { formula: `SUM(E${dataStartRow}:E${dataEndRow})` };
        totalRow.getCell(5).style = { ...styles.totalRow, numFmt: '#,##0.00' };
        row++;
    }

    return { endRow: row };
}

// =====================================================
// FUNCIONES AUXILIARES
// =====================================================

function groupByMaterialExcel(items) {
    const groups = {};
    items.forEach(r => {
        const groupName = r.group || 'SIN GRUPO';
        if (!groups[groupName]) groups[groupName] = { name: groupName, items: [], isMezcla: !!r.isMezcla };
        groups[groupName].items.push(r);
    });
    return Object.values(groups).sort((a, b) => {
        if (a.name === 'OTROS') return 1;
        if (b.name === 'OTROS') return -1;
        return a.name.localeCompare(b.name);
    });
}

function groupByTituloExcel(items) {
    const groups = {};
    items.forEach(r => {
        const titulo = r.titulo || 'SIN TITULO';
        if (!groups[titulo]) groups[titulo] = { titulo, items: [] };
        groups[titulo].items.push(r);
    });
    return Object.values(groups).sort((a, b) => {
        if (a.titulo === 'OTROS') return 1;
        if (b.titulo === 'OTROS') return -1;
        return a.titulo.localeCompare(b.titulo);
    });
}

function getGroupFactorExcel(isHtr, groupName) {
    if (typeof getGroupFactor === 'function') return getGroupFactor(isHtr, groupName);
    return isHtr ? 0.60 : 0.65;
}

function isAlgodonTextExcel(text) {
    if (typeof isAlgodonText === 'function') return isAlgodonText(text);
    return /ALGODON|PIMA|TANGUIS|UPLAND|ORGANICO|BCI/i.test(text || '');
}

function isAlgodonMixComponentExcel(value) {
    if (typeof isAlgodonMixComponent === 'function') return isAlgodonMixComponent(value);
    const norm = String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toUpperCase();
    return /\b(PIMA|UPLAND|TANGUIS|ORGANICO|ORGANIC|ORG)\b/.test(norm);
}

function getNeFromItemExcel(item) {
    if (typeof getNeFromItem === 'function') return getNeFromItem(item);
    return item.neValue || item.ne || null;
}

function getBaseMaterialTokenExcel(s) {
    if (typeof getBaseMaterialToken === 'function') return getBaseMaterialToken(s);
    if (!s) return '';
    const u = String(s).toUpperCase();
    let base = '';
    if (u.includes('PIMA')) base = 'PIMA';
    else if (u.includes('LYOCELL') || u.includes('TENCEL')) base = 'LYOCELL';
    else if (u.includes('MODAL')) base = 'MODAL';
    else if (u.includes('TANGUIS')) base = 'TANGUIS';
    let mod = '';
    if (u.includes('ORGANICO') || u.includes('ORGANIC')) mod = ' ORGANICO';
    else if (u.includes('BCI')) mod = ' BCI';
    if (!base && mod) return 'TANGUIS' + mod;
    if (base && mod) return base + mod;
    if (base) return base;
    return '';
}

function getClientCertExcel(client) {
    if (typeof getClientCert === 'function') return getClientCert(client);
    if (!client) return 'GOTS';
    const key = String(client).toUpperCase().trim();
    if (key === 'LLL') return 'OCS';
    return 'GOTS';
}

function computeWeightedNeGlobal() {
    if (typeof computeWeightedNe === 'function') {
        const allItems = (GLOBAL_DATA.nuevo || []).concat(GLOBAL_DATA.htr || []);
        return computeWeightedNe(allItems);
    }
    return null;
}

function getColumnLetter(colNum) {
    let letter = '';
    let temp;
    while (colNum > 0) {
        temp = (colNum - 1) % 26;
        letter = String.fromCharCode(temp + 65) + letter;
        colNum = Math.floor((colNum - temp - 1) / 26);
    }
    return letter;
}
