(function () {
    if (window.downloadHabilitadoProgExcel) return;

    const H = ['P', 'HOD', 'F.ING.Prog', 'F.HAB', 'STATUS', 'CLI', 'OC', 'COLOR', 'PDS', 'PDA', 'CERT', 'COMP/OTROS', 'OBSERVACIONES', 'VALIDA', 'H', 'PLANTA', 'LINEA', 'HAB'];
    const W = [5, 9, 9, 9, 11, 7, 16, 13, 8, 10, 9, 32, 30, 8, 5, 11, 10, 11];
    const S = [
        { f: 'PROG 1T', n: 'PROG_1T' },
        { f: 'PROG 2T', n: 'PROG_2T' },
        { f: 'PROG 3T', n: 'PROG_3T' },
        { f: 'S/DESTINO', n: 'PREP_S_DESTINO' }
    ];
    const BLUE = 'FF0070C0', GROUP = 'FFFFE699', GROUP_FONT = 'FF1E40AF', A = 'FFE8F0FE', B = 'FFFFFFFF', WARN = 'FFFCE8E8';
    const BORDER = { style: 'thin', color: { argb: 'FFDDE3EA' } };

    const b = (c) => { c.border = { top: BORDER, left: BORDER, bottom: BORDER, right: BORDER }; };
    const p = (row) => {
        const t = String(getVal(row, 'tipo-transfer') || getVal(row, 'TIPO-TRANSFER') || getVal(row, 'tipo_transfer') || '').toUpperCase().trim();
        const n = String(getVal(row, 'n.transfxpda') || getVal(row, 'N.TRANSFXPDA') || getVal(row, 'n_transfxpda') || '').trim();
        return (t === 'EN PIEZA' && !isNaN(parseInt(n, 10)) && n !== '') ? 'Pza' : 'Otros';
    };

    function status(row) {
        const ruta = String(getVal(row, 'RUTA TELA') || getVal(row, 'RUTA_TELA') || getVal(row, 'RUTA') || '').toUpperCase().trim();
        const corte = String(getVal(row, 'STATUS_CORTE') || getVal(row, 'STATUS') || getVal(row, 'estado_corte') || getVal(row, 'ESTADO_CORTE') || '').toUpperCase().trim();
        const bloq = String(getVal(row, 'estado_bloqueo') || getVal(row, 'ESTADO_BLOQUEO') || '').toUpperCase().trim();
        const lav = String(getVal(row, 'estado_lavada') || getVal(row, 'ESTADO_LAVADA') || '').toUpperCase().trim();
        const ev = String(getVal(row, 'estado_enumerado') || getVal(row, 'ESTADO_ENUMERADO') || '').toUpperCase().trim();
        if (ruta === 'ACABADA') {
            if (!corte || corte === 'X PROG') return 'x cortar';
            if (['PROG', 'PROG 1T', 'PROG 2T', 'PROG 3T'].includes(corte)) return 'Proc Corte';
            if (corte === 'OK') return (ev === 'OK ENM' || ev === 'OK PAQUETEO') ? 'x Hab' : 'x enm';
        } else if (ruta === 'LAVADA') {
            if (!bloq) return 'x pedir';
            if (bloq === 'X PROG') return 'x bloq';
            if (bloq === 'PROG') return 'x Bloq';
            if (bloq === 'OK') {
                if (lav !== 'OK') return 'x lavar';
                if (!corte || corte === 'X PROG') return 'x cortar';
                if (['PROG 1T', 'PROG 2T', 'PROG 3T'].includes(corte)) return 'Proc Corte';
                if (corte === 'OK') return (ev === 'OK ENM' || ev === 'OK PAQUETEO') ? 'x Hab' : 'x enm';
            }
        }
        return '';
    }

    function compPieces(row) {
        const out = [];
        const add = (label, value) => {
            const text = String(value || '').replace(/\s+/g, ' ').trim();
            const norm = text.toUpperCase();
            if (!text || text === '-' || norm === 'X' || norm === 'NO LLEVA') return;
            out.push({ label, value: text });
        };
        add('RIB', getVal(row, 'estado_rib') || getVal(row, 'ESTADO_RIB') || getVal(row, 'RIB') || '');

        const eb = getVal(row, 'ESTADO BLOQUES') || getVal(row, 'ESTADO_BLOQUES') || getVal(row, 'estado_bloques') || '';
        const ecb = getVal(row, 'estado_corte_bloques') || getVal(row, 'ESTADO_CORTE_BLOQUES') || getVal(row, 'ESTADO CORTE BLOQUES') || '';
        const ebn = String(eb || '').toUpperCase().trim();
        const ecbn = String(ecb || '').toUpperCase().trim();
        let bloq = '';
        if (ebn === 'NO LLEVA') bloq = 'NO LLEVA';
        else if (ebn.indexOf('OK CORTE') !== -1) {
            if (!ecbn) bloq = 'X PROG';
            else if (ecbn.indexOf('PROG') !== -1) bloq = 'PROG';
            else if (ecbn.indexOf('OK') !== -1) bloq = 'OK';
            else bloq = eb;
        } else bloq = eb;
        add('BLOQ?', bloq);
        add('COLL/TAP', getVal(row, 'estado_coll_tap') || getVal(row, 'ESTADO_COLL_TAP') || getVal(row, 'ESTADO COLL TAP') || '');

        const tt = String(getVal(row, 'tipo-transfer') || getVal(row, 'TIPO-TRANSFER') || getVal(row, 'tipo_transfer') || '').toUpperCase().trim();
        const nt = String(getVal(row, 'n.transfxpda') || getVal(row, 'N.TRANSFXPDA') || getVal(row, 'n_transfxpda') || '').trim();
        const et = String(getVal(row, 'estado_transfer') || getVal(row, 'ESTADO_TRANSFER') || '').trim();
        let trsf = '';
        if (tt === 'NO LLEVA' || nt.toUpperCase() === 'NO LLEVA') trsf = 'NO LLEVA';
        else if (tt === 'EN PIEZA' && !isNaN(parseInt(nt, 10)) && nt !== '') trsf = `Pza(x${parseInt(nt, 10)})-${et || 'X PROG'}`;
        else if (tt === 'EN PRENDA' && !isNaN(parseInt(nt, 10)) && nt !== '') trsf = `PDA[x${parseInt(nt, 10)}]-${et || 'X PROG'}`;
        else trsf = nt;
        add('TRSF', trsf);

        const bd = getVal(row, 'estado_bordado') || getVal(row, 'ESTADO_BORDADO') || '';
        const nbd = String(getVal(row, 'n.BDxpda') || getVal(row, 'n.bordadoxpda') || getVal(row, 'N.BDXPDA') || '').toUpperCase().trim();
        const bdn = String(bd || '').toUpperCase().trim();
        let bord = '';
        if (bdn === 'NO LLEVA' || nbd === 'NO LLEVA') bord = 'NO LLEVA';
        else if (bdn === 'PROG') bord = 'PROG';
        else if ((!bdn || bdn === undefined) && nbd && !isNaN(parseInt(nbd, 10))) bord = 'X PROG';
        else if (bdn === 'OK') bord = 'OK';
        else bord = bd;
        add('BORD', bord);

        const est = getVal(row, 'estado_estampado') || getVal(row, 'ESTADO_ESTAMPADO') || '';
        const nest = String(getVal(row, 'n.ESTAMPxpda') || getVal(row, 'n.ESTAMP xpda') || getVal(row, 'N.ESTAMPXPDA') || getVal(row, 'n.ESTAMPxpda ') || '').toUpperCase().trim();
        const estn = String(est || '').toUpperCase().trim();
        let estm = '';
        if (nest.indexOf('NO LLEVA') !== -1 || estn === 'NO LLEVA') estm = 'NO LLEVA';
        else if (estn === 'PROG') estm = 'PROG';
        else if ((!estn || estn === undefined) && nest && !isNaN(parseInt(nest, 10))) estm = 'X PROG';
        else if (estn === 'OK') estm = 'OK';
        else estm = est;
        add('ESTM', estm);
        return out;
    }

    const compRich = (row) => {
        const arr = compPieces(row);
        if (!arr.length) return '';
        const rt = [];
        arr.forEach((x, i) => {
            if (i) rt.push({ text: ' | ' });
            rt.push({ text: `${x.label}: `, font: { name: 'Calibri', size: 11, bold: true } });
            rt.push({ text: x.value, font: { name: 'Calibri', size: 11 } });
        });
        return { richText: rt };
    };

    function formatFHabExport(row) {
        try {
            const raw = typeof getRawFIngRealFromRow === 'function' ? getRawFIngRealFromRow(row) : '';
            if (!raw) return '';
            if (typeof formatDayMonthFromSheetDateLiteral === 'function') {
                return formatDayMonthFromSheetDateLiteral(raw) || '';
            }
            const text = String(raw).trim();
            const match = text.match(/Date\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)/);
            if (!match) return '';
            const mesesEs = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
            const monthNum = parseInt(match[2], 10);
            const day = String(parseInt(match[3], 10)).padStart(2, '0');
            return mesesEs[monthNum - 1] ? `${day}/${mesesEs[monthNum - 1]}` : '';
        } catch (e) {
            return '';
        }
    }

    const valida = (row) => {
        let idx = findHeaderIndexCaseInsensitive('VALIDACION');
        if (idx === -1) idx = findHeaderIndexCaseInsensitive('VALIDA');
        let raw = '';
        if (idx !== -1 && row[idx] !== undefined && row[idx] !== null) raw = row[idx];
        else {
            const byName = getVal(row, 'VALIDACION') || getVal(row, 'VALIDA');
            if (byName !== undefined && byName !== null) raw = byName;
        }
        const norm = String(raw).trim().toUpperCase();
        return (raw === true || raw === 1 || norm === 'TRUE' || norm === 'VERDADERO' || norm === '1' || norm === 'SI' || norm === 'X') ? '\u2611' : '\u2610';
    };

    function cloneNotasRows(rows) {
        return (rows || []).map((r) => ({
            cliente: String((r && r.cliente) || '').trim(),
            ops: String((r && r.ops) || '').trim(),
            color: String((r && r.color) || '').trim(),
            pds: String((r && r.pds) || '').trim(),
            comentario: String((r && r.comentario) || '').trim(),
            comentario_general: String((r && r.comentario_general) || '').trim()
        }));
    }

    async function getNotasRowsByTurn(turno) {
        const turnoNorm = String(turno || '').toUpperCase().trim();
        if (!turnoNorm) return [];

        let rows = [];
        try {
            if (typeof habilitadoHoja3RowsByTurn !== 'undefined'
                && habilitadoHoja3RowsByTurn
                && Array.isArray(habilitadoHoja3RowsByTurn[turnoNorm])) {
                rows = cloneNotasRows(habilitadoHoja3RowsByTurn[turnoNorm]);
            }
        } catch (e) { }

        if ((!rows || rows.length === 0)
            && typeof currentHabilitadoFilter !== 'undefined'
            && String(currentHabilitadoFilter || '').toUpperCase().trim() === turnoNorm
            && typeof habilitadoHoja3Rows !== 'undefined'
            && Array.isArray(habilitadoHoja3Rows)) {
            rows = cloneNotasRows(habilitadoHoja3Rows);
        }

        if ((!rows || rows.length === 0) && typeof fetchHoja3RowsByTurn === 'function') {
            try {
                const fetched = await fetchHoja3RowsByTurn(turnoNorm);
                if (Array.isArray(fetched)) rows = cloneNotasRows(fetched);
            } catch (e) {
                console.error('Error cargando notas para Excel', e);
            }
        }

        rows = (rows || []).filter((r) => !!(r.cliente || r.ops || r.color || r.pds || r.comentario || r.comentario_general));
        if (rows.length === 0) {
            rows = [{
                cliente: '',
                ops: '',
                color: '',
                pds: '',
                comentario: 'SIN NOTAS REGISTRADAS',
                comentario_general: ''
            }];
        }
        return rows;
    }

    function styleNotesHeader(cell) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F172A' } };
        cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        b(cell);
    }

    function styleNotesCell(cell) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };
        cell.font = { name: 'Calibri', size: 10, color: { argb: 'FF111827' } };
        cell.alignment = { horizontal: 'left', vertical: 'top', wrapText: true };
        b(cell);
    }

    function addNotesSection(ws, startRow, turno, rows) {
        const clientCol = 5;   // E
        const opsStartCol = 6; // F
        const opsEndCol = 8;   // H
        const colorCol = 9;    // I
        const pdsCol = 10;     // J
        const commentCol = 11; // K
        const generalCol = 12; // L

        ws.getColumn(clientCol).width = 12;
        ws.getColumn(opsStartCol).width = 14;
        ws.getColumn(opsStartCol + 1).width = 14;
        ws.getColumn(opsEndCol).width = 14;
        ws.getColumn(colorCol).width = 14;
        ws.getColumn(pdsCol).width = 9;
        ws.getColumn(commentCol).width = 34;
        ws.getColumn(generalCol).width = 34;

        const titleRow = ws.getRow(startRow);
        ws.mergeCells(startRow, clientCol, startRow, commentCol);
        titleRow.getCell(clientCol).value = `NOTAS ${turno}`;
        titleRow.getCell(clientCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
        titleRow.getCell(clientCol).font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        titleRow.getCell(clientCol).alignment = { horizontal: 'left', vertical: 'middle' };
        for (let c = clientCol; c <= commentCol; c++) b(titleRow.getCell(c));

        ws.mergeCells(startRow, generalCol, startRow, generalCol);
        titleRow.getCell(generalCol).value = 'COMENTARIOS GENERALES';
        titleRow.getCell(generalCol).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
        titleRow.getCell(generalCol).font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
        titleRow.getCell(generalCol).alignment = { horizontal: 'left', vertical: 'middle' };
        b(titleRow.getCell(generalCol));

        const headerRow = ws.getRow(startRow + 1);

        ws.mergeCells(startRow + 1, clientCol, startRow + 1, clientCol);
        headerRow.getCell(clientCol).value = 'CLIENTE';
        styleNotesHeader(headerRow.getCell(clientCol));

        ws.mergeCells(startRow + 1, opsStartCol, startRow + 1, opsEndCol);
        headerRow.getCell(opsStartCol).value = 'OPS';
        styleNotesHeader(headerRow.getCell(opsStartCol));
        headerRow.getCell(opsStartCol).alignment = { horizontal: 'center', vertical: 'middle' };

        ws.mergeCells(startRow + 1, colorCol, startRow + 1, colorCol);
        headerRow.getCell(colorCol).value = 'COLOR';
        styleNotesHeader(headerRow.getCell(colorCol));

        ws.mergeCells(startRow + 1, pdsCol, startRow + 1, pdsCol);
        headerRow.getCell(pdsCol).value = 'PDS';
        styleNotesHeader(headerRow.getCell(pdsCol));

        ws.mergeCells(startRow + 1, commentCol, startRow + 1, commentCol);
        headerRow.getCell(commentCol).value = 'COMENTARIO';
        styleNotesHeader(headerRow.getCell(commentCol));

        ws.mergeCells(startRow + 1, generalCol, startRow + 1, generalCol);
        headerRow.getCell(generalCol).value = 'COMENTARIOS GENERALES';
        styleNotesHeader(headerRow.getCell(generalCol));

        const totalRows = rows.length;
        for (let i = 0; i < totalRows; i++) {
            const row = rows[i] || {};
            const rowNumber = startRow + 2 + i;
            const excelRow = ws.getRow(rowNumber);
            const noteLen = Math.max(
                String(row.ops || '').length,
                String(row.comentario || '').length,
                String(row.comentario_general || '').length
            );

            excelRow.getCell(clientCol).value = row.cliente || '';
            styleNotesCell(excelRow.getCell(clientCol));
            excelRow.getCell(clientCol).alignment = { horizontal: 'left', vertical: 'middle' };

            ws.mergeCells(rowNumber, opsStartCol, rowNumber, opsEndCol);
            excelRow.getCell(opsStartCol).value = row.ops || '';
            styleNotesCell(excelRow.getCell(opsStartCol));
            excelRow.getCell(opsStartCol).alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
            for (let c = opsStartCol + 1; c <= opsEndCol; c++) b(excelRow.getCell(c));

            excelRow.getCell(colorCol).value = row.color || '';
            styleNotesCell(excelRow.getCell(colorCol));
            excelRow.getCell(colorCol).alignment = { horizontal: 'center', vertical: 'middle' };

            excelRow.getCell(pdsCol).value = row.pds || '';
            styleNotesCell(excelRow.getCell(pdsCol));
            excelRow.getCell(pdsCol).alignment = { horizontal: 'center', vertical: 'middle' };

            excelRow.getCell(commentCol).value = row.comentario || '';
            styleNotesCell(excelRow.getCell(commentCol));
            excelRow.getCell(commentCol).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

            excelRow.getCell(generalCol).value = row.comentario_general || '';
            styleNotesCell(excelRow.getCell(generalCol));
            excelRow.getCell(generalCol).alignment = { horizontal: 'left', vertical: 'middle', wrapText: true };

            excelRow.height = noteLen > 120 ? 36 : (noteLen > 60 ? 28 : 20);
        }

        return startRow + 1 + totalRows;
    }

    function visibleIndices(filterState) {
        const idxEv = findHeaderIndexCaseInsensitive('estado_enumerado');
        const idxHab = findHeaderIndexCaseInsensitive('estado_habilitado');
        let idxVal = findHeaderIndexCaseInsensitive('VALIDACION');
        if (idxVal === -1) idxVal = findHeaderIndexCaseInsensitive('VALIDA');
        const applyFilters = Array.isArray(habilitadoHeaderFilters) && habilitadoHeaderFilters.length > 0;
        const out = [];

        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            const ev = idxEv !== -1 && row[idxEv] !== undefined ? row[idxEv] : getVal(row, 'estado_enumerado');
            const evNorm = String(ev || '').toUpperCase().trim();
            const hab = idxHab !== -1 && row[idxHab] !== undefined ? row[idxHab] : getVal(row, 'estado_habilitado');
            const habNorm = String(hab || '').toUpperCase().trim();
            const planta = String(getVal(row, 'PLANTA') || '').toUpperCase().trim();

            let ok = false;
            if (filterState === 'X PROG') ok = (habNorm === '' || habNorm === 'X PROG');
            else if (filterState === 'S/DESTINO') ok = ((planta === 'S/DESTINO' && habNorm !== 'OK') || habNorm === 'OK S/DESTINO');
            else ok = (habNorm === filterState);
            if (!ok) continue;

            if (applyFilters) {
                let keep = true;
                for (const hf of habilitadoHeaderFilters) {
                    const fv = getHabilitadoFieldValue(row, hf.field, evNorm);
                    if (!matchesHabilitadoFilterValue(hf.field, fv, String(hf.value || '').trim().toUpperCase())) {
                        keep = false;
                        break;
                    }
                }
                if (!keep) continue;
            }
            out.push(i);
        }

        const ocKey = (row) => {
            const oc = String(getVal(row, 'OC') || '').trim();
            if (oc) return oc.toLowerCase();
            const op = String(getVal(row, 'OP') || '').trim();
            const corte = String(getVal(row, 'CORTE') || '').trim();
            if (op || corte) return `${op}-${corte}`.toLowerCase();
            return `${String(getVal(row, 'OP TELA') || '').trim()}-${String(getVal(row, 'PARTIDA') || '').trim()}`.toLowerCase();
        };
        const pm = (raw) => {
            const v = String(raw || '').trim().toUpperCase();
            if (v === 'D') return { b: 0, n: 0, t: '' };
            if (/^\d+$/.test(v)) return { b: 1, n: parseInt(v, 10), t: '' };
            if (!v) return { b: 3, n: Infinity, t: '' };
            return { b: 2, n: Infinity, t: v };
        };
        const cp = (a, b) => {
            const A = pm(a), B = pm(b);
            if (A.b !== B.b) return A.b - B.b;
            if (A.n !== B.n) return A.n - B.n;
            return A.t.localeCompare(B.t, undefined, { sensitivity: 'base' });
        };

        out.sort((a, b) => {
            const A = rawData[a], B = rawData[b];
            let c = ocKey(A).localeCompare(ocKey(B), undefined, { numeric: true, sensitivity: 'base' });
            if (c) return c;
            c = `${String(A[colMap["OP TELA"]] || "").trim()}-${String(A[colMap["PARTIDA"]] || "").trim()}`.toLowerCase()
                .localeCompare(`${String(B[colMap["OP TELA"]] || "").trim()}-${String(B[colMap["PARTIDA"]] || "").trim()}`.toLowerCase(), undefined, { numeric: true, sensitivity: 'base' });
            if (c) return c;
            const idxP = findPriorityHeaderIndex('habilitado');
            c = cp(idxP !== -1 ? String(A[idxP] || '').trim() : '', idxP !== -1 ? String(B[idxP] || '').trim() : '');
            if (c) return c;
            return (B[colMap["HOD"]] || 0) - (A[colMap["HOD"]] || 0);
        });

        if (filterState === 'PROG 1T' || filterState === 'PROG 2T' || filterState === 'PROG 3T') {
            out.sort((a, b) => {
                const A = rawData[a], B = rawData[b];
                const c1 = (p(A) === 'Pza' ? 0 : 1) - (p(B) === 'Pza' ? 0 : 1);
                if (c1) return c1;
                const idxP = findPriorityHeaderIndex('habilitado');
                const c2 = cp(idxP !== -1 ? String(A[idxP] || '').trim() : '', idxP !== -1 ? String(B[idxP] || '').trim() : '');
                if (c2) return c2;
                const c3 = ocKey(A).localeCompare(ocKey(B), undefined, { numeric: true, sensitivity: 'base' });
                if (c3) return c3;
                const c4 = `${String(A[colMap["OP TELA"]] || "").trim()}-${String(A[colMap["PARTIDA"]] || "").trim()}`.toLowerCase()
                    .localeCompare(`${String(B[colMap["OP TELA"]] || "").trim()}-${String(B[colMap["PARTIDA"]] || "").trim()}`.toLowerCase(), undefined, { numeric: true, sensitivity: 'base' });
                if (c4) return c4;
                return (B[colMap["HOD"]] || 0) - (A[colMap["HOD"]] || 0);
            });
        }
        return out;
    }

    const rowData = (row) => ({
        p: String(getPriorityValueFromRow(row, 'habilitado') || '').trim(),
        hod: formatValue(getVal(row, 'HOD'), 'date') || '',
        fing: formatValue(getVal(row, 'F.ING.COST'), 'date') || '',
        fhab: formatFHabExport(row) || '',
        status: status(row),
        cli: normalizeClientName(getVal(row, 'CLIENTE')) || '',
        oc: String(getVal(row, 'OC') || ((getVal(row, 'OP') || '') + '-' + (getVal(row, 'CORTE') || ''))).trim(),
        color: abbreviateHeather(getVal(row, 'COLOR') || ''),
        pds: formatThousands(parseFloat(getVal(row, 'PDS GIRADAS')) || 0, 0),
        pda: normalizePrenda(getVal(row, 'PRENDA') || ''),
        cert: normalizeTipoCert(getVal(row, 'TIPO CERTIFICADO') || ''),
        comp: compRich(row),
        obs: String(getVal(row, 'OBSERVACIONES') || getVal(row, 'OBSERVACION') || getVal(row, 'OBS') || '').replace(/\s+/g, ' ').trim(),
        valida: valida(row),
        h: isHabilitadoHMarcada(row) ? '\u2611' : '\u2610',
        planta: normalizeHabilitadoPlantaValue(getVal(row, 'PLANTA') || '') || 'XASIG',
        linea: String(getVal(row, 'LINEA') || '').trim() || 'XASIG',
        hab: String(getVal(row, 'estado_habilitado') || getVal(row, 'ESTADO_HABILITADO') || '').toString().toUpperCase().trim() || 'X PROG',
        priority1: String(getPriorityValueFromRow(row, 'habilitado') || '').trim() === '1'
    });

    function styleRow(row, data, fill) {
        const rowFill = data.priority1 ? WARN : fill;
        row.eachCell((cell, i) => {
            const h = H[i - 1];
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowFill } };
            cell.font = { name: 'Calibri', size: 11, color: { argb: 'FF1E293B' } };
            cell.alignment = { vertical: 'middle' };
            b(cell);
            if (h === 'P' || h === 'PDS' || h === 'PLANTA' || h === 'LINEA' || h === 'HAB' || h === 'HOD' || h === 'F.ING.PROG' || h === 'F.HAB') {
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
            } else if (h === 'OC' || h === 'CLI' || h === 'COLOR' || h === 'PDA' || h === 'CERT') {
                cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } else if (h === 'COMP/OTROS' || h === 'OBSERVACIONES') {
                cell.alignment = { horizontal: 'left', vertical: 'middle' };
            } else if (h === 'VALIDA' || h === 'H') {
                cell.alignment = { horizontal: 'center', vertical: 'middle' };
                cell.font = { name: 'Segoe UI Symbol', size: 12, color: { argb: 'FF2563EB' } };
            }
        });

        const s = String(data.status || '').toUpperCase().trim();
        const st = row.getCell(5);
        if (s.includes('X CORTAR') || s.includes('X ENM') || s.includes('X BLOQ') || s.includes('X LAVAR')) {
            st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5F8F1' } };
            st.font = { name: 'Calibri', size: 11, color: { argb: 'FF065F46' }, bold: true };
        } else if (s.includes('PROC CORTE')) {
            st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAF2FF' } };
            st.font = { name: 'Calibri', size: 11, color: { argb: 'FF075985' }, bold: true };
        } else if (s.includes('X HAB')) {
            st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0D3C78' } };
            st.font = { name: 'Calibri', size: 11, color: { argb: 'FFFFFFFF' }, bold: true };
        } else if (s.includes('X PEDIR')) {
            st.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE68A' } };
            st.font = { name: 'Calibri', size: 11, color: { argb: 'FF9A3412' }, bold: true };
        }

        const hb = String(data.hab || '').toUpperCase().trim();
        const hc = row.getCell(18);
        if (hb === 'PROG 1T' || hb === 'PROG 2T' || hb === 'PROG 3T') {
            hc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFDE68A' } };
            hc.font = { name: 'Calibri', size: 11, color: { argb: 'FF92400E' }, bold: true };
        } else if (hb === 'OK' || hb === 'OK S/DESTINO') {
            hc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE5F8F1' } };
            hc.font = { name: 'Calibri', size: 11, color: { argb: 'FF047857' }, bold: true };
        } else if (hb === 'DEPURADO') {
            hc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
            hc.font = { name: 'Calibri', size: 11, color: { argb: 'FF991B1B' }, bold: true };
        } else if (hb === 'X PROG' || !hb) {
            hc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3F4F6' } };
            hc.font = { name: 'Calibri', size: 11, color: { argb: 'FF374151' }, bold: true };
        }

        if (data.priority1) {
            row.eachCell((cell, i) => {
                if (![5, 14, 15, 18].includes(i)) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: WARN } };
            });
        }
    }

    async function sheet(wb, f, n) {
        const ws = wb.addWorksheet(n);
        ws.properties.defaultRowHeight = 18;
        ws.views = [{ state: 'frozen', ySplit: 1 }];
        ws.pageSetup = {
            paperSize: 9,
            orientation: 'landscape',
            fitToPage: true,
            fitToWidth: 1,
            fitToHeight: 0,
            margins: {
                left: 0.1969,
                right: 0.1969,
                top: 0.1969,
                bottom: 0.1969,
                header: 0.0,
                footer: 0.0
            }
        };
        ws.pageSetup.printTitlesRow = '1:1';
        ws.columns = W.map(width => ({ width }));
        ws.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: H.length } };

        const head = ws.addRow(H);
        head.height = 22;
        head.eachCell(cell => {
            cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } };
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            b(cell);
        });

        const idx = visibleIndices(f);
        const prog = (f === 'PROG 1T' || f === 'PROG 2T' || f === 'PROG 3T');
        let lastG = null, lastOP = null, band = 'a';
        let totalPds = 0;

        idx.forEach(ri => {
            const src = rawData[ri];
            const g = p(src);
            if (prog && g !== lastG) {
                let pdsTrsf = 0;
                let transferTrsf = 0;
                idx.forEach(i => {
                    const r = rawData[i];
                    if (p(r) !== g || isHabilitadoValidacionMarcada(r)) return;
                    const pdsRow = getHabilitadoPdsValue(r);
                    const transfersRow = getHabilitadoTransferMultiplierValue(r) * pdsRow;
                    pdsTrsf += pdsRow;
                    transferTrsf += transfersRow;
                });
                const label = g === 'Pza' ? 'LLEVA TRANSFER EN PIEZA' : 'No lleva transfer en pieza';
                const gr = ws.addRow([`${label} [${formatThousands(pdsTrsf, 0)}pds - ${formatThousands(transferTrsf, 0)}transfers]`]);
                ws.mergeCells(gr.number, 1, gr.number, H.length);
                gr.height = 20;
                for (let c = 1; c <= H.length; c++) {
                    const cell = gr.getCell(c);
                    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: GROUP } };
                    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: GROUP_FONT } };
                    cell.alignment = { horizontal: 'left', vertical: 'middle' };
                    b(cell);
                }
                lastG = g;
                lastOP = null;
                band = 'a';
            }

            const d = rowData(src);
            totalPds += parseFloat(getVal(src, 'PDS GIRADAS')) || 0;
            const r = ws.addRow([d.p, d.hod, d.fing, d.fhab, d.status, d.cli, d.oc, d.color, d.pds, d.pda, d.cert, d.comp || '', d.obs, d.valida, d.h, d.planta, d.linea, d.hab]);
            const op = `${String(getVal(src, 'OP TELA') || '').trim()}-${String(getVal(src, 'PARTIDA') || '').trim()}`;
            if (lastOP !== null && op !== lastOP) band = (band === 'a') ? 'b' : 'a';
            lastOP = op;
            styleRow(r, d, band === 'a' ? A : B);
            if (d.comp && d.comp.richText) r.getCell(12).value = d.comp;
            r.getCell(12).alignment = { horizontal: 'left', vertical: 'middle' };
            r.getCell(13).alignment = { horizontal: 'left', vertical: 'middle' };
            r.getCell(14).alignment = { horizontal: 'center', vertical: 'middle' };
            r.getCell(15).alignment = { horizontal: 'center', vertical: 'middle' };

            r.height = 20;
        });

        if (f === 'S/DESTINO') {
            const totalRow = ws.addRow(['TOTAL', '', '', '', '', '', '', Math.round(totalPds), '', '', '', '', '', '', '', '', '']);
            totalRow.height = 20;
            for (let c = 1; c <= H.length; c++) {
                const cell = totalRow.getCell(c);
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF000000' } };
                cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
                cell.border = { top: BORDER, left: BORDER, bottom: BORDER, right: BORDER };
            }
            totalRow.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
            totalRow.getCell(8).alignment = { horizontal: 'center', vertical: 'middle' };
            totalRow.getCell(8).numFmt = '#,##0';
        } else if (prog) {
            const notasRows = await getNotasRowsByTurn(f);
            const startRow = ws.lastRow ? ws.lastRow.number + 3 : 3;
            addNotesSection(ws, startRow, f, notasRows);
        }
    }

    window.downloadHabilitadoProgExcel = async function () {
        try {
            if (!rawData || rawData.length < 2) {
                alert('No hay datos disponibles para descargar.');
                return;
            }
            const wb = new ExcelJS.Workbook();
            wb.creator = 'PCP Confecciones';
            wb.created = new Date();
            wb.modified = new Date();
            for (const x of S) {
                await sheet(wb, x.f, x.n);
            }

            const now = new Date();
            const m = ['ene', 'feb', 'mar', 'abr', 'may', 'jun', 'jul', 'ago', 'sep', 'oct', 'nov', 'dic'][now.getMonth()];
            const file = `prog_habilitado_${String(now.getDate()).padStart(2, '0')}-${m}-${now.getFullYear()}_${String(now.getHours()).padStart(2, '0')}-${String(now.getMinutes()).padStart(2, '0')}`;
            const buffer = await wb.xlsx.writeBuffer();
            const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `${file}.xlsx`;
            a.click();
            setTimeout(() => URL.revokeObjectURL(url), 1000);
        } catch (e) {
            console.error('Error exporting Habilitado Excel', e);
            alert('No se pudo generar el Excel de Habilitado.');
        }
    };
})();
