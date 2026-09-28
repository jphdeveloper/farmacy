const appState = {
  historyData: null,
  inventoryData: null,
  priceListData: null,
  latestWorkbook: null
};

// Umbral mínimo de similitud (0-1) para aceptar una coincidencia aproximada
// entre la descripción del inventario y la lista de precios oficiales.
const FUZZY_MATCH_THRESHOLD = 0.8;

const elements = {
  historyFile: document.querySelector('#historyFile'),
  inventoryFile: document.querySelector('#inventoryFile'),
  priceListFile: document.querySelector('#priceListFile'),
  reportMonth: document.querySelector('#reportMonth'),
  periodMonths: document.querySelector('#periodMonths'),
  periodOptions: [...document.querySelectorAll('[data-period-months]')],
  clearBtn: document.querySelector('#clearBtn'),
  generateBtn: document.querySelector('#generateBtn'),
  downloadBtn: document.querySelector('#downloadBtn'),
  statusBox: document.querySelector('#statusBox'),
  resultsBody: document.querySelector('#resultsBody'),
  historyCount: document.querySelector('#historyCount'),
  inventoryCount: document.querySelector('#inventoryCount'),
  priceListCount: document.querySelector('#priceListCount'),
  resultCount: document.querySelector('#resultCount'),
  unitsToBuy: document.querySelector('#unitsToBuy'),
  estimatedValue: document.querySelector('#estimatedValue'),
  monthsUsed: document.querySelector('#monthsUsed'),
  savingsValue: document.querySelector('#savingsValue'),
  coverageText: document.querySelector('#coverageText'),
  priceCoverageCard: document.querySelector('#priceCoverageCard'),
  priceCoverageText: document.querySelector('#priceCoverageText'),
  currentYear: document.querySelector('#currentYear')
};

elements.reportMonth.value = getCurrentMonthValue();
if (elements.currentYear) {
  elements.currentYear.textContent = String(new Date().getFullYear());
}

elements.reportMonth.addEventListener('click', () => {
  if (typeof elements.reportMonth.showPicker !== 'function') return;
  try {
    elements.reportMonth.showPicker();
  } catch {
    elements.reportMonth.focus();
  }
});

elements.historyFile.addEventListener('change', async () => {
  const file = elements.historyFile.files?.[0];
  if (!file) return;
  setStatus(`Leyendo histórico: ${file.name}`);
  try {
    appState.historyData = await parseHistoryWorkbook(file);
    elements.historyCount.textContent = `${appState.historyData.rows.length.toLocaleString('es-CO')} filas`;
    setStatus(`Histórico cargado: ${appState.historyData.rows.length.toLocaleString('es-CO')} filas y ${appState.historyData.months.length} meses detectados.`);
    autoGenerateIfReady();
  } catch (error) {
    appState.historyData = null;
    setStatus(`No se pudo leer el histórico: ${error.message}`);
  }
});

elements.inventoryFile.addEventListener('change', async () => {
  const file = elements.inventoryFile.files?.[0];
  if (!file) return;
  setStatus(`Leyendo inventario: ${file.name}`);
  try {
    appState.inventoryData = await parseInventoryWorkbook(file);
    elements.inventoryCount.textContent = `${appState.inventoryData.rows.length.toLocaleString('es-CO')} productos`;
    setStatus(`Inventario cargado: ${appState.inventoryData.rows.length.toLocaleString('es-CO')} productos listos para cruzar.`);
    autoGenerateIfReady();
  } catch (error) {
    appState.inventoryData = null;
    setStatus(`No se pudo leer el inventario: ${error.message}`);
  }
});

elements.priceListFile.addEventListener('change', async () => {
  const file = elements.priceListFile.files?.[0];
  if (!file) return;
  setStatus(`Leyendo lista de precios oficiales: ${file.name}`);
  try {
    appState.priceListData = await parsePriceListWorkbook(file);
    elements.priceListCount.textContent = `${appState.priceListData.productCount.toLocaleString('es-CO')} productos`;
    setStatus(`Lista de precios cargada: ${appState.priceListData.productCount.toLocaleString('es-CO')} productos de ${appState.priceListData.providerCount.toLocaleString('es-CO')} proveedores.`);
    await autoGenerateIfReady();
  } catch (error) {
    appState.priceListData = null;
    elements.priceListCount.textContent = '0 productos';
    setStatus(`No se pudo leer la lista de precios: ${error.message}`);
  }
});

elements.generateBtn.addEventListener('click', async () => {
  await generateReport();
});

elements.periodOptions.forEach((option) => {
  option.addEventListener('click', async () => {
    elements.periodOptions.forEach((item) => {
      const isActive = item === option;
      item.classList.toggle('is-active', isActive);
      item.setAttribute('aria-pressed', String(isActive));
    });
    await autoGenerateIfReady();
  });
});

elements.clearBtn?.addEventListener('click', () => {
  clearAllFields();
});

elements.downloadBtn.addEventListener('click', () => {
  if (!appState.latestWorkbook) return;
  const reportDate = elements.reportMonth.value || getCurrentMonthValue();
  XLSX.writeFile(appState.latestWorkbook, `reporte_compra_medicamentos_${reportDate.replace('-', '_')}.xlsx`);
});

async function autoGenerateIfReady() {
  if (appState.historyData && appState.inventoryData) {
    await generateReport();
  }
}

async function generateReport() {
  if (!appState.historyData || !appState.inventoryData) {
    setStatus('Faltan archivos por cargar. Necesito histórico e inventario.');
    return;
  }

  const reportMonth = parseMonthValue(elements.reportMonth.value || getCurrentMonthValue());
  const monthsRequested = getSelectedPeriodMonths();
  const monthsToUse = appState.historyData.months.filter((month) => month < reportMonth).slice(-monthsRequested);

  if (!monthsToUse.length) {
    setStatus('No hay meses históricos anteriores al mes seleccionado. Cambia el mes del reporte o revisa los datos.');
    return;
  }

  const historicalConsumption = aggregateHistoricalConsumption(appState.historyData.rows, monthsToUse);
  const priceListData = appState.priceListData;

  const reportRows = appState.inventoryData.rows.map((item) => {
    const monthlyData = historicalConsumption.get(item.reference) || { totalExits: 0 };
    const averageConsumption = monthlyData.totalExits / monthsToUse.length;
    const currentStock = toNumber(item.stock);
    const internalCost = toNumber(item.cost);

    const priceMatch = priceListData ? findOfficialPrice(item.description, priceListData) : { matchType: 'sin_lista' };
    const officialPrice = priceMatch.price ?? null;
    const usedCost = officialPrice != null ? officialPrice : internalCost;

    const recommendedUnits = Math.max(0, Math.ceil(averageConsumption - currentStock));
    const estimatedValue = roundMoney(recommendedUnits * usedCost);
    const priceDifference = officialPrice != null ? roundNumber(officialPrice - internalCost, 2) : 0;
    // Positivo = ahorro (el precio oficial es menor al costo interno registrado); negativo = sobrecosto.
    const savingsImpact = officialPrice != null ? roundMoney(recommendedUnits * (internalCost - officialPrice)) : 0;

    return {
      reference: item.reference,
      description: item.description,
      averageConsumption: roundNumber(averageConsumption, 2),
      currentStock: roundNumber(currentStock, 2),
      recommendedUnits,
      internalCost: roundNumber(internalCost, 2),
      officialPrice: officialPrice != null ? roundNumber(officialPrice, 2) : null,
      cost: roundNumber(usedCost, 2),
      priceDifference,
      savingsImpact,
      suggestedProvider: priceMatch.provider || '—',
      matchType: priceMatch.matchType,
      matchScore: priceMatch.score ?? null,
      matchedPriceDescription: priceMatch.matchedDescription || '',
      estimatedValue,
      status: recommendedUnits > 0 ? 'Compra sugerida' : 'Sin compra'
    };
  });

  const sortedRows = reportRows
    .filter((row) => row.recommendedUnits > 0 || row.averageConsumption > 0)
    .sort((a, b) => b.estimatedValue - a.estimatedValue || b.recommendedUnits - a.recommendedUnits || a.reference.localeCompare(b.reference));

  const priceCoverage = summarizePriceCoverage(reportRows, priceListData);

  appState.latestWorkbook = buildWorkbook(sortedRows, monthsToUse, reportMonth, appState.historyData, appState.inventoryData, priceCoverage);

  renderReport(sortedRows, monthsToUse.length, appState.historyData, appState.inventoryData, priceCoverage);
  elements.downloadBtn.disabled = false;
  setStatus(`Reporte generado para ${formatMonth(reportMonth)} con ${monthsToUse.length} mes${monthsToUse.length === 1 ? '' : 'es'} de histórico.`);
}

function summarizePriceCoverage(allRows, priceListData) {
  if (!priceListData) return null;

  const exact = allRows.filter((row) => row.matchType === 'exacta').length;
  const approximate = allRows.filter((row) => row.matchType === 'aproximada').length;
  const noMatch = allRows.filter((row) => row.matchType === 'sin_match').length;
  const totalSavingsImpact = allRows.reduce((sum, row) => sum + row.savingsImpact, 0);

  return {
    total: allRows.length,
    exact,
    approximate,
    noMatch,
    totalSavingsImpact: roundMoney(totalSavingsImpact),
    priceListProductCount: priceListData.productCount,
    priceListProviderCount: priceListData.providerCount
  };
}

function renderReport(rows, monthsUsed, history, inventory, priceCoverage) {
  elements.resultCount.textContent = `${rows.length.toLocaleString('es-CO')} compras`;
  elements.monthsUsed.textContent = `${monthsUsed}`;

  const unitsToBuy = rows.reduce((sum, row) => sum + row.recommendedUnits, 0);
  const estimatedValue = rows.reduce((sum, row) => sum + row.estimatedValue, 0);

  elements.unitsToBuy.textContent = unitsToBuy.toLocaleString('es-CO');
  elements.estimatedValue.textContent = formatCurrency(estimatedValue);

  const matchedRefs = rows.length;
  const inventoryRefs = inventory.rows.length;
  const historyRefs = history.referenceSet.size;
  const unmatchedHistory = [...history.referenceSet].filter((ref) => !inventory.referenceSet.has(ref)).length;
  const unmatchedInventory = [...inventory.referenceSet].filter((ref) => !history.referenceSet.has(ref)).length;

  elements.coverageText.textContent = `Cruce hecho sobre ${matchedRefs.toLocaleString('es-CO')} referencias. Histórico único: ${historyRefs.toLocaleString('es-CO')}. Inventario único: ${inventoryRefs.toLocaleString('es-CO')}. Sin cruce en histórico: ${unmatchedHistory}. Sin consumo histórico: ${unmatchedInventory}.`;

  if (priceCoverage) {
    const totalSavings = rows.reduce((sum, row) => sum + row.savingsImpact, 0);
    elements.savingsValue.textContent = formatCurrency(totalSavings);
    elements.savingsValue.classList.toggle('text-positive', totalSavings >= 0);
    elements.savingsValue.classList.toggle('text-negative', totalSavings < 0);

    elements.priceCoverageCard.style.display = '';
    const { total, exact, approximate, noMatch } = priceCoverage;
    const pct = (n) => (total > 0 ? ((n / total) * 100).toFixed(1) : '0.0');
    elements.priceCoverageText.textContent = `De ${total.toLocaleString('es-CO')} productos del inventario: ${exact.toLocaleString('es-CO')} (${pct(exact)}%) con precio oficial exacto, ${approximate.toLocaleString('es-CO')} (${pct(approximate)}%) con coincidencia aproximada (revisar) y ${noMatch.toLocaleString('es-CO')} (${pct(noMatch)}%) sin precio oficial (se usó el costo interno del inventario).`;
  } else {
    elements.savingsValue.textContent = '$0';
    elements.priceCoverageCard.style.display = 'none';
  }

  if (!rows.length) {
    elements.resultsBody.innerHTML = '<tr><td class="px-4 py-10 text-center text-slate-400" colspan="10">No hay referencias con sugerencia de compra para el periodo elegido.</td></tr>';
    return;
  }

  elements.resultsBody.innerHTML = rows.slice(0, 500).map((row) => `
    <tr class="transition hover:bg-white/5">
      <td data-label="Referencia" class="px-4 py-3 font-medium text-cyan-200">${escapeHtml(row.reference)}</td>
      <td data-label="Descripción" class="max-w-[26rem] px-4 py-3 text-slate-200">${escapeHtml(row.description)}</td>
      <td data-label="Promedio" class="px-4 py-3 tabular-nums">${formatDecimal(row.averageConsumption)}</td>
      <td data-label="Stock" class="px-4 py-3 tabular-nums">${formatDecimal(row.currentStock)}</td>
      <td data-label="A comprar" class="px-4 py-3 tabular-nums font-semibold text-emerald-300">${row.recommendedUnits.toLocaleString('es-CO')}</td>
      <td data-label="Costo usado" class="px-4 py-3 tabular-nums">${formatCurrency(row.cost)}</td>
      <td data-label="Proveedor sugerido" class="px-4 py-3 text-slate-200">${escapeHtml(row.suggestedProvider)}</td>
      <td data-label="Origen precio" class="px-4 py-3">${renderMatchBadge(row.matchType)}</td>
      <td data-label="Valor" class="px-4 py-3 tabular-nums">${formatCurrency(row.estimatedValue)}</td>
      <td data-label="Estado" class="px-4 py-3"><span class="badge badge-emerald">${escapeHtml(row.status)}</span></td>
    </tr>
  `).join('');

  if (rows.length > 500) {
    elements.resultsBody.insertAdjacentHTML('beforeend', '<tr><td class="px-4 py-4 text-center text-slate-400" colspan="10">Mostrando solo las primeras 500 filas del reporte.</td></tr>');
  }
}

function renderMatchBadge(matchType) {
  const variants = {
    exacta: { label: 'Exacta', className: 'badge-emerald' },
    aproximada: { label: 'Aproximada · revisar', className: 'badge-amber' },
    sin_match: { label: 'Sin precio oficial', className: 'badge-slate' },
    sin_lista: { label: 'Sin lista cargada', className: 'badge-slate' }
  };
  const variant = variants[matchType] || variants.sin_lista;
  return `<span class="badge ${variant.className}">${variant.label}</span>`;
}

function buildWorkbook(rows, monthsUsed, reportMonth, historyData, inventoryData, priceCoverage) {
  const totalUnits = rows.reduce((sum, row) => sum + row.recommendedUnits, 0);
  const totalValue = rows.reduce((sum, row) => sum + row.estimatedValue, 0);
  const weightedUnitCost = totalUnits > 0 ? roundNumber(totalValue / totalUnits, 2) : 0;
  const topItems = rows.slice(0, 10);

  const summaryAoa = [
    ['REPORTE DE COMPRA DE MEDICAMENTOS'],
    ['Generado', new Date()],
    ['Mes objetivo', formatMonth(reportMonth)],
    ['Meses analizados', monthsUsed.map(formatMonth).join(', ')],
    [],
    ['INDICADOR', 'VALOR'],
    ['Referencias sugeridas', rows.length],
    ['Unidades totales a comprar', totalUnits],
    ['Valor estimado total (COP)', totalValue],
    ['Costo unitario promedio ponderado (COP)', weightedUnitCost],
    ['Impacto ahorro / sobrecosto vs. costo interno (COP)', priceCoverage ? priceCoverage.totalSavingsImpact : 0],
    [],
    ['COBERTURA DE PRECIOS OFICIALES', priceCoverage ? '' : 'No se cargó lista de precios'],
    ['Productos con coincidencia exacta', priceCoverage ? priceCoverage.exact : 0],
    ['Productos con coincidencia aproximada (revisar)', priceCoverage ? priceCoverage.approximate : 0],
    ['Productos sin precio oficial (usa costo interno)', priceCoverage ? priceCoverage.noMatch : 0],
    [],
    ['TOP 10 REFERENCIAS POR VALOR ESTIMADO'],
    ['Referencia', 'Descripción', 'Unidades', 'Costo Unitario', 'Valor Estimado']
  ];

  for (const item of topItems) {
    summaryAoa.push([
      item.reference,
      item.description,
      item.recommendedUnits,
      item.cost,
      item.estimatedValue
    ]);
  }

  const summarySheet = XLSX.utils.aoa_to_sheet(summaryAoa);
  summarySheet['!merges'] = [
    XLSX.utils.decode_range('A1:E1'),
    XLSX.utils.decode_range('A13:E13'),
    XLSX.utils.decode_range('A18:E18')
  ];
  summarySheet['!cols'] = [
    { wch: 24 },
    { wch: 58 },
    { wch: 14 },
    { wch: 18 },
    { wch: 18 }
  ];

  const topDataStart = 20;
  for (let rowIndex = topDataStart; rowIndex < topDataStart + topItems.length; rowIndex += 1) {
    setCellNumberFormat(summarySheet, rowIndex, 3, '#,##0');
    setCellNumberFormat(summarySheet, rowIndex, 4, '$#,##0');
    setCellNumberFormat(summarySheet, rowIndex, 5, '$#,##0');
  }
  setCellNumberFormat(summarySheet, 2, 2, 'yyyy-mm-dd hh:mm');
  setCellNumberFormat(summarySheet, 8, 2, '#,##0');
  setCellNumberFormat(summarySheet, 9, 2, '$#,##0');
  setCellNumberFormat(summarySheet, 10, 2, '$#,##0.00');
  setCellNumberFormat(summarySheet, 11, 2, '$#,##0');
  setCellNumberFormat(summarySheet, 14, 2, '#,##0');
  setCellNumberFormat(summarySheet, 15, 2, '#,##0');
  setCellNumberFormat(summarySheet, 16, 2, '#,##0');

  const matchLabels = {
    exacta: 'Exacta',
    aproximada: 'Aproximada (revisar)',
    sin_match: 'Sin precio oficial',
    sin_lista: 'Sin lista cargada'
  };

  const detailRows = rows.map((row, index) => {
    const participation = totalValue > 0 ? row.estimatedValue / totalValue : 0;
    return {
      Ranking: index + 1,
      Referencia: row.reference,
      Descripcion: row.description,
      PromedioMensual: row.averageConsumption,
      StockActual: row.currentStock,
      UnidadesAComprar: row.recommendedUnits,
      CostoInterno: row.internalCost,
      CostoOficial: row.officialPrice ?? '',
      ProveedorSugerido: row.suggestedProvider,
      OrigenPrecio: matchLabels[row.matchType] || matchLabels.sin_lista,
      CostoUnitarioUsado: row.cost,
      DiferenciaVsInterno: row.priceDifference,
      ValorEstimado: row.estimatedValue,
      ImpactoAhorroSobrecosto: row.savingsImpact,
      ParticipacionValor: participation,
      Prioridad: getPriorityByValue(row.estimatedValue, totalValue),
      Estado: row.status
    };
  });

  const reportSheet = XLSX.utils.json_to_sheet(detailRows);
  reportSheet['!autofilter'] = { ref: `A1:Q${Math.max(2, detailRows.length + 1)}` };
  reportSheet['!cols'] = [
    { wch: 9 },
    { wch: 16 },
    { wch: 56 },
    { wch: 16 },
    { wch: 14 },
    { wch: 16 },
    { wch: 14 },
    { wch: 14 },
    { wch: 30 },
    { wch: 20 },
    { wch: 16 },
    { wch: 16 },
    { wch: 16 },
    { wch: 18 },
    { wch: 14 },
    { wch: 12 },
    { wch: 16 }
  ];

  for (let rowIndex = 2; rowIndex <= detailRows.length + 1; rowIndex += 1) {
    setCellNumberFormat(reportSheet, rowIndex, 1, '#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 4, '#,##0.00');
    setCellNumberFormat(reportSheet, rowIndex, 5, '#,##0.00');
    setCellNumberFormat(reportSheet, rowIndex, 6, '#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 7, '$#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 8, '$#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 11, '$#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 12, '$#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 13, '$#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 14, '$#,##0');
    setCellNumberFormat(reportSheet, rowIndex, 15, '0.00%');
  }

  const matchedRefs = rows.length;
  const unmatchedHistory = [...historyData.referenceSet].filter((ref) => !inventoryData.referenceSet.has(ref)).length;
  const unmatchedInventory = [...inventoryData.referenceSet].filter((ref) => !historyData.referenceSet.has(ref)).length;

  const traceAoa = [
    ['TRAZABILIDAD DEL REPORTE'],
    ['Mes objetivo', formatMonth(reportMonth)],
    ['Meses históricos usados', monthsUsed.map(formatMonth).join(', ')],
    ['Referencias históricas únicas', historyData.referenceSet.size],
    ['Referencias de inventario únicas', inventoryData.referenceSet.size],
    ['Referencias con compra sugerida', matchedRefs],
    ['Sin cruce en inventario', unmatchedHistory],
    ['Sin consumo histórico', unmatchedInventory]
  ];

  if (priceCoverage) {
    traceAoa.push(
      [],
      ['LISTA DE PRECIOS OFICIALES'],
      ['Productos cargados en la lista', priceCoverage.priceListProductCount ?? ''],
      ['Proveedores cargados en la lista', priceCoverage.priceListProviderCount ?? ''],
      ['Umbral mínimo de similitud para match aproximado', `${Math.round(FUZZY_MATCH_THRESHOLD * 100)}%`],
      ['Productos con coincidencia exacta', priceCoverage.exact],
      ['Productos con coincidencia aproximada (revisar)', priceCoverage.approximate],
      ['Productos sin precio oficial (usa costo interno)', priceCoverage.noMatch]
    );
  }

  const traceSheet = XLSX.utils.aoa_to_sheet(traceAoa);
  traceSheet['!merges'] = [XLSX.utils.decode_range('A1:B1')];
  if (priceCoverage) traceSheet['!merges'].push(XLSX.utils.decode_range('A10:B10'));
  traceSheet['!cols'] = [{ wch: 44 }, { wch: 38 }];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, summarySheet, 'Resumen_Ejecutivo');
  XLSX.utils.book_append_sheet(workbook, reportSheet, 'Detalle_Compra');

  if (priceCoverage && priceCoverage.approximate > 0) {
    const reviewRows = rows
      .filter((row) => row.matchType === 'aproximada')
      .map((row) => ({
        Referencia: row.reference,
        DescripcionInventario: row.description,
        DescripcionListaPrecios: row.matchedPriceDescription || '',
        Similitud: row.matchScore != null ? Math.round(row.matchScore * 100) / 100 : '',
        ProveedorSugerido: row.suggestedProvider,
        PrecioOficial: row.officialPrice,
        CostoInterno: row.internalCost,
        DiferenciaVsInterno: row.priceDifference
      }));
    const reviewSheet = XLSX.utils.json_to_sheet(reviewRows);
    reviewSheet['!cols'] = [
      { wch: 16 }, { wch: 56 }, { wch: 56 }, { wch: 12 }, { wch: 30 }, { wch: 16 }, { wch: 16 }, { wch: 16 }
    ];
    reviewSheet['!autofilter'] = { ref: `A1:H${Math.max(2, reviewRows.length + 1)}` };
    for (let rowIndex = 2; rowIndex <= reviewRows.length + 1; rowIndex += 1) {
      setCellNumberFormat(reviewSheet, rowIndex, 4, '0%');
      setCellNumberFormat(reviewSheet, rowIndex, 6, '$#,##0');
      setCellNumberFormat(reviewSheet, rowIndex, 7, '$#,##0');
      setCellNumberFormat(reviewSheet, rowIndex, 8, '$#,##0');
    }
    XLSX.utils.book_append_sheet(workbook, reviewSheet, 'Revisar_Precios');
  }

  XLSX.utils.book_append_sheet(workbook, traceSheet, 'Trazabilidad');
  return workbook;
}

function setCellNumberFormat(worksheet, row, column, format) {
  const cellRef = XLSX.utils.encode_cell({ r: row - 1, c: column - 1 });
  if (worksheet[cellRef]) {
    worksheet[cellRef].z = format;
  }
}

function getPriorityByValue(value, totalValue) {
  if (!totalValue || value <= 0) return 'Baja';
  const ratio = value / totalValue;
  if (ratio >= 0.07) return 'Alta';
  if (ratio >= 0.02) return 'Media';
  return 'Baja';
}

async function parseHistoryWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (!rows.length) throw new Error('El archivo histórico está vacío.');

  const headerIndex = findHeaderRow(rows, ['fecha', 'referencia', 'descripcion', 'salidas']);
  if (headerIndex < 0) throw new Error('No se encontraron encabezados en el histórico.');

  const headers = rows[headerIndex].map((value) => normalizeKey(value));
  const indexMap = createIndexMap(headers);
  const data = [];
  const months = [];
  const monthKeys = new Set();
  const referenceSet = new Set();

  for (let index = headerIndex + 1; index < rows.length; index += 1) {
    const row = rows[index];
    const reference = normalizeReference(row[indexMap.reference]);
    if (!reference) continue;

    const date = parseExcelDate(row[indexMap.date]);
    if (!date) continue;

    const exits = toNumber(row[indexMap.exits]);
    const monthKey = `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;

    data.push({ reference, date, monthKey, exits });
    if (!monthKeys.has(monthKey)) {
      monthKeys.add(monthKey);
      months.push(new Date(date.getFullYear(), date.getMonth(), 1));
    }
    referenceSet.add(reference);
  }

  if (!data.length) throw new Error('No se detectaron registros válidos en el histórico.');

  months.sort((a, b) => a - b);
  return { rows: data, months, referenceSet };
}

async function parseInventoryWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
  if (!rows.length) throw new Error('El archivo de inventario está vacío.');

  const headerIndex = findHeaderRow(rows, ['referencia', 'descripcion', 'stock', 'costo']);
  if (headerIndex < 0) throw new Error('No se encontraron encabezados en el inventario.');

  const headers = rows[headerIndex].map((value) => normalizeKey(value));
  const indexMap = createInventoryIndexMap(headers);
  const data = [];
  const referenceSet = new Set();

  for (let index = headerIndex + 1; index < rows.length; index += 1) {
    const row = rows[index];
    const reference = normalizeReference(row[indexMap.reference]);
    const description = cleanText(row[indexMap.description]);
    if (!reference) continue;

    data.push({ reference, description, stock: toNumber(row[indexMap.stock]), cost: toNumber(row[indexMap.cost]) });
    referenceSet.add(reference);
  }

  if (!data.length) throw new Error('No se detectaron productos válidos en el inventario.');

  return { rows: data, referenceSet };
}

async function parsePriceListWorkbook(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

  // La hoja con los precios puede no ser la primera del libro, así que se busca
  // en todas las hojas la que tenga los encabezados Proveedor / Producto / Valor.
  let sheetRows = null;
  let headerIndex = -1;
  for (const sheetName of workbook.SheetNames) {
    const candidateRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '' });
    const candidateHeaderIndex = findHeaderRow(candidateRows, ['proveedor', 'producto', 'valor']);
    if (candidateHeaderIndex >= 0) {
      sheetRows = candidateRows;
      headerIndex = candidateHeaderIndex;
      break;
    }
  }

  if (!sheetRows) throw new Error('No se encontró una hoja con las columnas Proveedor, Producto y Valor.');

  const headers = sheetRows[headerIndex].map((value) => normalizeKey(value));
  const providerIdx = headers.findIndex((header) => header === 'proveedor');
  const productIdx = headers.findIndex((header) => header === 'producto');
  const valueIdx = headers.findIndex((header) => header === 'valor');

  // Índice de precios: descripción normalizada -> { minValor, proveedor, originalDescription }
  const priceIndex = new Map();
  // Cubetas por primer token para acotar la búsqueda difusa (fuzzy) a candidatos relevantes.
  const buckets = new Map();
  const providers = new Set();

  let lastProvider = '';
  let productCount = 0;

  for (let index = headerIndex + 1; index < sheetRows.length; index += 1) {
    const row = sheetRows[index];
    const providerCell = cleanText(row[providerIdx]);
    if (providerCell) lastProvider = providerCell;

    const productText = cleanText(row[productIdx]);
    if (!productText) continue;

    const value = toNumber(row[valueIdx]);
    if (!value) continue;

    const normalized = normalizeDescription(productText);
    if (!normalized) continue;

    productCount += 1;
    if (lastProvider) providers.add(lastProvider);

    const existing = priceIndex.get(normalized);
    if (!existing || value < existing.minValor) {
      priceIndex.set(normalized, { minValor: value, proveedor: lastProvider || '—', originalDescription: productText });
    }

    const firstToken = normalized.split(' ')[0];
    if (!buckets.has(firstToken)) buckets.set(firstToken, []);
    if (!buckets.get(firstToken).includes(normalized)) buckets.get(firstToken).push(normalized);
  }

  if (!priceIndex.size) throw new Error('No se detectaron productos con precio válido en la lista de proveedores.');

  return { priceIndex, buckets, productCount: priceIndex.size, providerCount: providers.size };
}

function findOfficialPrice(description, priceListData) {
  const normalized = normalizeDescription(description);
  if (!normalized) return { matchType: 'sin_match' };

  const exact = priceListData.priceIndex.get(normalized);
  if (exact) {
    return {
      matchType: 'exacta',
      price: exact.minValor,
      provider: exact.proveedor,
      matchedDescription: exact.originalDescription
    };
  }

  // Sin coincidencia exacta: se busca por similitud de texto dentro de los productos
  // que comparten la primera palabra (normalmente el principio activo), y solo se
  // acepta la coincidencia si los números presentes (dosis, concentración, volumen)
  // son exactamente los mismos en ambos lados, para evitar cruces peligrosos entre
  // presentaciones distintas de un mismo medicamento.
  const firstToken = normalized.split(' ')[0];
  const candidates = priceListData.buckets.get(firstToken) || [];
  if (!candidates.length) return { matchType: 'sin_match' };

  const targetNumbers = extractNumbers(normalized);
  let bestScore = 0;
  let bestKey = null;

  for (const candidateKey of candidates) {
    if (!setsEqual(targetNumbers, extractNumbers(candidateKey))) continue;
    const score = diceCoefficient(normalized, candidateKey);
    if (score > bestScore) {
      bestScore = score;
      bestKey = candidateKey;
    }
  }

  if (bestKey && bestScore >= FUZZY_MATCH_THRESHOLD) {
    const match = priceListData.priceIndex.get(bestKey);
    return {
      matchType: 'aproximada',
      price: match.minValor,
      provider: match.proveedor,
      matchedDescription: match.originalDescription,
      score: bestScore
    };
  }

  return { matchType: 'sin_match' };
}

function normalizeDescription(value) {
  return cleanText(value)
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]+/g, ' ')
    .trim();
}

function extractNumbers(normalizedText) {
  const matches = normalizedText.match(/\d+(?:\.\d+)?/g);
  return new Set(matches || []);
}

function setsEqual(setA, setB) {
  if (setA.size !== setB.size) return false;
  for (const value of setA) {
    if (!setB.has(value)) return false;
  }
  return true;
}

function diceCoefficient(textA, textB) {
  const bigramsA = toBigramSet(textA);
  const bigramsB = toBigramSet(textB);
  if (!bigramsA.size || !bigramsB.size) return 0;
  let intersection = 0;
  for (const bigram of bigramsA) {
    if (bigramsB.has(bigram)) intersection += 1;
  }
  return (2 * intersection) / (bigramsA.size + bigramsB.size);
}

function toBigramSet(text) {
  const compact = text.replace(/\s+/g, '');
  const bigrams = new Set();
  for (let i = 0; i < compact.length - 1; i += 1) {
    bigrams.add(compact.slice(i, i + 2));
  }
  return bigrams;
}

function aggregateHistoricalConsumption(rows, monthsUsed) {
  const monthSet = new Set(monthsUsed.map((month) => `${month.getFullYear()}-${String(month.getMonth() + 1).padStart(2, '0')}`));
  const map = new Map();

  for (const row of rows) {
    if (!monthSet.has(row.monthKey)) continue;
    if (!map.has(row.reference)) {
      map.set(row.reference, { totalExits: 0 });
    }
    map.get(row.reference).totalExits += row.exits;
  }

  return map;
}

function createIndexMap(headers) {
  const find = (...candidates) => headers.findIndex((header) => candidates.some((candidate) => header === candidate || header.includes(candidate)));
  return {
    date: find('fecha'),
    reference: find('referencia', 'codigoreferencia', 'codigo'),
    description: find('descripcion'),
    exits: find('salidas', 'salida')
  };
}

function createInventoryIndexMap(headers) {
  const find = (...candidates) => headers.findIndex((header) => candidates.some((candidate) => header === candidate || header.includes(candidate)));
  return {
    reference: find('referencia', 'codigo'),
    description: find('descripcion'),
    stock: find('stock', 'existencia'),
    cost: find('costo', 'precio')
  };
}

function findHeaderRow(rows, requiredFields) {
  return rows.findIndex((row) => {
    const normalized = row.map((value) => normalizeKey(value));
    return requiredFields.every((field) => normalized.some((cell) => cell === field));
  });
}

function parseExcelDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (!parsed) return null;
    return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  if (typeof value === 'string') {
    const cleaned = value.trim();
    const match = cleaned.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (match) return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
    const parsed = new Date(cleaned);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }
  return null;
}

function parseMonthValue(value) {
  const [year, month] = value.split('-').map(Number);
  return new Date(year, month - 1, 1);
}

function getCurrentMonthValue() {
  const now = new Date();
  return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
}

function formatMonth(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function cleanText(value) {
  return String(value ?? '').trim();
}

function normalizeReference(value) {
  return cleanText(value)
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^A-Z0-9]/g, '');
}

function normalizeKey(value) {
  return cleanText(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function toNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  if (typeof value === 'string') {
    const raw = value.trim().replace(/[^0-9,.-]/g, '');
    const lastComma = raw.lastIndexOf(',');
    const lastDot = raw.lastIndexOf('.');
    let cleaned = raw;

    if (lastComma >= 0 && lastDot >= 0) {
      const decimalSeparator = lastComma > lastDot ? ',' : '.';
      const thousandsSeparator = decimalSeparator === ',' ? '.' : ',';
      cleaned = raw.replaceAll(thousandsSeparator, '').replace(decimalSeparator, '.');
    } else if (lastComma >= 0) {
      const decimals = raw.length - lastComma - 1;
      cleaned = decimals === 3 ? raw.replaceAll(',', '') : raw.replace(',', '.');
    } else if (lastDot >= 0) {
      const decimals = raw.length - lastDot - 1;
      cleaned = decimals === 3 ? raw.replaceAll('.', '') : raw;
    }

    const parsed = Number(cleaned);
    return Number.isFinite(parsed) ? parsed : 0;
  }
  return 0;
}

function roundNumber(value, decimals = 2) {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

function roundMoney(value) {
  return Math.round(value);
}

function formatDecimal(value) {
  return new Intl.NumberFormat('es-CO', { minimumFractionDigits: 0, maximumFractionDigits: 2 }).format(value);
}

function formatCurrency(value) {
  return new Intl.NumberFormat('es-CO', { style: 'currency', currency: 'COP', maximumFractionDigits: 0 }).format(value || 0);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function setStatus(message) {
  elements.statusBox.textContent = message;
}

function clearAllFields() {
  appState.historyData = null;
  appState.inventoryData = null;
  appState.priceListData = null;
  appState.latestWorkbook = null;

  elements.historyFile.value = '';
  elements.inventoryFile.value = '';
  elements.priceListFile.value = '';
  elements.reportMonth.value = getCurrentMonthValue();
  elements.periodOptions.forEach((option, index) => {
    const isDefault = index === 0;
    option.classList.toggle('is-active', isDefault);
    option.setAttribute('aria-pressed', String(isDefault));
  });

  elements.historyCount.textContent = '0 filas';
  elements.inventoryCount.textContent = '0 productos';
  elements.priceListCount.textContent = '0 productos';
  elements.resultCount.textContent = '0 compras';
  elements.unitsToBuy.textContent = '0';
  elements.estimatedValue.textContent = '$0';
  elements.monthsUsed.textContent = '0';
  elements.savingsValue.textContent = '$0';
  elements.coverageText.textContent = 'Carga los archivos para conocer cuántas referencias quedaron sin cruce entre histórico e inventario.';
  elements.priceCoverageCard.style.display = 'none';
  elements.priceCoverageText.textContent = '';
  elements.resultsBody.innerHTML = '<tr><td colspan="10">Sin datos todavía.</td></tr>';
  elements.downloadBtn.disabled = true;
  setStatus('Campos limpiados. Puedes cargar nuevos archivos.');
}

function getSelectedPeriodMonths() {
  const selectedOption = elements.periodOptions.find((option) => option.classList.contains('is-active'));
  return Number(selectedOption?.dataset.periodMonths || 3);
}