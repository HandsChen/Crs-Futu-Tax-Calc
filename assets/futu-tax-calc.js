
// ============================================================
// 0. GLOBAL STATE (multi-year architecture)
// ============================================================
const S = {
  files: { bill: [], div: [], pnl: [], crs: [] },
  // Raw data: all years, never mutated per-year
  raw: { trades: [], fundFlows: [], pnlRecords: [], divSummary: [] },
  // Parsed data: all years (dividends, interest etc.)
  parsed: { dividends: [], interest: [], fundDiv: [], fundTrades: [], noteIncome: [], fees: [], other: [] },
  // Global FIFO cost basis — built from ALL historical trades
  costBasis: {},
  // Multi-year results
  allYears: [],        // detected years, sorted e.g. ['2023','2024','2025']
  activeYear: null,     // currently viewing year (string) or null for "all"
  years: {},            // { '2025': { tax, capGains } }
  checkState: {},       // per-detail check tracking
};

// ============================================================
// 1. EXCHANGE RATES (央行月末中间价, 2022-2025)
// ============================================================
const RATES = {
  USD: {
    2022: {1:6.3646,2:6.3222,3:6.3507,4:6.4280,5:6.6607,6:6.6957,7:6.7371,8:6.8451,9:7.0488,10:7.1638,11:7.1211,12:6.9735},
    2023: {1:6.7750,2:6.9417,3:6.8873,4:6.9154,5:7.0635,6:7.1550,7:7.1444,8:7.1783,9:7.1778,10:7.1782,11:7.1381,12:7.0984},
    2024: {1:7.1057,2:7.1058,3:7.0964,4:7.1056,5:7.1095,6:7.1082,7:7.1400,8:7.1030,9:7.1055,10:7.0880,11:7.0789,12:7.0288},
    2025: {1:7.1698,2:7.1738,3:7.1782,4:7.2000,5:7.1848,6:7.1700,7:7.1400,8:7.1030,9:7.1055,10:7.0880,11:7.0789,12:7.0288}
  },
  HKD: {
    2022: {1:0.81783,2:0.81251,3:0.81590,4:0.82568,5:0.85506,6:0.85920,7:0.86152,8:0.87029,9:0.89775,10:0.91515,11:0.90947,12:0.89330},
    2023: {1:0.86957,2:0.88963,3:0.88266,4:0.88652,5:0.90226,6:0.91568,7:0.91209,8:0.91634,9:0.91715,10:0.91710,11:0.91296,12:0.90667},
    2024: {1:0.90709,2:0.90877,3:0.90723,4:0.90789,5:0.90933,6:0.90963,7:0.91400,8:0.91157,9:0.91298,10:0.91227,11:0.90760,12:0.90322},
    2025: {1:0.92298,2:0.91970,3:0.92030,4:0.92300,5:0.92170,6:0.91700,7:0.91400,8:0.91157,9:0.91298,10:0.91227,11:0.90760,12:0.90322}
  }
};

function getRate(ccy, dateStr) {
  const d = parseDate(dateStr);
  const yr = d.getFullYear();
  const mo = d.getMonth() + 1;
  const table = RATES[ccy];
  if (!table) return ccy === 'USD' ? 7.1 : 0.91;
  const yrData = table[yr];
  if (!yrData) { const yrs = Object.keys(table).sort(); return table[yrs[yrs.length-1]][mo] || (ccy==='USD'?7.1:0.91); }
  return yrData[mo] || (ccy === 'USD' ? 7.1 : 0.91);
}

function parseDate(str) {
  if (!str) return new Date();
  const s = String(str).trim();
  const m1 = s.match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (m1) return new Date(+m1[1], +m1[2]-1, +m1[3]);
  const m2 = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) return new Date(+m2[3], +m2[1]-1, +m2[2]);
  const m3 = s.match(/^(\d{4})(\d{2})(\d{2})/);
  if (m3) return new Date(+m3[1], +m3[2]-1, +m3[3]);
  return new Date();
}

function yearFromDate(str) {
  return parseDate(str).getFullYear();
}

// ============================================================
// 2. LOGGING
// ============================================================
function log(msg, cls) {
  const la = document.getElementById('log-area');
  const line = document.createElement('div');
  line.className = 'log-line' + (cls ? ' ' + cls : '');
  line.textContent = '  ' + msg;
  la.appendChild(line);
  la.scrollTop = la.scrollHeight;
}

function setProgress(pct) {
  document.getElementById('prog-bar').style.width = pct + '%';
}

// ============================================================
// 3. FILE PARSER
// ============================================================
function fuzzyMatch(headers, candidates) {
  for (const c of candidates) {
    for (const h of headers) {
      if (h && String(h).includes(c)) return h;
    }
  }
  return null;
}

function colIdx(headers, key) {
  for (let i = 0; i < headers.length; i++) {
    if (headers[i] && String(headers[i]).includes(key)) return i;
  }
  return -1;
}

function parseSheetToObjects(ws) {
  const range = XLSX.utils.decode_range(ws['!ref']);
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({r:range.s.r, c:c})];
    headers.push(cell ? String(cell.v || '').trim() : '');
  }
  const rows = [];
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const row = {};
    let hasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({r:r, c:c})];
      const val = cell ? cell.v : null;
      if (val !== null && val !== undefined && val !== '') hasData = true;
      row[headers[c]] = val;
    }
    if (hasData) rows.push(row);
  }
  return rows;
}

function parseBill(wb) {
  let trades = [];
  const tradeSheet = wb.Sheets['证券-交易流水'];
  if (tradeSheet) {
    const raw = parseSheetToObjects(tradeSheet);
    for (const r of raw) {
      const time = r['成交时间'];
      const sym = r['代码名称'];
      const mkt = r['交易所/市场'];
      const dir = r['方向'];
      const ccy = r['币种'];
      const qty = parseFloat(r['数量/面值'] || 0);
      const price = parseFloat(r['价格'] || 0);
      const amt = parseFloat(r['成交金额'] || 0);
      const fee = parseFloat(r['总费用'] || 0);
      const cat = r['品类'] || '';
      if (!time || !sym) continue;
      trades.push({time: String(time), symbol: String(sym), market: String(mkt||''), direction: String(dir||''),
        currency: String(ccy||''), qty, price, amount: amt, fee, category: String(cat) });
    }
    log(`交易流水: ${trades.length} 条`, 'log-ok');
  } else {
    log('未找到「证券-交易流水」sheet', 'log-err');
  }

  let fundFlows = [];
  const ffSheet = wb.Sheets['证券-资金进出'];
  if (ffSheet) {
    const raw = parseSheetToObjects(ffSheet);
    for (const r of raw) {
      const date = r['日期'];
      const type = r['类型'];
      const dir = r['方向'];
      const ccy = r['币种'];
      const amt = parseFloat(r['变动金额'] || 0);
      const note = r['备注'] || '';
      if (!date) continue;
      fundFlows.push({date: String(date), type: String(type||''), direction: String(dir||''),
        currency: String(ccy||''), amount: amt, note: String(note) });
    }
    log(`资金进出: ${fundFlows.length} 条`, 'log-ok');
  }

  return { trades, fundFlows };
}

function parsePnl(wb) {
  let records = [];
  const sheet = wb.Sheets['证券-盈亏记录-股票和期权交易'];
  if (!sheet) { log('未找到盈亏记录 sheet', 'log-err'); return records; }

  const range = XLSX.utils.decode_range(sheet['!ref']);
  const headers = [];
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({r:2, c:c})];
    headers.push(cell ? String(cell.v || '').trim() : '');
  }

  for (let r = 3; r <= range.e.r; r++) {
    const row = {};
    let hasData = false;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({r:r, c:c})];
      const val = cell ? cell.v : null;
      if (val !== null && val !== undefined && val !== '') hasData = true;
      row[headers[c]] = val;
    }
    if (!hasData) continue;
    const sym = row['代码名称'];
    if (!sym) continue;

    records.push({
      symbol: String(sym),
      market: String(row['市场/交易所'] || ''),
      date: String(row['交易日期'] || ''),
      currency: String(row['币种'] || ''),
      direction: String(row['方向'] || ''),
      qty: parseFloat(row['数量'] || 0),
      closePrice: parseFloat(row['平仓价格'] || 0),
      revenue: parseFloat(row['收入（含费用）'] || 0),
      cost: parseFloat(row['成本（含费用）'] || 0),
      gain: parseFloat(row['盈亏'] || 0),
    });
  }
  log(`盈亏记录: ${records.length} 条`, 'log-ok');
  return records;
}

function parseDivSummary(wb) {
  let records = [];
  const sheet = wb.Sheets['股息、利息及其他收入'];
  if (!sheet) { log('未找到股息利息汇总 sheet', 'log-err'); return records; }
  const rows = parseSheetToObjects(sheet);
  for (const r of rows) {
    records.push({
      year: r['年份'] || '',
      account: String(r['账户名称']||''),
      dividend: parseFloat(r['全年股息']||0),
      interest: parseFloat(r['全年利息']||0),
      other: parseFloat(r['全年其他收入']||0),
      currency: String(r['币种']||''),
    });
  }
  log(`股息利息汇总: ${records.length} 条`, 'log-ok');
  return records;
}

// ============================================================
// 4. CATEGORIZE FUND FLOWS
// ============================================================
function categorizeFundFlows(flows) {
  const result = {
    dividends: [], interest: [], fundDiv: [],
    fundTrades: [], noteIncome: [], fees: [], other: [],
  };

  for (const r of flows) {
    const note = (r.note || '').toUpperCase();
    const type = (r.type || '');

    if (note.includes('ACCOUNT UPGRADE') || note.includes('TRANSFER FROM') || note.includes('TRANSFER TO')) {
      continue;
    }

    if (note.includes('INTEREST') || type.includes('利息')) {
      result.interest.push(r);
      continue;
    }

    if (note.includes('ADR FEE') || note.includes('ADR费')) {
      result.fees.push(r);
      continue;
    }

    const isFund = note.includes('SUBSCRIPTION') || note.includes('REDEMPTION') || note.includes('申购') || note.includes('赎回');
    if (isFund) {
      result.fundTrades.push(r);
      continue;
    }

    const isFundDiv = (note.includes('FUND') && (note.includes('DIVIDEND') || note.includes('DISTRIBUTION'))) ||
                      note.includes('基金分红');
    if (isFundDiv) {
      result.fundDiv.push(r);
      continue;
    }

    const isNote = note.includes('NOTE') || note.includes('FTFC') || note.includes('SHARKFIN') ||
                   note.includes('结构化') || type.includes('结构化');
    if (isNote && r.direction === 'In') {
      result.noteIncome.push(r);
      continue;
    }

    const isDiv = note.includes('DIVIDEND') || note.includes('DIVIDENDS') || note.includes('F/D') ||
                  note.includes('I/D') || note.includes('S/D') || type.includes('公司行动');
    const isTax = note.includes('WITHHOLDING TAX') || note.includes('WITHHOLD TAX') || note.includes('预扣税') ||
                  note.includes('DIVIDEND TAX') || note.includes('DIVIDEND FEE');
    const isHandling = note.includes('HANDLING CHARGE') || note.includes('HANDING FEE') || note.includes('SCRIP CHARGE');

    if (isDiv) {
      if (isTax || isHandling) {
        if (r.direction === 'Out') {
          result.dividends.push(r);
        } else {
          result.fees.push(r);
        }
      } else if (r.direction === 'In') {
        result.dividends.push(r);
      } else {
        result.other.push(r);
      }
      continue;
    }

    if (isTax && r.direction === 'Out') {
      result.dividends.push(r);
      continue;
    }

    if (isHandling) {
      result.fees.push(r);
      continue;
    }

    result.other.push(r);
  }

  return result;
}

// ============================================================
// 5. BUILD DIVIDEND ITEMS
// ============================================================
function buildDividendItems(dividendFlows) {
  const items = [];
  const ins = dividendFlows.filter(r => r.direction === 'In');
  const outs = dividendFlows.filter(r => r.direction === 'Out');

  const usedOuts = new Set();
  for (const rIn of ins) {
    const note = (rIn.note || '').toUpperCase();
    const symMatch = note.match(/^(\w+)\s/);
    const symbol = symMatch ? symMatch[1] : 'UNKNOWN';
    const ccy = rIn.currency;
    const gross = Math.abs(rIn.amount);

    let taxAmt = 0;
    const inDate = parseDate(rIn.date);
    for (let i = 0; i < outs.length; i++) {
      if (usedOuts.has(i)) continue;
      const o = outs[i];
      if (o.currency !== ccy) continue;
      const outDate = parseDate(o.date);
      const dayDiff = Math.abs((inDate - outDate) / (1000*60*60*24));
      if (dayDiff <= 3) {
        const ratio = Math.abs(o.amount) / gross;
        if (ratio < 0.5) {
          taxAmt = Math.abs(o.amount);
          usedOuts.add(i);
          break;
        }
      }
    }

    items.push({
      symbol, date: rIn.date, currency: ccy,
      grossAmount: gross, taxAmount: taxAmt,
      netAmount: gross - taxAmt, note: rIn.note,
    });
  }

  for (let i = 0; i < outs.length; i++) {
    if (usedOuts.has(i)) continue;
    const o = outs[i];
    items.push({
      symbol: 'TAX', date: o.date, currency: o.currency,
      grossAmount: Math.abs(o.amount), taxAmount: Math.abs(o.amount), netAmount: 0,
      note: o.note,
    });
  }

  return items;
}

// ============================================================
// 6. BUILD FUND DIVIDEND ITEMS
// ============================================================
function buildFundDivItems(fundDivFlows) {
  return fundDivFlows.filter(r => r.direction === 'In').map(r => ({
    date: r.date,
    currency: r.currency,
    amount: Math.abs(r.amount),
    note: r.note,
  }));
}

// ============================================================
// 7. FIFO COST CALCULATION (global — uses ALL historical data)
// ============================================================
function buildFIFOCost(trades, pnlRecords) {
  const costBasis = {};

  function normalizeDate(dateStr) {
    if (!dateStr) return '';
    const s = String(dateStr).trim();
    const m1 = s.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
    if (m1) return m1[1] + String(m1[2]).padStart(2,'0') + String(m1[3]).padStart(2,'0');
    const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m2) return m2[3] + String(m2[1]).padStart(2,'0') + String(m2[2]).padStart(2,'0');
    const m3 = s.match(/^(\d{4})(\d{2})(\d{2})/);
    if (m3) return m3[1]+m3[2]+m3[3];
    return '';
  }

  // If we have PnL records, use their cost directly
  if (pnlRecords && pnlRecords.length > 0) {
    const pnlByTrade = {};
    for (const p of pnlRecords) {
      if (p.direction === '卖出平仓' || p.direction.includes('卖出') || p.direction.includes('SELL')) {
        const nDate = normalizeDate(p.date);
        const key = `${p.symbol}|${p.market}|${nDate}`;
        pnlByTrade[key] = { cost: Math.abs(p.cost), qty: Math.abs(p.qty), gain: p.gain };
      }
    }

    for (const t of trades) {
      const isSell = t.direction.includes('卖出') || t.direction.includes('SELL');
      if (!isSell) continue;
      const nDate = normalizeDate(t.time);
      const key = `${t.symbol}|${t.market}|${nDate}`;
      const pnl = pnlByTrade[key];
      if (pnl) {
        const symKey = t.symbol + '|' + t.market;
        if (!costBasis[symKey]) costBasis[symKey] = [];
        costBasis[symKey].push({
          qty: Math.abs(t.qty), costPerUnit: pnl.cost / Math.abs(p.qty), totalCost: pnl.cost, gain: pnl.gain,
          sellDate: t.time, sellPrice: t.price, sellAmt: Math.abs(t.amount),
        });
      }
    }
    return costBasis;
  }

  // No PnL records - use pure FIFO from ALL trade history (cross-year buys included)
  const buyQueues = {};
  const sells = [];

  // Sort all trades by time (chronologically)
  const sorted = [...trades].sort((a, b) => String(a.time).localeCompare(String(b.time)));

  for (const t of sorted) {
    const symKey = t.symbol + '|' + t.market;
    const dir = String(t.direction || '');
    const isBuy = dir.includes('买入') || dir.includes('BUY');
    const isTransferIn = dir.includes('转仓转入') || dir.includes('TRANSFER IN') || dir.includes('转入');
    const isSell = dir.includes('卖出') || dir.includes('SELL');
    const absQty = Math.abs(t.qty);

    if (isBuy || isTransferIn) {
      // 转仓转入：使用转仓当天股价作为成本基准
      if (!buyQueues[symKey]) buyQueues[symKey] = [];
      buyQueues[symKey].push({ qty: absQty, price: t.price, isTransfer: isTransferIn });
    } else if (isSell) {
      sells.push({ ...t, symKey, absQty });
    }
  }

  // Track estimated costs for UI warning
  const estimatedCosts = [];

  for (const s of sells) {
    const queue = buyQueues[s.symKey] || [];
    let remaining = s.absQty;
    let totalCost = 0;
    while (remaining > 0 && queue.length > 0) {
      const batch = queue[0];
      const used = Math.min(remaining, batch.qty);
      totalCost += used * batch.price;
      batch.qty -= used;
      remaining -= used;
      if (batch.qty <= 0) queue.shift();
    }

    // FIFO couldn't match all shares — try PnL, then estimate
    if (remaining > 0) {
      let estimated = false;
      // Try PnL record for this sell (fuzzy: same symbol, date ±3 days)
      if (pnlRecords && pnlRecords.length > 0) {
        const sellDate = parseDate(s.time);
        for (const p of pnlRecords) {
          if (p.symbol !== s.symbol) continue;
          if (!p.direction.includes('卖出') && !p.direction.includes('SELL') && p.direction !== '卖出平仓') continue;
          const pDate = parseDate(p.date);
          const diffDays = Math.abs((sellDate - pDate) / (1000*60*60*24));
          if (diffDays <= 3) {
            const unmatchedQty = remaining;
            totalCost += unmatchedQty * (Math.abs(p.cost) / Math.max(Math.abs(p.qty), 1));
            remaining = 0;
            estimated = false; // Have PnL cost, not estimated
            break;
          }
        }
      }

      if (remaining > 0) {
        // Still can't find cost — estimate using sell price * 0 (conservative: cost=0)
        // Log a warning; UI will show caution
        estimatedCosts.push({
          symbol: s.symbol,
          date: s.time,
          unmatchedQty: remaining,
          estimatedCost: 0,
        });
        remaining = 0; // Don't fail — just use cost=0 and warn
      }
    }

    if (!costBasis[s.symKey]) costBasis[s.symKey] = [];
    const costPerUnit = s.absQty > 0 ? totalCost / s.absQty : 0;
    const isEstimated = estimatedCosts.some(e => e.symbol === s.symbol && e.date === s.time);
    costBasis[s.symKey].push({
      qty: s.absQty, costPerUnit, totalCost,
      gain: Math.abs(s.amount) - totalCost,
      sellDate: s.time, sellPrice: s.price, sellAmt: Math.abs(s.amount),
      estimated: isEstimated,
    });
  }

  // Store estimated costs globally for UI warning
  if (estimatedCosts.length > 0) {
    S.estimatedCosts = (S.estimatedCosts || []).concat(estimatedCosts);

// --- State for user cost overrides ---
S.costOverrides = S.costOverrides || {};
  }

  return costBasis;
}

// ============================================================
// 8. CLASSIFY CAPITAL GAINS
// ============================================================
function classifyCapitalGains(trades, costBasis) {
  const usTrades = [];
  const hkTrades = [];
  const fundTrades = [];

  for (const t of trades) {
    const isSell = t.direction.includes('卖出') || t.direction.includes('SELL');
    if (!isSell) continue;

    const symKey = t.symbol + '|' + t.market;
    const mkt = (t.market || '').toUpperCase();
    const cat = (t.category || '');

    if (cat === '基金' || cat.includes('基金') || cat === 'FUND') {
      const costs = costBasis[symKey] || [];
      const match = costs.find(c => c.sellDate === t.time);
      fundTrades.push({
        symbol: t.symbol, date: t.time, currency: t.currency,
        qty: Math.abs(t.qty), sellAmt: Math.abs(t.amount),
        cost: match ? match.totalCost : 0,
        gain: match ? match.gain : 0,
      });
      continue;
    }

    const costs = costBasis[symKey] || [];
    const match = costs.find(c => c.sellDate === t.time);
    const entry = {
      symbol: t.symbol, date: t.time, currency: t.currency,
      qty: Math.abs(t.qty), sellAmt: Math.abs(t.amount),
      cost: match ? match.totalCost : 0,
      gain: match ? match.gain : (Math.abs(t.amount) - (match ? match.totalCost : 0)),
    };

    if (mkt === 'US' || mkt.includes('US')) {
      usTrades.push(entry);
    } else {
      hkTrades.push(entry);
    }
  }

  return { usTrades, hkTrades, fundTrades };
}

// ============================================================
// 9. TAX ENGINE (unchanged — single year input)
// ============================================================
function calculateTax(data) {
  const { dividends, interest, fundDiv, noteIncome } = data;
  const { usTrades, hkTrades, fundTrades: capFundTrades } = data.capGains;

  // ---- US ----
  const usDiv = dividends.filter(d => d.currency === 'USD');
  const usInt = interest.filter(d => d.currency === 'USD');

  let usDivGross = 0, usDivTax = 0;
  const usDivDetail = usDiv.map(d => {
    const rate = d.grossAmount > 0 ? getRate('USD', d.date) : getRate('USD', null);
    const grossCNY = d.grossAmount * rate;
    const taxCNY = d.taxAmount * rate;
    usDivGross += grossCNY; usDivTax += taxCNY;
    return { ...d, rate, grossCNY, taxCNY };
  });

  let usIntGross = 0;
  const usIntDetail = usInt.map(d => {
    const rate = getRate('USD', d.date);
    const cny = Math.abs(d.amount) * rate;
    usIntGross += cny;
    return { ...d, rate, cny };
  });

  let usCapNet = 0;
  const usCapDetail = usTrades.map(t => {
    const rate = getRate('USD', t.date);
    const gainCNY = t.gain * rate;
    usCapNet += gainCNY;
    return { ...t, rate, gainCNY };
  });

  const usTaxable = usDivGross + usIntGross + Math.max(0, usCapNet);
  const usCredit = usDivTax;
  const usTaxBefore = (usDivGross + usIntGross + Math.max(0, usCapNet)) * 0.2;
  const usTax = Math.max(0, usTaxBefore - usCredit);

  // ---- HK ----
  const hkDiv = dividends.filter(d => d.currency === 'HKD');
  const hkInt = interest.filter(d => d.currency === 'HKD');

  let hkDivGross = 0, hkDivTax = 0;
  const hkDivDetail = hkDiv.map(d => {
    const rate = d.grossAmount > 0 ? getRate('HKD', d.date) : getRate('HKD', null);
    const grossCNY = d.grossAmount * rate;
    const taxCNY = d.taxAmount * rate;
    hkDivGross += grossCNY; hkDivTax += taxCNY;
    return { ...d, rate, grossCNY, taxCNY };
  });

  let hkIntGross = 0;
  const hkIntDetail = hkInt.map(d => {
    const rate = getRate('HKD', d.date);
    const cny = Math.abs(d.amount) * rate;
    hkIntGross += cny;
    return { ...d, rate, cny };
  });

  let hkCapNet = 0;
  const hkCapDetail = hkTrades.map(t => {
    const rate = getRate('HKD', t.date);
    const gainCNY = t.gain * rate;
    hkCapNet += gainCNY;
    return { ...t, rate, gainCNY };
  });

  // Fund capital gains
  let fundCapNet = 0;
  const fundCapDetail = capFundTrades.map(t => {
    const ccy = t.currency;
    const rate = ccy === 'USD' ? getRate('USD', t.date) : getRate('HKD', t.date);
    const gainCNY = t.gain * rate;
    fundCapNet += gainCNY;
    return { ...t, rate, gainCNY };
  });

  // Fund dividends
  let fundDivNet = 0;
  const fundDivDetail = fundDiv.map(d => {
    const rate = d.currency === 'USD' ? getRate('USD', d.date) : getRate('HKD', d.date);
    const cny = d.amount * rate;
    fundDivNet += cny;
    return { ...d, rate, cny };
  });

  // Note income
  let noteNet = 0;
  const noteDetail = noteIncome.map(d => {
    const rate = d.currency === 'USD' ? getRate('USD', d.date) : getRate('HKD', d.date);
    const cny = Math.abs(d.amount) * rate;
    noteNet += cny;
    return { ...d, rate, cny };
  });

  const hkAllCapNet = hkCapNet + fundCapNet;
  const hkTaxable = hkDivGross + hkIntGross + Math.max(0, hkAllCapNet) + fundDivNet + noteNet;
  const hkCredit = hkDivTax;
  const hkTaxBefore = (hkDivGross + hkIntGross + Math.max(0, hkAllCapNet) + fundDivNet + noteNet) * 0.2;
  const hkTax = Math.max(0, hkTaxBefore - hkCredit);

  return {
    us: {
      div: { gross: usDivGross, tax: usDivTax, detail: usDivDetail, taxDue: usDivGross * 0.2, netTax: Math.max(0, usDivGross * 0.2 - usDivTax) },
      int: { gross: usIntGross, detail: usIntDetail, taxDue: usIntGross * 0.2 },
      cap: { net: usCapNet, detail: usCapDetail, taxDue: Math.max(0, usCapNet * 0.2) },
      totalTaxable: usTaxable, credit: usCredit, taxDue: usTax,
    },
    hk: {
      div: { gross: hkDivGross, tax: hkDivTax, detail: hkDivDetail, taxDue: hkDivGross * 0.2, netTax: Math.max(0, hkDivGross * 0.2 - hkDivTax) },
      int: { gross: hkIntGross, detail: hkIntDetail, taxDue: hkIntGross * 0.2 },
      cap: { net: hkCapNet, detail: hkCapDetail, taxDue: Math.max(0, hkCapNet * 0.2) },
      fundCap: { net: fundCapNet, detail: fundCapDetail, taxDue: Math.max(0, fundCapNet * 0.2) },
      fundDiv: { gross: fundDivNet, detail: fundDivDetail, taxDue: fundDivNet * 0.2 },
      note: { gross: noteNet, detail: noteDetail, taxDue: noteNet * 0.2 },
      totalTaxable: hkTaxable, credit: hkCredit, taxDue: hkTax,
    },
    totalTaxable: usTaxable + hkTaxable,
    totalCredit: usCredit + hkCredit,
    totalTax: usTax + hkTax,
  };
}

// ============================================================
// 10. YEAR DETECTION
// ============================================================
function detectYears() {
  const years = new Set();
  // From trades
  for (const t of S.raw.trades) {
    const yr = yearFromDate(t.time);
    if (yr >= 2020 && yr <= 2030) years.add(yr);
  }
  // From fund flows
  for (const f of S.raw.fundFlows) {
    const yr = yearFromDate(f.date);
    if (yr >= 2020 && yr <= 2030) years.add(yr);
  }
  // From parsed dividends
  for (const d of S.parsed.dividends) {
    const yr = yearFromDate(d.date);
    if (yr >= 2020 && yr <= 2030) years.add(yr);
  }
  // From interest
  for (const d of S.parsed.interest) {
    const yr = yearFromDate(d.date);
    if (yr >= 2020 && yr <= 2030) years.add(yr);
  }
  // From fund dividends
  for (const d of S.parsed.fundDiv) {
    const yr = yearFromDate(d.date);
    if (yr >= 2020 && yr <= 2030) years.add(yr);
  }
  // From note income
  for (const d of S.parsed.noteIncome) {
    const yr = yearFromDate(d.date);
    if (yr >= 2020 && yr <= 2030) years.add(yr);
  }
  S.allYears = Array.from(years).sort((a,b) => a - b);
  log(`检测到年度: ${S.allYears.join(', ')}`, 'log-ok');
}

// Helper: filter an array of items by year
function filterByYear(items, getDateFn, year) {
  return items.filter(item => yearFromDate(getDateFn(item)) === year);
}

// ============================================================
// 11. MAIN PROCESSING PIPELINE (multi-year)
// ============================================================
async function processAll() {
  setProgress(0);
  const la = document.getElementById('log-area');
  la.innerHTML = '';
  la.style.display = 'block';
  document.getElementById('progress-area').style.display = 'block';

  try {
    // --- Parse ALL bill files ---
    log('正在解析年度账单（多文件）...');
    setProgress(5);
    let allTrades = [];
    let allFundFlows = [];
    for (let i = 0; i < S.files.bill.length; i++) {
      log(`  解析: ${S.files.bill[i].name}`);
      const billWb = XLSX.read(await S.files.bill[i].arrayBuffer(), { type: 'array' });
      const { trades, fundFlows } = parseBill(billWb);
      allTrades = allTrades.concat(trades);
      allFundFlows = allFundFlows.concat(fundFlows);
    }
    S.raw.trades = allTrades;
    S.raw.fundFlows = allFundFlows;
    log(`  共 ${allTrades.length} 条交易`, 'log-ok');
    setProgress(25);

    // Keep legacy variable names for downstream code
    const trades = allTrades;
    const fundFlows = allFundFlows;

    // --- Parse ALL pnl files ---
    let allPnl = [];
    if (S.files.pnl && S.files.pnl.length > 0) {
      log('正在解析盈亏记录（多文件）...');
      for (let i = 0; i < S.files.pnl.length; i++) {
        log(`  解析: ${S.files.pnl[i].name}`);
        const pnlWb = XLSX.read(await S.files.pnl[i].arrayBuffer(), { type: 'array' });
        const recs = parsePnl(pnlWb);
        allPnl = allPnl.concat(recs);
      }
      S.raw.pnlRecords = allPnl;
      log(`  共 ${allPnl.length} 条盈亏记录`, 'log-ok');
    } else {
      log('⚠ 未提供盈亏记录，FIFO成本将从交易流水推算', 'log-err');
    }
    setProgress(40);

    // --- Parse ALL div files ---
    let allDiv = [];
    if (S.files.div && S.files.div.length > 0) {
      log('正在解析股息利息汇总（多文件）...');
      for (let i = 0; i < S.files.div.length; i++) {
        log(`  解析: ${S.files.div[i].name}`);
        const divWb = XLSX.read(await S.files.div[i].arrayBuffer(), { type: 'array' });
        const recs = parseDivSummary(divWb);
        allDiv = allDiv.concat(recs);
      }
      S.raw.divSummary = allDiv;
    }
    setProgress(50);

    // --- Categorize fund flows (global, all years) ---
    log('正在分类资金流水...');
    const cats = categorizeFundFlows(fundFlows);
    S.parsed.dividends = buildDividendItems(cats.dividends);
    S.parsed.interest = cats.interest;
    S.parsed.fundDiv = buildFundDivItems(cats.fundDiv);
    S.parsed.fundTrades = cats.fundTrades;
    S.parsed.noteIncome = cats.noteIncome;
    S.parsed.fees = cats.fees;
    S.parsed.other = cats.other;
    log(`  股息 ${S.parsed.dividends.length}笔 · 利息 ${S.parsed.interest.length}笔 · 基金分红 ${S.parsed.fundDiv.length}笔 · 基金交易 ${S.parsed.fundTrades.length}笔 · 票据 ${S.parsed.noteIncome.length}笔`, 'log-ok');
    setProgress(60);

    // --- Detect years ---
    detectYears();
    if (S.allYears.length === 0) {
      log('⚠ 无法从数据中检测年份，默认按2025年处理', 'log-err');
      S.allYears = [2025];
    }
    setProgress(65);

    // --- FIFO (global — all years) ---
    log('正在计算全局FIFO成本（使用所有年度数据）...');
    S.costBasis = buildFIFOCost(trades, S.raw.pnlRecords);
    setProgress(75);

    // --- Per-year calculation ---
    S.years = {};
    S.checkState = {};
    let ckIdCounter = 0;

    for (const yr of S.allYears) {
      const yearStr = String(yr);
      log(`  计算 ${yearStr}年度...`);

      // Filter trades for this year
      const yearTrades = filterByYear(trades, t => t.time, yr);
      const yearCapGains = classifyCapitalGains(yearTrades, S.costBasis);

      // Filter dividends, interest, fundDiv, noteIncome for this year
      const yearDividends = filterByYear(S.parsed.dividends, d => d.date, yr);
      const yearInterest = filterByYear(S.parsed.interest, d => d.date, yr);
      const yearFundDiv = filterByYear(S.parsed.fundDiv, d => d.date, yr);
      const yearNoteIncome = filterByYear(S.parsed.noteIncome, d => d.date, yr);

      // Calculate tax for this year
      const yearTax = calculateTax({
        dividends: yearDividends,
        interest: yearInterest,
        fundDiv: yearFundDiv,
        noteIncome: yearNoteIncome,
        capGains: yearCapGains,
      });

      S.years[yearStr] = { tax: yearTax, capGains: yearCapGains };

      // Pre-populate checkState (use year prefix to avoid collisions)
      // We'll do this in render time to keep it simple

      const usCap = yearTax.us.cap.detail.length;
      const hkCap = yearTax.hk.cap.detail.length;
      const usDiv = yearTax.us.div.detail.length;
      const hkDiv = yearTax.hk.div.detail.length;
      const active = usCap + hkCap + usDiv + hkDiv;
      log(`    ${yearStr}：应税 ¥${yearTax.totalTaxable.toFixed(0)} · 应缴 ¥${yearTax.totalTax.toFixed(0)} · 交易 ${yearTrades.filter(t=>t.direction.includes('卖')||t.direction.includes('SELL')).length}笔`);
    }
    setProgress(90);

    // --- Set active year (most recent) ---
    S.activeYear = String(S.allYears[S.allYears.length - 1]);

    // --- Render ---
    log('正在渲染结果...');
    renderYearBar();
    renderResults();
    setProgress(100);
    log('✅ 计算完成！', 'log-ok');

    updateSteps(3);
  } catch (err) {
    log('❌ 处理出错: ' + err.message, 'log-err');
    console.error(err);
  }
}

// ============================================================
// 12. RENDER YEAR BAR
// ============================================================
function renderYearBar() {
  const bar = document.getElementById('year-bar');
  bar.style.display = 'flex';
  // Clear existing tabs (keep the label)
  const existingTabs = bar.querySelectorAll('.year-tab');
  existingTabs.forEach(t => t.remove());

  for (const yr of S.allYears) {
    const btn = document.createElement('button');
    btn.className = 'year-tab';
    btn.textContent = yr + ' 年';
    btn.dataset.year = String(yr);
    if (S.activeYear === String(yr)) btn.classList.add('active');
    btn.addEventListener('click', () => switchYear(String(yr)));
    bar.appendChild(btn);
  }
}

function switchYear(yearStr) {
  S.activeYear = yearStr;

  // Update year tabs
  document.querySelectorAll('#year-bar .year-tab').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.year === yearStr);
  });

  // Re-render
  renderResults();

  // Update export button label
  document.getElementById('btnExport').textContent = `📥 下载 ${yearStr}年度报税底稿 (.md)`;
}

// ============================================================
// 13. UI RENDERER (multi-year aware)
// ============================================================
function getCurrentTax() {
  if (!S.activeYear || !S.years[S.activeYear]) return null;
  return S.years[S.activeYear].tax;
}

function renderResults() {
  const T = getCurrentTax();
  if (!T) return;

  document.getElementById('empty-state').style.display = 'none';
  document.getElementById('result-area').style.display = 'block';

  // Show cost warning if any costs were estimated
  const costWarn = document.getElementById('cost-warning');
  const costWarnSymbols = document.getElementById('cost-warn-symbols');
  if (S.estimatedCosts && S.estimatedCosts.length > 0) {
    const symbols = [...new Set(S.estimatedCosts.map(e => e.symbol))].join(', ');
    costWarnSymbols.textContent = symbols;
    costWarn.style.display = 'block';
  } else {
    costWarn.style.display = 'none';
  }

  const yearLabel = S.activeYear + '年度';

  // Stat cards
  const statRow = document.getElementById('stat-row');
  const checkCount = Object.keys(S.checkState).filter(k => k.startsWith(S.activeYear + '-')).length;
  const detailCount = (T.us.div.detail||[]).length + (T.us.int.detail||[]).length + (T.us.cap.detail||[]).length
    + (T.hk.div.detail||[]).length + (T.hk.int.detail||[]).length + (T.hk.cap.detail||[]).length
    + (T.hk.fundCap.detail||[]).length + (T.hk.fundDiv.detail||[]).length + (T.hk.note.detail||[]).length;

  statRow.innerHTML = `
    <div class="stat stat-pri"><div class="value">¥${T.totalTax.toFixed(0)}</div><div class="label">应缴税额</div><div class="sub">${yearLabel}</div></div>
    <div class="stat stat-war"><div class="value">¥${T.totalTaxable.toFixed(0)}</div><div class="label">应税所得</div><div class="sub">美国 + 香港</div></div>
    <div class="stat stat-suc"><div class="value">¥${T.totalCredit.toFixed(0)}</div><div class="label">已缴税抵扣</div><div class="sub">源泉税</div></div>
    <div class="stat"><div class="value" style="color:var(--gray-500);">${checkCount}/${detailCount}</div><div class="label">核对进度</div><div class="sub">点击税目展开核对</div></div>
  `;

  // Card subtitle
  const cardTitle = document.querySelector('#result-area .card-title');
  if (cardTitle) {
    cardTitle.innerHTML = `📋 ${yearLabel} — 分国分项明细 &amp; 逐笔核对`;
  }

  // Render tab contents
  renderUSTab(T);
  renderHKTab(T);
  renderRateTab();
}

// ─── US Tab ───
function renderUSTab(T) {
  const us = T.us;
  const rows = [];
  const prefix = S.activeYear + '-';

  if (us.div.detail.length > 0) {
    const detailId = prefix + 'us-div';
    S.checkState[detailId] = { total: us.div.detail.length, checked: 0 };
    const checkBadge = `<span class="badge badge-us" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${us.div.detail.length} 已核</span>`;
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-us">股息红利</span></td>
        <td class="cc">${us.div.detail.length} 笔</td>
        <td>USD ${us.div.detail.reduce((s,d)=>s+d.grossAmount,0).toFixed(2)}</td>
        <td class="num">${us.div.gross.toFixed(2)}</td>
        <td class="num">${us.div.tax.toFixed(2)}</td>
        <td class="num pos">${us.div.netTax.toFixed(2)}</td>
        <td style="text-align:center;">${checkBadge}</td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>#</th><th>股票</th><th>日期</th><th class="num">毛额</th><th class="num">源泉税</th><th class="num">汇率</th><th class="num">折CNY毛额</th><th class="num">折CNY已缴税</th><th class="chk-col">✓</th></tr></thead>
              <tbody>
                ${us.div.detail.map((d,i) => `
                  <tr>
                    <td>${i+1}</td><td>${d.symbol}</td><td>${d.date}</td>
                    <td class="num">${d.grossAmount.toFixed(2)}</td>
                    <td class="num">${d.taxAmount.toFixed(2)}</td>
                    <td class="num">${d.rate.toFixed(4)}</td>
                    <td class="num">${d.grossCNY.toFixed(2)}</td>
                    <td class="num">${d.taxCNY.toFixed(2)}</td>
                    <td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td>
                  </tr>`).join('')}
              </tbody>
              <tfoot><tr class="dft"><td colspan="3">合计</td><td class="num">${us.div.detail.reduce((s,d)=>s+d.grossAmount,0).toFixed(2)}</td><td class="num">${us.div.detail.reduce((s,d)=>s+d.taxAmount,0).toFixed(2)}</td><td></td><td class="num">${us.div.gross.toFixed(2)}</td><td class="num">${us.div.tax.toFixed(2)}</td><td></td></tr></tfoot>
            </table>
          </div>
        </td>
      </tr>`);
  }

  if (us.int.detail.length > 0) {
    const detailId = prefix + 'us-int';
    S.checkState[detailId] = { total: us.int.detail.length, checked: 0 };
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-us">利息所得</span></td>
        <td class="cc">${us.int.detail.length} 笔</td>
        <td>USD ${us.int.detail.reduce((s,d)=>s+Math.abs(d.amount),0).toFixed(2)}</td>
        <td class="num">${us.int.gross.toFixed(2)}</td><td class="num">0</td>
        <td class="num pos">${us.int.taxDue.toFixed(2)}</td>
        <td style="text-align:center;"><span class="badge badge-us" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${us.int.detail.length} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>来源</th><th>日期</th><th class="num">金额</th><th class="num">汇率</th><th class="num">折CNY</th><th class="chk-col">✓</th></tr></thead>
              <tbody>
                ${us.int.detail.map(d => `
                  <tr><td>账户余额利息</td><td>${d.date}</td><td class="num">${Math.abs(d.amount).toFixed(2)}</td><td class="num">${d.rate.toFixed(4)}</td><td class="num">${d.cny.toFixed(2)}</td><td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td></tr>`).join('')}
              </tbody>
            </table>
          </div>
        </td>
      </tr>`);
  }

  if (us.cap.detail.length > 0) {
    const detailId = prefix + 'us-cap';
    S.checkState[detailId] = { total: us.cap.detail.length, checked: 0 };
    const netLabel = us.cap.net < 0 ? '净亏损' : '净盈利';
    const netClass = us.cap.net < 0 ? 'neg' : 'pos';
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-us">财产转让</span></td>
        <td class="cc">${us.cap.detail.length} 笔</td>
        <td>${netLabel}</td>
        <td class="num ${netClass}">${us.cap.net.toFixed(0)}</td><td class="num">0</td>
        <td class="num ${us.cap.taxDue > 0 ? 'pos' : ''}">${us.cap.taxDue.toFixed(0)}</td>
        <td style="text-align:center;"><span class="badge badge-us" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${us.cap.detail.length} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>#</th><th>股票</th><th>日期</th><th class="num">数量</th><th class="num">卖价</th><th class="num">FIFO成本</th><th class="num">盈亏</th><th class="num">汇率</th><th class="num">折CNY</th><th class="chk-col">✓</th></tr></thead>
              <tbody>
                ${us.cap.detail.map((d,i) => `
                  <tr>
                    <td>${i+1}</td><td>${d.symbol}${d.estimated ? ' ⚠️':''}</td><td>${String(d.date).substring(0,10)}</td>
                    <td class="num">${d.qty}</td><td class="num">${(d.sellAmt/d.qty||0).toFixed(2)}</td>
                    <td class="num cost-cell"><input class="cost-edit" type="text" value="${(d.cost/d.qty||0).toFixed(2)}" data-symbol="${d.symbol}" data-date="${String(d.date).substring(0,10)}" data-market="us" data-orig-value="${(d.cost/d.qty||0).toFixed(2)}" onchange="onCostEdit(this)" /></td>
                    <td class="num gain-cell ${d.gain>=0?'pos':'neg'}">${d.gain>=0?'+':''}${d.gain.toFixed(0)}</td>
                    <td class="num">${d.rate.toFixed(4)}</td>
                    <td class="num gain-cny-cell ${d.gainCNY>=0?'pos':'neg'}">${d.gainCNY>=0?'+':''}${d.gainCNY.toFixed(0)}</td>
                    <td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td>
                  </tr>`).join('')}
              </tbody>
              <tfoot><tr class="dft"><td colspan="6">合计</td><td class="num ${us.cap.net>=0?'pos':'neg'}">${us.cap.net>=0?'+':''}${us.cap.net.toFixed(0)}</td><td></td><td class="num">${us.cap.net.toFixed(0)}</td><td></td></tr></tfoot>
            </table>
            ${us.cap.net < -0.01 ? '<div class="dn">💡 美股财产转让净亏损，应税=0。亏损可同国同年同税目内互抵。</div>' : ''}
          </div>
        </td>
      </tr>`);
  }

  rows.push(`
    <tr style="font-weight:700;background:var(--primary-light);">
      <td colspan="3">🇺🇸 美国小计</td><td></td>
      <td class="num">${us.totalTaxable.toFixed(2)}</td>
      <td class="num">${us.credit.toFixed(2)}</td>
      <td class="num" style="color:var(--primary);">${us.taxDue.toFixed(2)}</td><td></td>
    </tr>`);

  document.getElementById('tab-us').innerHTML = `
    <table class="summary-table">
      <thead><tr><th style="width:36px;"></th><th>税目</th><th>笔数</th><th>原币金额</th><th class="num">折合CNY</th><th class="num">境外已缴税</th><th class="num">应缴(20%)</th><th style="width:70px;text-align:center;">核对</th></tr></thead>
      <tbody>${rows.join('')}</tbody>
    </table>
    <div class="fn"><strong>核对状态：</strong> 请点击各税目行展开明细逐项勾选确认</div>
  `;
}

// ─── HK Tab ───
function renderHKTab(T) {
  const hk = T.hk;
  const rows = [];
  const prefix = S.activeYear + '-';

  if (hk.cap.detail.length > 0) {
    const detailId = prefix + 'hk-cap';
    const total = hk.cap.detail.length;
    S.checkState[detailId] = { total, checked: 0 };
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-hk">财产转让(股票)</span></td>
        <td class="cc">${total} 笔</td>
        <td>HKD ${hk.cap.detail.reduce((s,d)=>s+d.sellAmt,0).toFixed(0)}</td>
        <td class="num pos">${hk.cap.net.toFixed(0)}</td><td class="num">0</td>
        <td class="num pos">${hk.cap.taxDue.toFixed(0)}</td>
        <td style="text-align:center;"><span class="badge badge-hk" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${total} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>#</th><th>代码</th><th>名称</th><th>日期</th><th class="num">数量</th><th class="num">卖价</th><th class="num">FIFO成本</th><th class="num">盈亏</th><th class="num">汇率</th><th class="num">折CNY</th><th class="chk-col">✓</th></tr></thead>
              <tbody>
                ${hk.cap.detail.map((d,i) => {
                  const symName = d.symbol;
                  return `
                  <tr>
                    <td>${i+1}</td><td>${symName}${d.estimated ? ' ⚠️':''}</td><td>${symName}</td><td>${String(d.date).substring(0,10)}</td>
                    <td class="num">${d.qty}</td><td class="num">${(d.sellAmt/d.qty||0).toFixed(2)}</td>
                    <td class="num cost-cell"><input class="cost-edit" type="text" value="${(d.cost/d.qty||0).toFixed(2)}" data-symbol="${symName}" data-date="${String(d.date).substring(0,10)}" data-market="hk" data-orig-value="${(d.cost/d.qty||0).toFixed(2)}" onchange="onCostEdit(this)" /></td>
                    <td class="num gain-cell ${d.gain>=0?'pos':'neg'}">${d.gain>=0?'+':''}${d.gain.toFixed(0)}</td>
                    <td class="num">${d.rate.toFixed(5)}</td>
                    <td class="num gain-cny-cell ${d.gainCNY>=0?'pos':'neg'}">${d.gainCNY>=0?'+':''}${d.gainCNY.toFixed(0)}</td>
                    <td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td>
                  </tr>`;
                }).join('')}
              </tbody>
              <tfoot><tr class="dft"><td colspan="6">合计</td><td></td><td class="num pos">+${hk.cap.net.toFixed(0)}</td><td></td><td class="num pos">+${hk.cap.net.toFixed(0)}</td><td></td></tr></tfoot>
            </table>
            <div class="dn">⚠ 请核对 FIFO 成本。如有转仓股票，成本需手动确认。<br>📅 FIFO 成本基于全部历史数据（含往年买入），确保跨年成本计算准确。</div>
          </div>
        </td>
      </tr>`);
  }

  if (hk.fundCap.detail.length > 0) {
    const detailId = prefix + 'hk-fcap';
    const total = hk.fundCap.detail.length;
    S.checkState[detailId] = { total, checked: 0 };
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-hk">财产转让(基金)</span></td>
        <td class="cc">${total} 只</td>
        <td>${hk.fundCap.detail.map(d=>d.currency+' '+d.gain.toFixed(0)).join(' + ') || '-'}</td>
        <td class="num ${hk.fundCap.net>=0?'pos':'neg'}">${hk.fundCap.net.toFixed(0)}</td><td class="num">0</td>
        <td class="num ${hk.fundCap.taxDue>0?'pos':''}">${hk.fundCap.taxDue.toFixed(0)}</td>
        <td style="text-align:center;"><span class="badge badge-hk" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${total} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>基金</th><th>申赎记录</th><th class="num">盈亏原币</th><th class="num">加权汇率</th><th class="num">折CNY</th><th class="chk-col">✓</th></tr></thead>
              <tbody>${hk.fundCap.detail.map(d => `<tr><td>${d.symbol}${d.estimated ? ' ⚠️':''}</td><td>FIFO赎回</td><td class="num ${d.gain>=0?'pos':'neg'}">${d.gain>=0?'+':''}${d.gain.toFixed(2)} ${d.currency}</td><td class="num">${d.rate.toFixed(4)}</td><td class="num ${d.gainCNY>=0?'pos':'neg'}">${d.gainCNY>=0?'+':''}${d.gainCNY.toFixed(0)}</td><td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td></tr>`).join('')}</tbody>
            </table>
          </div>
        </td>
      </tr>`);
  }

  if (hk.div.detail.length > 0) {
    const detailId = prefix + 'hk-div';
    const total = hk.div.detail.length;
    S.checkState[detailId] = { total, checked: 0 };
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-hk">股息红利</span></td>
        <td class="cc">${total} 笔</td>
        <td>HKD ${hk.div.detail.reduce((s,d)=>s+d.grossAmount,0).toFixed(0)}</td>
        <td class="num">${hk.div.gross.toFixed(0)}</td><td class="num">${hk.div.tax.toFixed(0)}</td>
        <td class="num pos">${hk.div.netTax.toFixed(0)}</td>
        <td style="text-align:center;"><span class="badge badge-hk" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${total} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>#</th><th>股票</th><th>日期</th><th class="num">毛额(HKD)</th><th class="num">股息税</th><th class="num">汇率</th><th class="num">折CNY</th><th class="num">折CNY税</th><th class="chk-col">✓</th></tr></thead>
              <tbody>${hk.div.detail.map((d,i) => `<tr><td>${i+1}</td><td>${d.symbol}</td><td>${d.date}</td><td class="num">${d.grossAmount.toFixed(0)}</td><td class="num">${d.taxAmount.toFixed(0)}</td><td class="num">${d.rate.toFixed(5)}</td><td class="num">${d.grossCNY.toFixed(0)}</td><td class="num">${d.taxCNY.toFixed(0)}</td><td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td></tr>`).join('')}</tbody>
              <tfoot><tr class="dft"><td colspan="3">合计</td><td class="num">${hk.div.detail.reduce((s,d)=>s+d.grossAmount,0).toFixed(0)}</td><td class="num">${hk.div.detail.reduce((s,d)=>s+d.taxAmount,0).toFixed(0)}</td><td></td><td class="num">${hk.div.gross.toFixed(0)}</td><td class="num">${hk.div.tax.toFixed(0)}</td><td></td></tr></tfoot>
            </table>
          </div>
        </td>
      </tr>`);
  }

  if (hk.fundDiv.detail.length > 0) {
    const detailId = prefix + 'hk-fdiv';
    const total = hk.fundDiv.detail.length;
    S.checkState[detailId] = { total, checked: 0 };
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-hk">基金分红</span></td>
        <td class="cc">${total} 笔</td>
        <td>${Object.entries(hk.fundDiv.detail.reduce((s,d)=>{ s[d.currency]=(s[d.currency]||0)+d.amount; return s; },{})).map(([cur,amt])=>cur+' '+amt.toFixed(2)).join(', ')}</td>
        <td class="num">${hk.fundDiv.gross.toFixed(0)}</td><td class="num">0</td>
        <td class="num pos">${hk.fundDiv.taxDue.toFixed(0)}</td>
        <td style="text-align:center;"><span class="badge badge-hk" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${total} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>#</th><th>日期</th><th>币种</th><th class="num">金额</th><th class="num">汇率</th><th class="num">折CNY</th><th class="chk-col">✓</th></tr></thead>
              <tbody>${hk.fundDiv.detail.map((d,i) => `<tr><td>${i+1}</td><td>${d.date}</td><td>${d.currency}</td><td class="num">${d.amount.toFixed(2)}</td><td class="num">${d.rate.toFixed(4)}</td><td class="num">${d.cny.toFixed(2)}</td><td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td></tr>`).join('')}</tbody>
              <tfoot><tr class="dft"><td colspan="3">合计</td><td></td><td></td><td class="num">${hk.fundDiv.gross.toFixed(0)}</td><td></td></tr></tfoot>
            </table>
          </div>
        </td>
      </tr>`);
  }

  if (hk.note.detail.length > 0) {
    const detailId = prefix + 'hk-note';
    const total = hk.note.detail.length;
    S.checkState[detailId] = { total, checked: 0 };
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-hk">结构化票据</span></td>
        <td class="cc">${total} 笔</td>
        <td>${hk.note.detail.map(d=>d.currency+' '+Math.abs(d.amount).toFixed(0)).join(', ')}</td>
        <td class="num">${hk.note.gross.toFixed(0)}</td><td class="num">0</td>
        <td class="num pos">${hk.note.taxDue.toFixed(0)}</td>
        <td style="text-align:center;"><span class="badge badge-hk" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${total} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>#</th><th>日期</th><th>币种</th><th class="num">金额</th><th class="num">汇率</th><th class="num">折CNY</th><th class="chk-col">✓</th></tr></thead>
              <tbody>${hk.note.detail.map((d,i) => `<tr><td>${i+1}</td><td>${d.date}</td><td>${d.currency}</td><td class="num">${Math.abs(d.amount).toFixed(2)}</td><td class="num">${d.rate.toFixed(4)}</td><td class="num">${d.cny.toFixed(0)}</td><td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td></tr>`).join('')}</tbody>
            </table>
          </div>
        </td>
      </tr>`);
  }

  if (hk.int.detail.length > 0) {
    const detailId = prefix + 'hk-int';
    const total = hk.int.detail.length;
    S.checkState[detailId] = { total, checked: 0 };
    rows.push(`
      <tr class="erow" onclick="toggleRow(this)" data-target="${detailId}">
        <td><span class="arr">▶</span></td>
        <td class="lc"><span class="badge badge-hk">利息所得</span></td>
        <td class="cc">${total} 笔</td>
        <td>HKD ${hk.int.detail.reduce((s,d)=>s+Math.abs(d.amount),0).toFixed(2)}</td>
        <td class="num">${hk.int.gross.toFixed(2)}</td><td class="num">0</td>
        <td class="num pos">${hk.int.taxDue.toFixed(2)}</td>
        <td style="text-align:center;"><span class="badge badge-hk" style="background:#fef3c7;color:#92400e;font-size:0.72rem;" id="chk-${detailId}">0/${total} 已核</span></td>
      </tr>
      <tr class="drow" id="${detailId}">
        <td colspan="8">
          <div class="dinner">
            <table>
              <thead><tr><th>来源</th><th>日期</th><th class="num">金额</th><th class="num">汇率</th><th class="num">折CNY</th><th class="chk-col">✓</th></tr></thead>
              <tbody>${hk.int.detail.map(d => `<tr><td>账户利息</td><td>${d.date}</td><td class="num">${Math.abs(d.amount).toFixed(2)}</td><td class="num">${d.rate.toFixed(4)}</td><td class="num">${d.cny.toFixed(2)}</td><td class="chk-col"><input type="checkbox" onchange="updateCheck(this,'${detailId}')"></td></tr>`).join('')}</tbody>
            </table>
          </div>
        </td>
      </tr>`);
  }

  rows.push(`
    <tr style="font-weight:700;background:var(--primary-light);">
      <td colspan="3">🇭🇰 香港小计</td><td></td>
      <td class="num">${hk.totalTaxable.toFixed(0)}</td>
      <td class="num">${hk.credit.toFixed(0)}</td>
      <td class="num" style="color:var(--primary);">${hk.taxDue.toFixed(0)}</td><td></td>
    </tr>`);

  document.getElementById('tab-hk').innerHTML = `
    <table class="summary-table">
      <thead><tr><th style="width:36px;"></th><th>税目</th><th>笔数</th><th>原币金额</th><th class="num">折合CNY</th><th class="num">境外已缴税</th><th class="num">应缴(20%)</th><th style="width:70px;text-align:center;">核对</th></tr></thead>
      <tbody>${rows.join('')}</tbody>
    </table>
    <div class="fn"><strong>核对状态：</strong> 请点击各税目行展开明细逐项勾选确认</div>
  `;
}

// ═══ Inline Cost Editing ═══
function onCostEdit(input) {
  var market = input.dataset.market;
  var symbol = input.dataset.symbol;
  var date = input.dataset.date;
  var newVal = parseFloat(input.value);
  var oldVal = parseFloat(input.dataset.origValue);

  if (isNaN(newVal) || newVal < 0) { input.value = oldVal.toFixed(2); return; }
  newVal = parseFloat(newVal.toFixed(2));
  input.value = newVal.toFixed(2);
  if (Math.abs(newVal - oldVal) < 0.001) return;

  var T = getCurrentTax(); if (!T) return;
  var marketData = T[market];
  var detail = marketData.cap.detail;
  var entry = detail.find(function(e) { return e.symbol === symbol && String(e.date).substring(0,10) === date; });
  if (!entry) return;

  var key = market + '|' + symbol + '|' + date;
  S.costOverrides[key] = newVal;

  var newCost = newVal * entry.qty;
  entry.cost = newCost;
  entry.gain = entry.sellAmt - newCost;
  entry.gainCNY = entry.gain * entry.rate;
  entry.estimated = false;
  input.classList.add('overridden');

  var row = input.closest('tr');
  if (row) {
    var gc = row.querySelector('.gain-cell');
    var gcc = row.querySelector('.gain-cny-cell');
    if (gc) { gc.textContent = (entry.gain>=0?'+':'')+entry.gain.toFixed(0); gc.className='num gain-cell '+(entry.gain>=0?'pos':'neg'); }
    if (gcc) { gcc.textContent = (entry.gainCNY>=0?'+':'')+entry.gainCNY.toFixed(0); gcc.className='num gain-cny-cell '+(entry.gainCNY>=0?'pos':'neg'); }
  }

  recalcTaxAfterEdit(T, market, marketData.cap);
}

function recalcTaxAfterEdit(T, market, cat) {
  cat.net = cat.detail.reduce(function(s,d){return s+d.gain;},0);
  cat.taxable = Math.max(0, cat.net);
  cat.taxDue = cat.taxable * 0.2;

  var m = T[market];
  m.totalTaxable = (m.cap.taxable||0) + (m.div&&m.div.gross||0) + (m.fundCap&&m.fundCap.taxable||0) + (m.fundDiv&&m.fundDiv.gross||0) + (m.noteInc&&m.noteInc.gross||0) + (m.intInc&&m.intInc.gross||0);
  m.totalCredit = (m.div&&m.div.tax||0) + (m.fundDiv&&m.fundDiv.tax||0);
  m.taxDue = Math.max(0, m.totalTaxable * 0.2 - m.totalCredit);

  T.totalTaxable = T.us.totalTaxable + T.hk.totalTaxable;
  T.totalCredit = T.us.totalCredit + T.hk.totalCredit;
  T.totalTax = T.us.taxDue + T.hk.taxDue;

  var sr = document.getElementById('stat-row');
  if (sr) sr.innerHTML = '<div class="stat-card"><div class="sv pos">'+T.totalTax.toFixed(0)+'</div><div class="sl2">应缴个税 (CNY)</div></div><div class="stat-card"><div class="sv">'+T.totalTaxable.toFixed(0)+'</div><div class="sl2">应税所得 (CNY)</div></div><div class="stat-card"><div class="sv">'+T.totalCredit.toFixed(0)+'</div><div class="sl2">境外抵免 (CNY)</div></div><div class="stat-card"><div class="sv">'+(T.totalTaxable>0?(T.totalTax/T.totalTaxable*100).toFixed(1):'0.0')+'%</div><div class="sl2">实际税负率</div></div>';

  var tab = document.getElementById(market==='us'?'tab-us':'tab-hk');
  if (!tab) return;
  var prefix = S.activeYear + '-';
  var detailId = prefix + market + '-cap';

  var drow = document.getElementById(detailId);
  if (drow && drow.previousElementSibling) {
    var erow = drow.previousElementSibling;
    var cells = erow.querySelectorAll('td');
    if (cells.length >= 7) {
      cells[3].textContent = cat.net < 0 ? '净亏损' : '净盈利';
      cells[4].textContent = cat.net.toFixed(0);
      cells[4].className = 'num ' + (cat.net >= 0 ? 'pos' : 'neg');
      cells[6].textContent = cat.taxDue.toFixed(0);
      cells[6].className = 'num ' + (cat.taxDue > 0 ? 'pos' : '');
    }
  }

  if (drow) {
    var tfoot = drow.querySelector('tfoot');
    if (tfoot) {
      var tds = tfoot.querySelectorAll('td');
      for (var i = 0; i < tds.length; i++) {
        var cn = tds[i].className || '';
        if (cn.indexOf('pos')>=0 || cn.indexOf('neg')>=0) {
          tds[i].textContent = (cat.net>=0?'+':'')+cat.net.toFixed(0);
          tds[i].className = 'num '+(cat.net>=0?'pos':'neg');
          break;
        }
      }
    }
  }

  var st = tab.querySelector('tr[style*="font-weight:700"]');
  if (st) {
    var sc = st.querySelectorAll('td');
    if (sc.length >= 7) {
      sc[4].textContent = m.totalTaxable.toFixed(2);
      sc[5].textContent = m.totalCredit.toFixed(2);
      sc[6].textContent = m.taxDue.toFixed(2);
    }
  }
}

// ─── Rate Tab (with year highlighting) ───
function renderRateTab() {
  const currentYear = parseInt(S.activeYear);
  const years = [2022, 2023, 2024, 2025];
  let usdHtml = '', hkdHtml = '';

  for (const yr of years) {
    const u = RATES.USD[yr]; const h = RATES.HKD[yr];
    const isActiveYear = (currentYear === yr);
    const rowClass = isActiveYear ? 'rate-hl' : '';

    usdHtml += `<tr class="${rowClass}"><td colspan="4" style="font-weight:700;background:${isActiveYear?'#fef3c7':'var(--gray-100)'};">${yr}年 ${isActiveYear ? '👈 当前年度' : ''}</td></tr>`;
    for (let m = 1; m <= 12; m++) {
      usdHtml += `<tr class="${rowClass}"><td>${m}月</td><td>${u[m].toFixed(4)}</td>`;
      m++;
      usdHtml += `<td>${m}月</td><td>${u[m].toFixed(4)}</td></tr>`;
    }

    hkdHtml += `<tr class="${rowClass}"><td colspan="4" style="font-weight:700;background:${isActiveYear?'#fef3c7':'var(--gray-100)'};">${yr}年 ${isActiveYear ? '👈 当前年度' : ''}</td></tr>`;
    for (let m = 1; m <= 12; m++) {
      hkdHtml += `<tr class="${rowClass}"><td>${m}月</td><td>${h[m].toFixed(5)}</td>`;
      m++;
      hkdHtml += `<td>${m}月</td><td>${h[m].toFixed(5)}</td></tr>`;
    }
  }

  const yearInfo = currentYear
    ? `<div class="alert-info">👁️ 当前查看：<strong style="color:var(--primary);">${currentYear}年度</strong> — 汇率表自动高亮（黄色），切换顶部年度标签联动变更。</div>` : '';

  document.getElementById('tab-rate').innerHTML = `
    ${yearInfo}
    <div style="display:flex;gap:24px;flex-wrap:wrap;">
      <div style="flex:1;min-width:280px;">
        <h3 style="font-size:0.95rem;margin-bottom:8px;">USD/CNY 月末汇率</h3>
        <table class="summary-table"><thead><tr><th>月份</th><th>汇率</th><th>月份</th><th>汇率</th></tr></thead><tbody>${usdHtml}</tbody></table>
      </div>
      <div style="flex:1;min-width:280px;">
        <h3 style="font-size:0.95rem;margin-bottom:8px;">HKD/CNY 月末汇率</h3>
        <table class="summary-table"><thead><tr><th>月份</th><th>汇率</th><th>月份</th><th>汇率</th></tr></thead><tbody>${hkdHtml}</tbody></table>
      </div>
    </div>
    <div class="alert alert-tip" style="margin-top:12px;">💱 央行月末中间价（2022-2025内置）。系统自动按交易日所在月末匹配汇率折算。FIFO成本使用全部历史数据（含往年买入）。</div>
  `;
}

// ============================================================
// 14. INTERACTIONS
// ============================================================
function toggleRow(rowEl) {
  const detailId = rowEl.getAttribute('data-target');
  const detail = document.getElementById(detailId);
  if (!detail) return;
  const isOpen = rowEl.classList.contains('open');
  if (isOpen) { rowEl.classList.remove('open'); detail.classList.remove('open'); }
  else { rowEl.classList.add('open'); detail.classList.add('open'); }
}

function updateCheck(cb, detailId) {
  const detail = document.getElementById(detailId);
  if (!detail) return;
  const allCbs = detail.querySelectorAll('input[type="checkbox"]');
  let checked = 0;
  allCbs.forEach(c => { if (c.checked) checked++; });

  S.checkState[detailId] = { total: allCbs.length, checked };

  // Update badge
  const badgeId = 'chk-' + detailId;
  const badge = document.getElementById(badgeId);
  if (!badge) {
    const srow = document.querySelector(`[data-target="${detailId}"]`);
    if (!srow) return;
    const statusCell = srow.querySelector('td:last-child .badge');
    if (!statusCell) return;
    if (checked === allCbs.length) {
      statusCell.textContent = '✓ 已核';
      statusCell.style.background = '#d1fae5'; statusCell.style.color = '#065f46';
    } else {
      statusCell.textContent = checked + '/' + allCbs.length + ' 已核';
      statusCell.style.background = '#fef3c7'; statusCell.style.color = '#92400e';
    }
    return;
  }

  if (checked === allCbs.length) {
    badge.textContent = '✓ 已核';
    badge.style.background = '#d1fae5'; badge.style.color = '#065f46';
  } else {
    badge.textContent = checked + '/' + allCbs.length + ' 已核';
    badge.style.background = '#fef3c7'; badge.style.color = '#92400e';
  }
}

// Tab switching (main tabs: US/HK/Rate)
document.getElementById('main-tabs').addEventListener('click', function(e) {
  if (!e.target.classList.contains('tab')) return;
  const tabId = e.target.getAttribute('data-tab');
  document.querySelectorAll('.tab-content').forEach(el => el.style.display = 'none');
  document.querySelectorAll('#main-tabs .tab').forEach(el => el.classList.remove('active'));
  document.getElementById(tabId).style.display = 'block';
  e.target.classList.add('active');
});

function updateSteps(step) {
  for (let i = 1; i <= 4; i++) {
    const el = document.getElementById('step-' + i);
    if (i < step) el.className = 'step done';
    else if (i === step) el.className = 'step active';
    else el.className = 'step';
  }
}

// ============================================================
// 15. FILE LIST UI + REMOVE
// ============================================================
function renderFileList(zone) {
  const container = document.getElementById('flist-' + zone);
  if (!container) return;
  const files = S.files[zone] || [];
  container.innerHTML = '';
  files.forEach((f, idx) => {
    const item = document.createElement('div');
    item.className = 'fitem';
    item.innerHTML = `<span class="fname">${f.name}</span><span class="fdel" data-zone="${zone}" data-idx="${idx}" title="删除">✕</span>`;
    container.appendChild(item);
  });
  // Bind delete buttons
  container.querySelectorAll('.fdel').forEach(btn => {
    btn.addEventListener('click', function(e) {
      e.stopPropagation();
      const z = this.getAttribute('data-zone');
      const i = parseInt(this.getAttribute('data-idx'));
      removeFile(z, i);
    });
  });
  // Update zone visual state
  const zoneEl = document.getElementById('zone-' + zone);
  if (files.length > 0) zoneEl.classList.add('loaded');
  else zoneEl.classList.remove('loaded');
}

function removeFile(zone, idx) {
  S.files[zone].splice(idx, 1);
  renderFileList(zone);
  // Update calc button state
  if (S.files.bill && S.files.bill.length > 0) {
    document.getElementById('btnCalc').disabled = false;
    updateSteps(1);
  } else {
    document.getElementById('btnCalc').disabled = true;
  }
}

// ============================================================
// 16. FILE UPLOAD HANDLERS (multi-file)
// ============================================================
document.querySelectorAll('.upload-zone input[type=file]').forEach(input => {
  input.addEventListener('change', async function() {
    const zone = this.getAttribute('data-zone');
    if (!this.files || this.files.length === 0) return;

    // Append all selected files
    for (const file of this.files) {
      S.files[zone].push(file);
    }
    renderFileList(zone);
    this.value = ''; // reset so re-selecting same file triggers change

    if (S.files.bill && S.files.bill.length > 0) {
      document.getElementById('btnCalc').disabled = false;
      updateSteps(1);
    }
  });
});

// Calculate button
document.getElementById('btnCalc').addEventListener('click', async function() {
  if (!S.files.bill || S.files.bill.length === 0) return;
  this.disabled = true;
  this.textContent = '⏳ 计算中...';
  updateSteps(2);
  await processAll();
  this.disabled = false;
  this.textContent = '🔄 重新计算';
});

// Reset
function doReset() {
  S.files = { bill: [], div: [], pnl: [], crs: [] };
  S.raw = { trades: [], fundFlows: [], pnlRecords: [], divSummary: [] };
  S.parsed = { dividends: [], interest: [], fundDiv: [], fundTrades: [], noteIncome: [], fees: [], other: [] };
  S.costBasis = {}; S.allYears = []; S.activeYear = null; S.years = {}; S.checkState = {};
  document.querySelectorAll('.upload-zone').forEach(z => z.classList.remove('loaded'));
  document.querySelectorAll('.upload-zone .fname').forEach(f => f.textContent = '');
  document.querySelectorAll('.upload-zone input[type=file]').forEach(i => i.value = '');
  document.getElementById('btnCalc').disabled = true;
  document.getElementById('btnCalc').textContent = '🔍 解析并计算税额';
  document.getElementById('result-area').style.display = 'none';
  document.getElementById('empty-state').style.display = 'block';
  document.getElementById('progress-area').style.display = 'none';
  document.getElementById('log-area').innerHTML = '';
  document.getElementById('log-area').style.display = 'none';
  document.getElementById('year-bar').style.display = 'none';
  document.getElementById('year-bar').querySelectorAll('.year-tab').forEach(t => t.remove());
  updateSteps(1);
  document.getElementById('tab-us').innerHTML = '';
  document.getElementById('tab-hk').innerHTML = '';
  document.getElementById('tab-rate').innerHTML = '';
  document.getElementById('btnExport').textContent = '📥 下载年度报税底稿 (.md)';
}

document.getElementById('btnReset').addEventListener('click', doReset);
document.getElementById('btnReset2').addEventListener('click', doReset);

// ============================================================
// 16. EXPORT
// ============================================================

function buildMarkdown(T, yearLabel) {
  return `# ${yearLabel}境外所得个税申报底稿

> 生成时间: ${new Date().toISOString().substring(0,10)}
> 计税方法: 分国不分项 | 汇率: 央行月末中间价
` +
`## 应纳税额汇总

| 项目 | 金额(CNY) |
|------|----------|
| 境外应税所得合计 | ${T.totalTaxable.toFixed(2)} |
| 境外已缴税合计 | ${T.totalCredit.toFixed(2)} |
| **应缴税额** | **${T.totalTax.toFixed(2)}** |

## 分国明细

### 🇺🇸 美国
| 税目 | 应税所得 | 境外已缴税 | 应缴税 |
|------|---------|----------|-------|
| 股息红利 | ${T.us.div.gross.toFixed(2)} | ${T.us.div.tax.toFixed(2)} | ${T.us.div.netTax.toFixed(2)} |
| 利息所得 | ${T.us.int.gross.toFixed(2)} | 0 | ${T.us.int.taxDue.toFixed(2)} |
| 财产转让 | ${T.us.cap.net.toFixed(2)} | 0 | ${T.us.cap.taxDue.toFixed(2)} |
| **小计** | **${T.us.totalTaxable.toFixed(2)}** | **${T.us.credit.toFixed(2)}** | **${T.us.taxDue.toFixed(2)}** |

### 🇭🇰 中国香港
| 税目 | 应税所得 | 境外已缴税 | 应缴税 |
|------|---------|----------|-------|
| 财产转让(股票) | ${T.hk.cap.net.toFixed(2)} | 0 | ${T.hk.cap.taxDue.toFixed(2)} |
| 财产转让(基金) | ${T.hk.fundCap.net.toFixed(2)} | 0 | ${T.hk.fundCap.taxDue.toFixed(2)} |
| 股息红利 | ${T.hk.div.gross.toFixed(2)} | ${T.hk.div.tax.toFixed(2)} | ${T.hk.div.netTax.toFixed(2)} |
| 基金分红 | ${T.hk.fundDiv.gross.toFixed(2)} | 0 | ${T.hk.fundDiv.taxDue.toFixed(2)} |
| 结构化票据 | ${T.hk.note.gross.toFixed(2)} | 0 | ${T.hk.note.taxDue.toFixed(2)} |
| 利息所得 | ${T.hk.int.gross.toFixed(2)} | 0 | ${T.hk.int.taxDue.toFixed(2)} |
| **小计** | **${T.hk.totalTaxable.toFixed(2)}** | **${T.hk.credit.toFixed(2)}** | **${T.hk.taxDue.toFixed(2)}** |

---
> 本底稿由富途境外投资个税计算器自动生成。
> FIFO成本基于全部历史数据（含往年买入）。
> 税率: 20% | 申报渠道: 自然人电子税务局 Web端
`;
}

// Export current year
document.getElementById('btnExport').addEventListener('click', function() {
  const T = getCurrentTax();
  if (!T) { alert('请先完成计算'); return; }

  const yearLabel = S.activeYear + '年';
  const md = buildMarkdown(T, yearLabel);

  const blob = new Blob([md], { type: 'text/markdown;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = `${S.activeYear}境外所得个税申报底稿.md`;
  a.click();
  URL.revokeObjectURL(url);
  updateSteps(4);
});

// Init
updateSteps(1);
