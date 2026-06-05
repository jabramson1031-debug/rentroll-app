import * as XLSX from 'xlsx'

// ─── Style helpers ────────────────────────────────────────────────────────────
const pct  = '0.00%'
const usd  = '$#,##0'
const usd2 = '$#,##0.00'
const num2 = '#,##0.00'
const num0 = '#,##0'

function cell(v, z) {
  if (v === null || v === undefined) return { v: '', t: 's' }
  if (typeof v === 'number') return z ? { v, t: 'n', z } : { v, t: 'n' }
  return { v: String(v), t: 's' }
}

function header(v) { return { v, t: 's' } }

function buildSheet(rows) {
  return XLSX.utils.aoa_to_sheet(rows)
}

// ─── INPUTS ───────────────────────────────────────────────────────────────────
function buildInputs(data) {
  const { inputs: I } = data
  const rows = [
    [header('INPUT PAGE — MULTIFAMILY / MIXED USE')],
    [],
    [header('PROPERTY INFORMATION'), '', header('CURRENT DEBT'), '', header('EXPENSE CONTROLS')],
    ['Property Name',       I.propertyName,    'Loan Amount',      cell(I.debt.loanAmount, usd),   'Insurance ($/unit)',       cell(I.expenses.insurance, usd)],
    ['Address',             I.address,          'Interest Rate',    cell(I.debt.rate, pct),          'Water & Sewer ($/unit)',   cell(I.expenses.waterSewer, usd)],
    ['Borough / City',      I.borough,          'Term (yrs)',       cell(I.debt.term, num0),         'Super Salary ($/unit)',    cell(I.expenses.superSalary, usd)],
    ['Zip Code',            I.zip,              'Amortization',     cell(I.debt.amort, num0),        'Repairs & Maint ($/unit)', cell(I.expenses.repairsMaint, usd)],
    ['Total Units',         cell(I.totalUnits, num0), '',           '',                              'Common Electric ($/sf)',   cell(I.expenses.commonElectric, usd2)],
    ['Residential Units',   cell(I.residentialUnits, num0), '', '', '',                              'Fuel – Oil ($/unit)',      cell(I.expenses.fuelOil, usd)],
    ['Commercial Units',    cell(I.commercialUnits, num0), '', '',  '',                              'General Admin ($/unit)',   cell(I.expenses.generalAdmin, usd)],
    ['Gross SF',            cell(I.grossSF, num0),          '', '', '',                              'Management Fee',           cell(I.expenses.managementPct, pct)],
    ['Loss Factor',         cell(I.lossFactor, pct),        '', '', '',                              'Vacancy Rate',             cell(I.expenses.vacancyPct, pct)],
    ['Net SF',              cell(I.netSF, num0),            '', '', ''],
    ['Property Taxes',      cell(I.taxes, usd)],
    ['Tax Class',           I.taxClass],
    [],
    [header('MARKET RATE RENT CONTROLS (Monthly)'), '', header('PRICING'), '', header('GROWTH RATES')],
    ['Studio',       cell(I.marketRents.studio,   usd), 'Listing Price',    cell(I.listingPrice,    usd), 'RS Growth Rate',  cell(I.rsGrowthRate, pct)],
    ['1 Bedroom',    cell(I.marketRents.oneBed,   usd), 'Investment Value', cell(I.investmentValue, usd), 'FM Growth Rate',  cell(I.fmGrowthRate, pct)],
    ['2 Bedroom',    cell(I.marketRents.twoBed,   usd)],
    ['3 Bedroom',    cell(I.marketRents.threeBed, usd)],
    ['4 Bedroom',    cell(I.marketRents.fourBed,  usd)],
    ['5 Bedroom',    cell(I.marketRents.fiveBed,  usd)],
  ]
  const ws = buildSheet(rows)
  ws['!cols'] = [{ wch: 28 }, { wch: 22 }, { wch: 22 }, { wch: 18 }, { wch: 28 }, { wch: 18 }]
  return ws
}

// ─── RENT ROLL ────────────────────────────────────────────────────────────────
function buildRentRoll(data) {
  const { inputs: I, rentRoll } = data
  const rows = [
    [header('RENT ROLL'), I.propertyName],
    [],
    ['Unit #', 'Status', 'Type', 'Bedrooms', 'SF', 'Actual Rent/Mo', 'Pro Forma Rent/Mo',
     'Annual Actual', 'Annual Pro Forma', 'Lease Expiration', 'Notes'],
    ...rentRoll.map(u => [
      u.unit, u.status, u.type, u.bedrooms,
      cell(u.sf,           num0),
      cell(u.actualRent,   usd),
      cell(u.proFormaRent, usd),
      cell((u.actualRent   || 0) * 12, usd),
      cell((u.proFormaRent || 0) * 12, usd),
      u.leaseExp || '', u.notes || ''
    ]),
    [],
    ['TOTALS', '', '', '',
      cell(rentRoll.reduce((s,u)=>s+(u.sf||0),0),           num0),
      cell(rentRoll.reduce((s,u)=>s+(u.actualRent||0),0),   usd),
      cell(rentRoll.reduce((s,u)=>s+(u.proFormaRent||0),0), usd),
      cell(rentRoll.reduce((s,u)=>s+(u.actualRent||0)*12,0),  usd),
      cell(rentRoll.reduce((s,u)=>s+(u.proFormaRent||0)*12,0),usd),
    ],
  ]
  const ws = buildSheet(rows)
  ws['!cols'] = [
    {wch:10},{wch:8},{wch:14},{wch:14},{wch:10},
    {wch:18},{wch:20},{wch:16},{wch:18},{wch:16},{wch:24}
  ]
  return ws
}

// ─── CASH FLOW ────────────────────────────────────────────────────────────────
function buildCashFlow(data) {
  const { inputs: I, cashFlow: CF } = data
  const n = I.totalUnits

  const row = (label, act, pf, z=usd, indent=false) => [
    indent ? `  ${label}` : label,
    cell(act, z),
    act && CF.egi  ? cell(act/CF.egi,  pct) : '',
    act && n       ? cell(act/n,        usd) : '',
    cell(pf,  z),
    pf  && CF.pfEgi ? cell(pf/CF.pfEgi, pct) : '',
    pf  && n        ? cell(pf/n,         usd) : '',
  ]

  const rows = [
    [header('INCOME & EXPENSE ANALYSIS'), I.propertyName],
    [],
    ['', header('ACTUAL'), '% EGI', '$/UNIT', header('PRO FORMA'), '% EGI', '$/UNIT'],
    [header('GROSS POTENTIAL INCOME')],
    row('Gross Potential Residential Rent', CF.gprResi,   CF.pfGprResi,   usd, true),
    row('Gross Potential Commercial Rent',  CF.gprComm,   CF.pfGprComm,   usd, true),
    row('Gross Income',                     CF.grossIncome, CF.pfGrossIncome),
    row('Vacancy / Collection Loss',        -CF.vacancy,  -CF.pfVacancy),
    row('Effective Gross Income (EGI)',      CF.egi,        CF.pfEgi),
    [],
    [header('EXPENSES')],
    row('Property Taxes',      CF.taxes,        CF.pfTaxes,       usd, true),
    row('Insurance',           CF.insurance,    CF.pfInsurance,   usd, true),
    row('Water & Sewer',       CF.waterSewer,   CF.pfWaterSewer,  usd, true),
    row('Repairs & Maint.',    CF.repairsMaint, CF.pfRepairsMaint,usd, true),
    row('Common Electric',     CF.commonElec,   CF.pfCommonElec,  usd, true),
    row('Super Salary',        CF.superSalary,  CF.pfSuperSalary, usd, true),
    row('Management Fee',      CF.mgmtFee,      CF.pfMgmtFee,     usd, true),
    row('Fuel – Oil',          CF.fuelOil,      CF.pfFuelOil,     usd, true),
    row('Total Expenses',      CF.totalExp,     CF.pfTotalExp),
    [],
    row('NET OPERATING INCOME (NOI)', CF.noi, CF.pfNoi),
    [],
    ['Annual Debt Service',    cell(-CF.annualDebtService, usd), '', '', cell(-CF.annualDebtService, usd)],
    row('Cash Flow After Debt', CF.cashFlowAfterDebt, CF.pfCashFlowAfterDebt),
  ]

  const ws = buildSheet(rows)
  ws['!cols'] = [{wch:32},{wch:16},{wch:8},{wch:12},{wch:16},{wch:8},{wch:12}]
  return ws
}

// ─── PRICING ──────────────────────────────────────────────────────────────────
function buildPricing(data) {
  const { inputs: I, cashFlow: CF, pricing: PR } = data
  const dscr = CF.annualDebtService > 0 ? CF.pfNoi / CF.annualDebtService : 0

  const rows = [
    [header('PRICING ANALYTICS'), I.propertyName],
    [],
    ['LISTING PRICE', cell(I.listingPrice, usd), '', 'INVESTMENT VALUE', cell(I.investmentValue, usd)],
    [],
    ['METRIC', 'LISTING — CURRENT', 'LISTING — PRO FORMA', 'INV. VALUE — CURRENT', 'INV. VALUE — PRO FORMA'],
    ['Net Operating Income', cell(CF.noi, usd), cell(CF.pfNoi, usd), cell(CF.noi, usd), cell(CF.pfNoi, usd)],
    ['Cash Flow After Debt',  cell(CF.cashFlowAfterDebt, usd), cell(CF.pfCashFlowAfterDebt, usd), '', ''],
    ['Cap Rate',    cell(PR.listCapCurrent, pct), cell(PR.listCapPF, pct), cell(PR.invCapCurrent, pct), cell(PR.invCapPF, pct)],
    ['GRM',         cell(PR.listGRMCurrent, num2), cell(PR.listGRMPF, num2), cell(PR.invGRMCurrent, num2), cell(PR.invGRMPF, num2)],
    ['Cash-on-Cash',cell(PR.listCoCCurrent, pct), cell(PR.listCoCPF, pct), '', ''],
    ['$/SF',        cell(PR.listPPSF, usd2), '', cell(PR.invPPSF, usd2), ''],
    ['$/Unit',      cell(PR.listPPUnit, usd), '', cell(PR.invPPUnit, usd), ''],
    [],
    [header('PROPOSED FINANCING')],
    ['Loan Amount',       cell(I.debt.loanAmount, usd)],
    ['Interest Rate',     cell(I.debt.rate, pct)],
    ['Term',              cell(I.debt.term, num0)],
    ['Amortization',      cell(I.debt.amort, num0)],
    ['Monthly Payment',   cell(CF.monthly, usd)],
    ['Annual Debt Service', cell(CF.annualDebtService, usd)],
    ['DSCR (Pro Forma)',  cell(dscr, num2)],
    ['Loan-to-Value',     cell(I.debt.loanAmount / I.listingPrice, pct)],
    ['Equity Required',   cell(I.listingPrice - I.debt.loanAmount, usd)],
    [],
    [header('SALES COMPARABLE ANALYSIS')],
    ['Market Cap Rate', cell(PR.mktCap, pct)],
    ['Market GRM',      cell(PR.mktGRM, num2)],
    ['Market $/SF',     cell(PR.mktPPSF, usd)],
    ['Reconciled Comp Value', cell(PR.compValue, usd)],
  ]

  const ws = buildSheet(rows)
  ws['!cols'] = [{wch:28},{wch:20},{wch:20},{wch:22},{wch:22}]
  return ws
}

// ─── 10-YEAR DCF ──────────────────────────────────────────────────────────────
function buildDCF(data) {
  const { inputs: I, dcf: D } = data
  const yrs = [0,1,2,3,4,5,6,7,8,9,10]
  const ylabels = ['Year 0','Year 1','Year 2','Year 3','Year 4','Year 5','Year 6','Year 7','Year 8','Year 9','Year 10']

  const ncell = (arr, i, z=usd) => arr[i] !== null && arr[i] !== undefined ? cell(arr[i], z) : ''

  const rows = [
    [header('10-YEAR CASH FLOW ANALYSIS'), I.propertyName],
    [],
    ['GROWTH ASSUMPTIONS', '', ...yrs.slice(1).map(() => '')],
    ['RS Growth Rate',    '', ...yrs.slice(1).map(() => cell(I.rsGrowthRate, pct))],
    ['FM Growth Rate',    '', ...yrs.slice(1).map(() => cell(I.fmGrowthRate, pct))],
    ['Expense Growth',    '', ...yrs.slice(1).map(() => cell(0.02, pct))],
    ['Vacancy Rate',      '', ...yrs.slice(1).map(() => cell(I.expenses.vacancyPct, pct))],
    [],
    ['', ...ylabels],
    [],
    [header('INCOME')],
    ['GPR Residential',    ...D.gprResi.map(v=>cell(v,usd))],
    ['GPR Commercial',     ...D.gprComm.map(v=>cell(v,usd))],
    ['Vacancy Loss',       ...D.vacancy.map(v=>cell(-v,usd))],
    ['Effective Gross Income', ...D.egi.map(v=>cell(v,usd))],
    [],
    [header('EXPENSES')],
    ['Property Taxes',     ...D.taxes.map(v=>cell(v,usd))],
    ['Insurance',          ...D.insurance.map(v=>cell(v,usd))],
    ['Water & Sewer',      ...D.waterSewer.map(v=>cell(v,usd))],
    ['Repairs & Maint.',   ...D.repairsMaint.map(v=>cell(v,usd))],
    ['Common Electric',    ...D.commonElec.map(v=>cell(v,usd))],
    ['Super Salary',       ...D.superSalary.map(v=>cell(v,usd))],
    ['Management Fee',     ...D.mgmtFee.map(v=>cell(v,usd))],
    ['Fuel – Oil',         ...D.fuelOil.map(v=>cell(v,usd))],
    ['Total Expenses',     ...D.totalExp.map(v=>cell(v,usd))],
    [],
    [header('NET OPERATING INCOME'), ...D.noi.map(v=>cell(v,usd))],
    [],
    [header('ACQUISITION & DISPOSITION')],
    ['Purchase Price',     cell(-I.listingPrice, usd), ...Array(10).fill('')],
    ['Acquisition Costs (2%)', cell(-I.listingPrice*0.02,usd), ...Array(10).fill('')],
    ['Reversion Proceeds', '', ...Array(9).fill(''), cell(D.reversionProceeds,usd)],
    ['Net CF (Unleveraged)', ...D.netCFUnlev.map(v=>cell(v,usd))],
    [],
    [header('DEBT SERVICE')],
    ['Loan Amount',        cell(I.debt.loanAmount,usd), ...Array(10).fill('')],
    ['Interest',           '', ...D.interest.map(v=>cell(-v,usd))],
    ['Principal',          '', ...D.principal.map(v=>cell(-v,usd))],
    ['Cash Flow After Debt', ...D.netCFLev.map(v=>cell(v,usd))],
    [],
    [header('ANNUAL METRICS')],
    ['Cap Rate',           '', ...D.capRate.slice(1).map(v=>v!==null?cell(v,pct):'')],
    ['DSCR',               '', ...D.dscr.slice(1).map(v=>v!==null?cell(v,num2):'')],
    ['Unlev. Cash-on-Cash','', ...D.unlevCoC.slice(1).map(v=>v!==null?cell(v,pct):'')],
    ['Lev. Cash-on-Cash',  '', ...D.levCoC.slice(1).map(v=>v!==null?cell(v,pct):'')],
    [],
    [header('SUMMARY RETURNS')],
    ['Unlevered IRR',      cell(D.unlevIRR, pct)],
    ['Levered IRR',        cell(D.levIRR, pct)],
    ['Unlevered NPV (4%)', cell(D.unlevNPV, usd)],
    ['Levered NPV (4%)',   cell(D.levNPV, usd)],
    ['Equity Multiple',    D.equityMultiple ? cell(D.equityMultiple, num2) : 'N/A'],
    ['Exit Cap Rate',      cell(0.065, pct)],
    ['Exit Value (Year 10)', cell(D.reversionProceeds, usd)],
  ]

  const ws = buildSheet(rows)
  ws['!cols'] = [{wch:30}, ...yrs.map(()=>({wch:15}))]
  return ws
}

// ─── VALUE SUMMARY ────────────────────────────────────────────────────────────
function buildValueSummary(data) {
  const { inputs: I, cashFlow: CF, pricing: PR, dcf: D } = data
  const dscr = CF.annualDebtService > 0 ? CF.pfNoi / CF.annualDebtService : 0

  const rows = [
    [header('VALUE SUMMARY & RECOMMENDED LISTING PRICE'), I.propertyName],
    [],
    ['INVESTMENT VALUE (Income Approach)', cell(I.investmentValue, usd)],
    ['  Based on Pro Forma Cap Rate', cell(I.investmentValue, usd)],
    [],
    ['SALES COMPARABLE VALUE', cell(PR.compValue, usd)],
    [],
    ['RECOMMENDED LISTING PRICE', cell(I.listingPrice, usd)],
    [],
    ['KEY METRICS AT LISTING PRICE', 'CURRENT', 'PRO FORMA'],
    ['Net Operating Income',     cell(CF.noi, usd),   cell(CF.pfNoi, usd)],
    ['Effective Gross Income',   cell(CF.egi, usd),   cell(CF.pfEgi, usd)],
    ['Cap Rate',                 cell(PR.listCapCurrent, pct), cell(PR.listCapPF, pct)],
    ['GRM',                      cell(PR.listGRMCurrent, num2), cell(PR.listGRMPF, num2)],
    ['Cash-on-Cash Return',      cell(PR.listCoCCurrent, pct), cell(PR.listCoCPF, pct)],
    ['Price Per SF',             cell(PR.listPPSF, usd2), ''],
    ['Price Per Unit',           cell(PR.listPPUnit, usd), ''],
    ['DSCR (Pro Forma)',         '',  cell(dscr, num2)],
    ['Annual Debt Service',      cell(CF.annualDebtService, usd), ''],
    ['Equity Required',          cell(I.listingPrice - I.debt.loanAmount, usd), ''],
    [],
    ['10-YEAR RETURN SUMMARY',   'VALUE'],
    ['Unlevered IRR',            cell(D.unlevIRR, pct)],
    ['Levered IRR',              cell(D.levIRR, pct)],
    ['Unlevered NPV (4%)',       cell(D.unlevNPV, usd)],
    ['Levered NPV (4%)',         cell(D.levNPV, usd)],
    ['Equity Multiple',          D.equityMultiple ? cell(D.equityMultiple, num2) : 'N/A'],
    ['Exit Value (Year 10)',     cell(D.reversionProceeds, usd)],
  ]

  const ws = buildSheet(rows)
  ws['!cols'] = [{wch:36},{wch:20},{wch:20}]
  return ws
}

// ─── Main export ──────────────────────────────────────────────────────────────
export function generateExcel(data) {
  const wb = XLSX.utils.book_new()

  XLSX.utils.book_append_sheet(wb, buildInputs(data),       'INPUTS')
  XLSX.utils.book_append_sheet(wb, buildRentRoll(data),     'RENT ROLL')
  XLSX.utils.book_append_sheet(wb, buildCashFlow(data),     'CASH FLOW')
  XLSX.utils.book_append_sheet(wb, buildPricing(data),      'PRICING')
  XLSX.utils.book_append_sheet(wb, buildDCF(data),          '10 YEAR DCF')
  XLSX.utils.book_append_sheet(wb, buildValueSummary(data), 'VALUE SUMMARY')

  return XLSX.write(wb, { bookType: 'xlsx', type: 'array' })
}
