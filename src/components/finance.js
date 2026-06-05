// ─── IRR (Newton's method) ────────────────────────────────────────────────────
export function calcIRR(cashflows) {
  let rate = 0.1
  for (let i = 0; i < 200; i++) {
    let npv = 0, dnpv = 0
    cashflows.forEach((cf, t) => {
      const disc = Math.pow(1 + rate, t)
      npv  += cf / disc
      dnpv -= (t * cf) / (disc * (1 + rate))
    })
    if (Math.abs(dnpv) < 1e-12) break
    const newRate = rate - npv / dnpv
    if (Math.abs(newRate - rate) < 1e-9) return newRate
    rate = isFinite(newRate) ? newRate : 0.1
  }
  return rate
}

export function calcNPV(cashflows, rate) {
  return cashflows.reduce((acc, cf, t) => acc + cf / Math.pow(1 + rate, t), 0)
}

// ─── Monthly mortgage payment ────────────────────────────────────────────────
export function monthlyPayment(principal, annualRate, years) {
  if (principal <= 0 || annualRate <= 0) return 0
  const r = annualRate / 12
  const n = years * 12
  return principal * r / (1 - Math.pow(1 + r, -n))
}

// ─── Market rent lookup ──────────────────────────────────────────────────────
function getMarketRent(bedrooms, marketRents) {
  const map = {
    'Studio':     marketRents.studio,
    '1 Bedroom':  marketRents.oneBed,
    '2 Bedroom':  marketRents.twoBed,
    '3 Bedroom':  marketRents.threeBed,
    '4 Bedroom':  marketRents.fourBed,
    '5 Bedroom':  marketRents.fiveBed,
  }
  return map[bedrooms] || 0
}

// ─── Main model builder ───────────────────────────────────────────────────────
export function buildFinancialModel(rentRoll, assumptions) {
  const {
    propertyName, address, borough, zip,
    grossSF, taxes, taxClass,
    listingPrice, investmentValue,
    debt, expenses, marketRents,
    rsGrowthRate, fmGrowthRate,
  } = assumptions

  const residentialUnits = rentRoll.filter(u => u.type !== 'commercial').length
  const commercialUnits  = rentRoll.filter(u => u.type === 'commercial').length
  const totalUnits       = rentRoll.length
  const lossFactor       = 0.15
  const netSF            = grossSF * (1 - lossFactor)

  // ── Process rent roll with pro forma rents ──────────────────────────────
  const processedRR = rentRoll.map(u => {
    const isRS   = u.status === 'RS' || u.status === 'RC'
    const isComm = u.type === 'commercial'
    let proFormaRent = u.actualRent

    if (isRS) {
      proFormaRent = u.actualRent * (1 + rsGrowthRate)
    } else if (!isComm) {
      const mkt = getMarketRent(u.bedrooms, marketRents)
      if (mkt > u.actualRent) {
        proFormaRent = mkt
      } else {
        proFormaRent = u.actualRent * (1 + fmGrowthRate)
      }
    }
    return { ...u, proFormaRent: Math.round(proFormaRent) }
  })

  // ── Annual income ────────────────────────────────────────────────────────
  const resi = processedRR.filter(u => u.type !== 'commercial')
  const comm = processedRR.filter(u => u.type === 'commercial')

  const gprResi   = resi.reduce((s, u) => s + (u.actualRent  || 0) * 12, 0)
  const gprComm   = comm.reduce((s, u) => s + (u.actualRent  || 0) * 12, 0)
  const pfGprResi = resi.reduce((s, u) => s + (u.proFormaRent || 0) * 12, 0)
  const pfGprComm = comm.reduce((s, u) => s + (u.proFormaRent || 0) * 12, 0)

  const grossIncome   = gprResi + gprComm
  const pfGrossIncome = pfGprResi + pfGprComm
  const vacancy       = grossIncome   * expenses.vacancyPct
  const pfVacancy     = pfGrossIncome * expenses.vacancyPct
  const egi           = grossIncome   - vacancy
  const pfEgi         = pfGrossIncome - pfVacancy

  // ── Expenses ─────────────────────────────────────────────────────────────
  const eu = v => v * totalUnits          // per-unit expense
  const es = v => v * (grossSF || 0)      // per-SF expense
  const g3  = v => v * 1.03              // 3% growth for pro forma

  const insurance   = eu(expenses.insurance)
  const waterSewer  = eu(expenses.waterSewer)
  const superSalary = eu(expenses.superSalary)
  const repairsMaint= eu(expenses.repairsMaint)
  const commonElec  = es(expenses.commonElectric)
  const fuelOil     = eu(expenses.fuelOil)
  const generalAdmin= eu(expenses.generalAdmin)
  const mgmtFee     = egi  * expenses.managementPct
  const pfMgmtFee   = pfEgi * expenses.managementPct

  const totalExp = taxes + insurance + waterSewer + superSalary +
                   repairsMaint + commonElec + fuelOil + generalAdmin + mgmtFee

  const pfTaxes = taxes * 1.08   // taxes grow ~8% per year in NYC
  const pfTotalExp = pfTaxes + g3(insurance) + g3(waterSewer) + g3(superSalary) +
                     g3(repairsMaint) + g3(commonElec) + g3(fuelOil) + g3(generalAdmin) + pfMgmtFee

  const noi   = egi   - totalExp
  const pfNoi = pfEgi - pfTotalExp

  // ── Debt service ──────────────────────────────────────────────────────────
  const { loanAmount, rate, amort } = debt
  const monthly          = monthlyPayment(loanAmount, rate, amort)
  const annualDebtService= monthly * 12
  const cashFlowAfterDebt  = noi   - annualDebtService
  const pfCashFlowAfterDebt= pfNoi - annualDebtService
  const equity             = listingPrice - loanAmount

  // ── Pricing metrics ───────────────────────────────────────────────────────
  const safe = (n, d) => d > 0 ? n / d : 0

  const listCapCurrent = safe(noi,   listingPrice)
  const listCapPF      = safe(pfNoi, listingPrice)
  const invCapCurrent  = safe(noi,   investmentValue)
  const invCapPF       = safe(pfNoi, investmentValue)
  const listGRMCurrent = safe(listingPrice, grossIncome)
  const listGRMPF      = safe(listingPrice, pfGrossIncome)
  const invGRMCurrent  = safe(investmentValue, grossIncome)
  const invGRMPF       = safe(investmentValue, pfGrossIncome)
  const listCoCCurrent = equity > 0 ? safe(cashFlowAfterDebt,   equity) : 0
  const listCoCPF      = equity > 0 ? safe(pfCashFlowAfterDebt, equity) : 0
  const listPPSF       = safe(listingPrice,    grossSF)
  const listPPUnit     = safe(listingPrice,    totalUnits)
  const invPPSF        = safe(investmentValue, grossSF)
  const invPPUnit      = safe(investmentValue, totalUnits)
  const mktCap         = 0.06
  const mktGRM         = 13
  const mktPPSF        = grossSF > 0 ? (investmentValue / grossSF) * 0.9 : 800
  const compValue      = (safe(noi, mktCap) + investmentValue * 1.02) / 2

  // ── 10-year DCF ───────────────────────────────────────────────────────────
  const expGrow = (base, yr) => base * Math.pow(1.02, yr)

  // Year 0 = current, years 1-10 = projections
  const dcfGPRResi = [gprResi,   ...Array.from({length:10},(_,i)=> pfGprResi * Math.pow(1+fmGrowthRate, i))]
  const dcfGPRComm = [gprComm,   ...Array.from({length:10},(_,i)=> pfGprComm * Math.pow(1.04, i))]
  const dcfGross   = dcfGPRResi.map((v,i)=> v + dcfGPRComm[i])
  const dcfVacancy = dcfGross.map(v => v * expenses.vacancyPct)
  const dcfEGI     = dcfGross.map((v,i)=> v - dcfVacancy[i])

  const dcfTaxes = [taxes, ...Array.from({length:10},(_,i)=> expGrow(taxes, i+1))]
  const dcfIns   = [insurance, ...Array.from({length:10},(_,i)=> expGrow(insurance, i+1))]
  const dcfWS    = [waterSewer, ...Array.from({length:10},(_,i)=> expGrow(waterSewer, i+1))]
  const dcfRM    = [repairsMaint,...Array.from({length:10},(_,i)=> expGrow(repairsMaint, i+1))]
  const dcfCE    = [commonElec, ...Array.from({length:10},(_,i)=> expGrow(commonElec, i+1))]
  const dcfSS    = [superSalary,...Array.from({length:10},(_,i)=> expGrow(superSalary, i+1))]
  const dcfFuel  = [fuelOil,    ...Array.from({length:10},(_,i)=> expGrow(fuelOil, i+1))]
  const dcfMgmt  = dcfEGI.map(v => v * expenses.managementPct)
  const dcfTotalExp = dcfTaxes.map((v,i)=>
    v + dcfIns[i] + dcfWS[i] + dcfRM[i] + dcfCE[i] + dcfSS[i] + dcfFuel[i] + dcfMgmt[i])
  const dcfNOI = dcfEGI.map((v,i)=> v - dcfTotalExp[i])

  // Reversion at year 10 at 6.5% exit cap
  const exitCap = 0.065
  const reversionProceeds = dcfNOI[10] / exitCap

  // Remaining loan balance at year 10
  const remainingBal = loanAmount > 0
    ? loanAmount * Math.pow(1+rate,10) - annualDebtService * (Math.pow(1+rate,10)-1) / rate
    : 0

  const acquisitionCost = listingPrice * 1.02

  const netCFUnlev = [
    -acquisitionCost,
    ...dcfNOI.slice(1,10),
    dcfNOI[10] + reversionProceeds
  ]

  // Year-by-year interest on declining balance
  const dcfInterest = Array.from({length:10}, (_,i)=> {
    if (loanAmount <= 0) return 0
    let bal = loanAmount
    for (let y=0; y<i; y++) bal = bal*(1+rate) - annualDebtService
    return bal * rate
  })
  const dcfPrincipal = dcfInterest.map(int => annualDebtService - int)

  const netCFLev = [
    -(acquisitionCost - loanAmount),
    ...dcfNOI.slice(1,10).map(v => v - annualDebtService),
    dcfNOI[10] - annualDebtService + reversionProceeds - Math.max(0, remainingBal)
  ]

  const unlevIRR = calcIRR(netCFUnlev)
  const levIRR   = calcIRR(netCFLev)
  const unlevNPV = calcNPV(netCFUnlev, 0.04)
  const levNPV   = calcNPV(netCFLev,   0.04)

  const totalReturn = netCFLev.slice(1).reduce((s,v)=>s+v, 0)
  const equityMultiple = equity > 0 ? (equity + totalReturn) / equity : null

  const dcfCapRate  = [null, ...dcfNOI.slice(1).map(v => safe(v, listingPrice))]
  const dcfDSCR     = [null, ...dcfNOI.slice(1).map(v => safe(v, annualDebtService))]
  const dcfUnlevCoC = [null, ...dcfNOI.slice(1).map(v => safe(v, acquisitionCost))]
  const dcfLevCoC   = [null, ...netCFLev.slice(1).map(v => equity > 0 ? safe(v, equity) : 0)]

  return {
    inputs: {
      propertyName, address, borough, zip,
      totalUnits, residentialUnits, commercialUnits,
      grossSF, netSF, lossFactor,
      taxes, taxClass, listingPrice, investmentValue,
      debt, expenses, marketRents, rsGrowthRate, fmGrowthRate,
    },
    rentRoll: processedRR,
    cashFlow: {
      gprResi, gprComm, pfGprResi, pfGprComm,
      otherIncome: 0,
      grossIncome, pfGrossIncome,
      vacancy, pfVacancy, egi, pfEgi,
      insurance, pfInsurance: g3(insurance),
      waterSewer, pfWaterSewer: g3(waterSewer),
      repairsMaint, pfRepairsMaint: g3(repairsMaint),
      commonElec, pfCommonElec: g3(commonElec),
      superSalary, pfSuperSalary: g3(superSalary),
      mgmtFee, pfMgmtFee,
      fuelOil, pfFuelOil: g3(fuelOil),
      taxes, pfTaxes,
      totalExp, pfTotalExp,
      noi, pfNoi,
      monthly, annualDebtService,
      cashFlowAfterDebt, pfCashFlowAfterDebt,
    },
    pricing: {
      listCapCurrent, listCapPF, invCapCurrent, invCapPF,
      listGRMCurrent, listGRMPF, invGRMCurrent, invGRMPF,
      listCoCCurrent, listCoCPF,
      listPPSF, listPPUnit, invPPSF, invPPUnit,
      mktCap, mktGRM, mktPPSF, compValue,
    },
    dcf: {
      gprResi: dcfGPRResi, gprComm: dcfGPRComm,
      egi: dcfEGI, vacancy: dcfVacancy,
      taxes: dcfTaxes, insurance: dcfIns, waterSewer: dcfWS,
      repairsMaint: dcfRM, commonElec: dcfCE, superSalary: dcfSS,
      mgmtFee: dcfMgmt, fuelOil: dcfFuel, totalExp: dcfTotalExp, noi: dcfNOI,
      reversionProceeds,
      netCFUnlev, netCFLev,
      interest: dcfInterest, principal: dcfPrincipal,
      unlevIRR, levIRR, unlevNPV, levNPV, equityMultiple,
      capRate: dcfCapRate, dscr: dcfDSCR,
      unlevCoC: dcfUnlevCoC, levCoC: dcfLevCoC,
    },
  }
}
