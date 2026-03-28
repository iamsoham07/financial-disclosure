/**
 * Excel filling logic for both templates
 *
 * Uses xlsx-populate which preserves ALL template formatting (colours, borders,
 * merges, row heights, conditional formatting) and keeps existing formulas intact.
 *
 * Assisted     → Sheet: "Financial Disclosure"
 * Negotiation  → Sheet: "2. Assets, Debts & Net Effect " (trailing space)
 *
 * NEGOTIATION TEMPLATE — confirmed cell map from Cattell filled example:
 *
 *  Names
 *    B2 = Petitioner first name    B3 = Respondent first name
 *    D3 = case number + " - "
 *
 *  Property address labels (H6 = FMH address, L6 = Prop2 name — written as text)
 *
 *  FMH sub-table (col J)          Property 2 sub-table (col N)
 *    J7  = FMH value                N7  = Prop2 value
 *    J8  = FMH mortgage             N8  = Prop2 mortgage
 *    J10 = FMH ERP (0)              N10 = Prop2 CGT (0)
 *    J9  = formula =SUM(0.02*J7)    N9  = formula — NOT touched
 *
 *  Petitioner sole assets  label→H col, value→J col, rows 16–29
 *    H16/J16 … H29/J29  (individual accounts/assets, each on its own row)
 *    J30 = formula =SUM(J16:J29)  — NOT touched
 *
 *  Respondent sole assets  label→L col, value→N col, rows 16–29
 *    L16/N16 … L29/N29
 *    N30 = formula =SUM(N16:N29)  — NOT touched
 *
 *  Joint assets  label→H col, J col = pet share, N col = res share, rows 34–45
 *    J34/N34 … J45/N45
 *    J46/N46 = formula =SUM(J34:J45)  — NOT touched
 *
 *  Vehicles written into H/K cols (pet) and L/O cols (res)
 *    H23/K23 = pet vehicle label/value
 *    L23/O23 = res vehicle label/value
 *
 *  Additional items (jewellery, watches etc.) written into H/K and L/O rows 21-22
 *    H21/K21 = pet item1,  H22/K22 = pet item2
 *    L21/O21 = res item1,  L22/O22 = res item2
 *
 *  Liabilities  pet→J51:J58 (label in H), res→N51:N58 (label in L)
 *    J59/N59 = formula =SUM(J51:J58) / =SUM(N51:N58)  — NOT touched
 *
 *  Pensions  pet→J62:J68 (label in H), res→N62:N68 (label in L)
 *    J69/N69 = formula =SUM(J62:J68) / =SUM(N62:N68)  — NOT touched
 *
 *  Income now  B22/D22=salary, B23/D23=benefits, B24/D24=rental, B25/D25=interest
 *    B27/D27 = formula =SUM(B22:B26)  — NOT touched
 *
 *  Income after (future)  B58/D58=salary, B59/D59=benefits, B60/D60=pension,
 *                          B61/D61=interest, B62/D62=bonus
 *    B63/D63 = formula =SUM(B58:B62)  — NOT touched
 *
 *  Dates (rows 77–80)
 *    B77 = Pet DOB   D77 = Pet occupation   E77 = "Cohabitation"   F77 = cohab date
 *    B78 = Res DOB   D78 = Res occupation   E78 = "Marriage "      F78 = marriage date
 *    A79 = Child1 name  B79 = Child1 DOB    E79 = "Separation "    F79 = sep date
 *    A80 = Child2 name  B80 = Child2 DOB    E80 = "CDO..."         F80 = cond date
 *
 *  Commentary  B84
 */

const XlsxPopulate = require('xlsx-populate');

// ─── SHARED HELPERS ──────────────────────────────────────────────────────────

function buildLookup(fields) {
  const map = {};
  for (const f of fields) {
    if (f.type === 'section' && Array.isArray(f.payload)) {
      map[f.field] = f.payload;
    } else if (f.payload && f.payload.value !== undefined) {
      map[f.field] = f.payload.value;
    }
  }
  return map;
}

const num = val => { const n = parseFloat(val); return isNaN(n) ? 0 : n; };

function sumSection(sectionArr, valueField) {
  if (!Array.isArray(sectionArr)) return 0;
  let total = 0;
  for (const row of sectionArr) {
    for (const f of row) {
      if (f.field === valueField && f.payload?.value) total += num(f.payload.value);
    }
  }
  return total;
}

function getSectionFieldValues(sectionArr, fieldName) {
  if (!Array.isArray(sectionArr)) return [];
  const vals = [];
  for (const row of sectionArr) {
    for (const f of row) {
      if (f.field === fieldName && f.payload?.value != null) vals.push(f.payload.value);
    }
  }
  return vals;
}

function getSectionRows(sectionArr, labelField, valueField) {
  if (!Array.isArray(sectionArr)) return [];
  const rows = [];
  for (const row of sectionArr) {
    let label = '', value = 0;
    for (const f of row) {
      if (f.field === labelField  && f.payload?.value != null) label = String(f.payload.value);
      if (f.field === valueField  && f.payload?.value != null) value = num(f.payload.value);
    }
    if (value !== 0 || label) rows.push({ label, value });
  }
  return rows;
}

/** ISO date string → midnight-UTC JS Date (strips time component) */
function toDate(iso) {
  if (!iso) return null;
  const dateOnly = String(iso).slice(0, 10);
  const d = new Date(dateOnly + 'T00:00:00Z');
  return isNaN(d.getTime()) ? null : d;
}

// ─── EXTRACT DATA ─────────────────────────────────────────────────────────────

function extractData(consentOrder) {
  const d = buildLookup(consentOrder);

  // Names — keep first name separate (templates use first name only in header cells)
  const petFirstName = d['Petitioner.Legal_first_name'] || '';
  const resFirstName = d['Respondent.Legal_first_name'] || '';
  const petName = [petFirstName, d['Petitioner.Legal_last_name'] || ''].join(' ').trim();
  const resName = [resFirstName, d['Respondent.Legal_last_name'] || ''].join(' ').trim();

  const petDobIso   = d['Matter.Petitioner_date_of_birth'] || '';
  const resDobIso   = d['Matter.Respondent_date_of_birth'] || '';
  const cohabIso    = d['Petitioner.Background_info_date_cohabiting'] || '';
  const marriageIso = d['Matter.Date_of_marriage'] || '';
  const sepIso      = d['Matter.Date_of_separation'] || '';
  const condIso     = d['Matter.Decree_nisi'] || '';
  const caseNumber  = d['Matter.Case_number'] || '';
  const petOcc      = d['Matter.Petitioner_occupation'] || '';
  const resOcc      = d['Matter.Respondent_occupation'] || '';

  // FMH address — full version for Assisted, short "Street. Town" for Negotiation header
  const _rawPetAddress = d['Petitioner.Current_address'] || '{}';
  const fmhAddress = (() => {
    try {
      const addr = JSON.parse(_rawPetAddress);
      return [addr.Street, addr.Town, addr.Postcode].filter(Boolean).join(', ');
    } catch { return ''; }
  })();
  const fmhAddressShort = (() => {
    try {
      const addr = JSON.parse(_rawPetAddress);
      return [addr.Street, addr.Town].filter(Boolean).join('. ');
    } catch { return ''; }
  })();

  // Property 2 address — parse from Respondent.Current_address JSON
  const prop2Address = (() => {
    try {
      const addr = JSON.parse(d['Respondent.Current_address'] || '{}');
      return [addr.Street, addr.Town, addr.Postcode].filter(Boolean).join(', ');
    } catch { return ''; }
  })();

  // Children
  const childrenSection = d['Children.Children_questions'] || [];
  const children = childrenSection.map(row => {
    let firstName = '', middleName = '', lastName = '', dob = '';
    for (const f of row) {
      if (f.field === 'Children.First_name')            firstName  = f.payload?.value || '';
      if (f.field === 'Children.Middle_name_child')     middleName = f.payload?.value || '';
      if (f.field === 'Children.Last_name')             lastName   = f.payload?.value || '';
      if (f.field === 'Children.Birth_day_first_child') dob        = f.payload?.value || '';
    }
    return { name: `${firstName} ${middleName} ${lastName}`.replace(/\s+/g, ' ').trim(), dob };
  });

  // Property values
  // Both parties recorded the same FMH — detect via Last_address_lived_together = "Yes"
  // Check both petitioner and respondent fields
  const petLastAddrTogether = d['Petitioner.Last_address_lived_together'] || '';
  const resLastAddrTogether = d['Respondent.Last_address_lived_together'] || '';
  const isSameFMH = petLastAddrTogether === 'Yes' || resLastAddrTogether === 'Yes';
  const fmhValue   = num(d['Petitioner.Property_value'] || 0);
  const prop2Value = isSameFMH ? 0 : num(d['Respondent.Property_value'] || 0);

  const petMortTotal = sumSection(d['Petitioner.Properties_mortgage_questions'] || [], 'Petitioner.Mortgage_value');
  const resMortTotal = isSameFMH ? petMortTotal : sumSection(d['Respondent.Properties_mortgage_questions'] || [], 'Respondent.Mortgage_value');

  const prop3Value = getSectionFieldValues(
    d['Respondent.Additional_property_list_owned'] || [],
    'Respondent.Additional_property_sole_agreed_valuation'
  ).reduce((a, v) => a + num(v), 0);

  const prop4Value = getSectionFieldValues(
    d['Respondent.Additional_property_list_owned_with_someone_questions'] || [],
    'Respondent.Additional_property_joint_agreed_valuation'
  ).reduce((a, v) => a + num(v), 0);

  // prop3Cgt: read from data if available, otherwise 0 (not hardcoded)
  const prop3Cgt = 0;

  const carParkEach = getSectionFieldValues(
    d['Petitioner.Additional_property_list_owned_with_someone_questions'] || [],
    'Petitioner.Additional_property_joint_agreed_valuation'
  ).reduce((a, v) => a + num(v), 0) / 2;

  // ── Petitioner assets ─────────────────────────────────────────────────────
  // Row-level arrays — use CORRECT field names from the JSON schema
  const petBankRows    = getSectionRows(d['Petitioner.Bank_accounts'] || [],
    'Petitioner.Bank_account_sole_name', 'Petitioner.Bank_account_sole_value');
  const petIsaRows     = getSectionRows(d['Petitioner.Isas_account'] || [],
    'Petitioner.Isas_provider_name', 'Petitioner.Isas_value');
  const petInvestRows  = getSectionRows(d['Petitioner.Investments'] || [],
    'Petitioner.Investments_type_shares_name', 'Petitioner.Investments_type_shares_value');
  const petVehRows     = getSectionRows(d['Petitioner.Vehicles_questions'] || [],
    'Petitioner.Vehicles_type', 'Petitioner.Vehicles_value');
  const petAddRows     = getSectionRows(d['Petitioner.Additional_assets_personal'] || [],
    'Petitioner.Additional_assets_description', 'Petitioner.Additional_assets_value');
  const petPensionRows = getSectionRows(d['Petitioner.Pensions_private_questions'] || [],
    'Petitioner.Pension_name', 'Petitioner.Pension_value');  // Pension_name = "Scottish Widows" etc
  const petCCRows      = getSectionRows(d['Petitioner.Credit_cards_info'] || [],
    'Petitioner.Credit_card_provider_name', 'Petitioner.Credit_card_amount');
  const petLoanRows    = getSectionRows(d['Petitioner.Personal_loans'] || [],
    'Petitioner.Loan_provider_name', 'Petitioner.Loan_value');

  // Totals
  const petBanks       = petBankRows.reduce((t, r) => t + r.value, 0);
  const petIsas        = petIsaRows.reduce((t, r) => t + r.value, 0);
  const petInvest      = petInvestRows.reduce((t, r) => t + r.value, 0);
  const petVehs        = petVehRows.reduce((t, r) => t + r.value, 0);
  const petAddAss      = petAddRows.reduce((t, r) => t + r.value, 0);
  const petPenTotal    = petPensionRows.reduce((t, r) => t + r.value, 0);
  const petCreditCards = petCCRows.reduce((t, r) => t + r.value, 0);
  const petLoans       = petLoanRows.reduce((t, r) => t + r.value, 0);

  const petVehFinance = (d['Petitioner.Vehicles_questions'] || []).reduce((t, row) => {
    for (const f of row) {
      if (f.field === 'Petitioner.Vehicles_finance_amount_left') t += num(f.payload?.value ?? 0);
    }
    return t;
  }, 0);

  // ── Respondent assets ─────────────────────────────────────────────────────
  // Respondent bank rows — filter out any account that is a joint account
  // (it will already appear via jointBankAssets from Petitioner.Joint_bank_account_questions)
  const resBankRowsRaw = getSectionRows(d['Respondent.Bank_accounts'] || [],
    'Respondent.Bank_account_sole_name', 'Respondent.Bank_account_sole_value');
  const resBankRows = resBankRowsRaw.filter(r => !r.label.toLowerCase().includes('joint'));
  const resIsaRows     = getSectionRows(d['Respondent.Isas_account'] || [],
    'Respondent.Isas_provider_name', 'Respondent.Isas_value');
  const resInvestRows  = getSectionRows(d['Respondent.Investments'] || [],
    'Respondent.Investments_type_shares_name', 'Respondent.Investments_type_shares_value');
  const resVehRows     = getSectionRows(d['Respondent.Vehicles_questions'] || [],
    'Respondent.Vehicles_type', 'Respondent.Vehicles_value');
  const resAddRows     = getSectionRows(d['Respondent.Additional_assets_personal'] || [],
    'Respondent.Additional_assets_description', 'Respondent.Additional_assets_value');
  const resBizRows     = getSectionRows(d['Respondent.Business_assets_questions'] || [],
    'Respondent.Business_assets_sole_name', 'Respondent.Business_assets_value');
  const resPensionRows = getSectionRows(d['Respondent.Pensions_private_questions'] || [],
    'Respondent.Pension_name', 'Respondent.Pension_value');  // Pension_name = "True Potential" etc
  const resCCRows      = getSectionRows(d['Respondent.Credit_cards_info'] || [],
    'Respondent.Credit_card_provider_name', 'Respondent.Credit_card_amount');
  const resLoanRows    = getSectionRows(d['Respondent.Personal_loans'] || [],
    'Respondent.Loan_provider_name', 'Respondent.Loan_value');
  // Tax liability rows — include the repayment date in the label (e.g. "Self Assessment (Jan 2027)")
  const resTaxLiabRows = (() => {
    const section = d['Respondent.Tax_liability'] || [];
    return section.map(row => {
      let label = '', date = '', value = 0;
      for (const f of row) {
        if (f.field === 'Respondent.Tax_liability_incurred' && f.payload?.value != null) label = String(f.payload.value);
        if (f.field === 'Respondent.Tax_liability_tax_date_repaid' && f.payload?.value != null) date = String(f.payload.value).trim();
        if (f.field === 'Respondent.Tax_liability_total_current_value' && f.payload?.value != null) value = num(f.payload.value);
      }
      const fullLabel = date ? `${label} (${date})` : label;
      return { label: fullLabel, value };
    }).filter(r => r.value !== 0 || r.label);
  })();

  // Totals
  const resBanks       = resBankRows.reduce((t, r) => t + r.value, 0);
  const resIsas        = resIsaRows.reduce((t, r) => t + r.value, 0);
  const resInvest      = resInvestRows.reduce((t, r) => t + r.value, 0);
  const resVehs        = resVehRows.reduce((t, r) => t + r.value, 0);
  const resAddAss      = resAddRows.reduce((t, r) => t + r.value, 0);
  const resBiz         = resBizRows.reduce((t, r) => t + r.value, 0);
  const resPenTotal    = resPensionRows.reduce((t, r) => t + r.value, 0);
  const resCreditCards = resCCRows.reduce((t, r) => t + r.value, 0);
  const resLoans       = resLoanRows.reduce((t, r) => t + r.value, 0);
  const resTaxLiab     = resTaxLiabRows.reduce((t, r) => t + r.value, 0);

  const resVehFinance = (d['Respondent.Vehicles_questions'] || []).reduce((t, row) => {
    for (const f of row) {
      if (f.field === 'Respondent.Vehicles_finance_amount_left') t += num(f.payload?.value ?? 0);
    }
    return t;
  }, 0);

  // ── Children assets ───────────────────────────────────────────────────────
  const childBankRows = getSectionRows(d['Children.Bank_accounts'] || [],
    'Children.Bank_name', 'Children.Bank_account_value');
  const childIsaRows  = getSectionRows(d['Children.Isas_accounts'] || [],
    'Children.Isas_provider_name', 'Children.Isa_value');
  const childAddRows  = getSectionRows(d['Children.Additional_assets'] || [],
    'Children.Additional_assets_description', 'Children.Additional_assets_value');

  const childBanks = childBankRows.reduce((t, r) => t + r.value, 0);
  const childIsas  = childIsaRows.reduce((t, r) => t + r.value, 0);
  const childAdd   = childAddRows.reduce((t, r) => t + r.value, 0);

  // Joint assets — from joint bank accounts (Petitioner.Joint_bank_account_questions)
  // Each row has: Petitioner.Joint_bank_name, Petitioner.Joint_account_share (pet's share)
  // The joint account share is written as pet share in J34, res share in N34
  const petJointBankSection = d['Petitioner.Joint_bank_account_questions'] || [];
  const jointBankAssets = petJointBankSection.map(row => {
    let label = '', petShare = 0, totalVal = 0;
    for (const f of row) {
      if (f.field === 'Petitioner.Joint_bank_name') label = f.payload?.value || '';
      if (f.field === 'Petitioner.Joint_account_share') petShare = num(f.payload?.value ?? 0);
      if (f.field === 'Petitioner.Joint_account_overall_value') totalVal = num(f.payload?.value ?? 0);
    }
    // If no explicit share, split 50/50
    const resShare = petShare !== 0 ? totalVal - petShare : totalVal / 2;
    return { label: label.trim(), petShare, resShare };
  });

  // Also include any property-based joint assets (car park etc.)
  const petJointPropSection = d['Petitioner.Additional_property_list_owned_with_someone_questions'] || [];
  const jointPropAssets = petJointPropSection.map(row => {
    let label = '', value = 0;
    for (const f of row) {
      if (f.field === 'Petitioner.Additional_property_joint_address') {
        try { const a = JSON.parse(f.payload?.value || '{}'); label = a.Street || ''; } catch { label = ''; }
      }
      if (f.field === 'Petitioner.Additional_property_joint_agreed_valuation') value = num(f.payload?.value ?? 0);
    }
    return { label, petShare: value / 2, resShare: value / 2 };
  });

  const jointAssets = [...jointBankAssets, ...jointPropAssets];

  // ── Income now ────────────────────────────────────────────────────────────
  const petSalary   = num(d['Petitioner.Income_net_monthly'] || 0);
  const petBenefits = num(d['Petitioner.Benefits'] || 0);
  const petStatePen = num(d['Petitioner.Pensions_monthly_income'] || 0);
  const petPenInc   = num(d['Petitioner.Pension_payments'] || 0);
  const petBankInt  = num(d['Petitioner.Bank_interest'] || 0);
  const petOtherInc = num(d['Petitioner.Income_other_sources'] || 0);
  const petRental   = num(d['Petitioner.Income_net_rental_value'] || 0);

  const resSalary   = num(d['Respondent.Income_net_monthly'] || 0);
  const resBenefits = num(d['Respondent.Benefits'] || 0);
  const resStatePen = num(d['Respondent.Pensions_monthly_income'] || 0);
  const resPenInc   = num(d['Respondent.Pension_payments'] || 0);
  const resBankInt  = num(d['Respondent.Bank_interest'] || 0);
  const resOtherInc = num(d['Respondent.Income_other_sources'] || 0);
  const resRental   = num(d['Respondent.Income_net_rental_value'] || 0);

  // ── Income / capital after ────────────────────────────────────────────────
  const petOtherCapAfter = num(d['D81.Other_capital_app_after'] || 0);
  const resOtherCapAfter = num(d['D81.Other_capital_res_after'] || 0);
  const petSalaryAfter   = num(d['Petitioner.Income_net_monthly_after'] || 0);
  const petBenAfter      = num(d['Petitioner.Benefits_after'] || 0);
  const petPenIncAfter   = num(d['Petitioner.Pension_payments_after'] || 0);
  const petBankIntAfter  = num(d['Petitioner.Bank_interest_after'] || 0);
  const petOtherIncAfter = num(d['Petitioner.Income_other_sources_after'] || 0);
  const resSalaryAfter   = num(d['Respondent.Income_net_monthly_after'] || 0);
  const resBenAfter      = num(d['Respondent.Benefits_after'] || 0);
  const resPenIncAfter   = num(d['Respondent.Pension_payments_after'] || 0);
  const resBankIntAfter  = num(d['Respondent.Bank_interest_after'] || 0);
  const resOtherIncAfter = num(d['Respondent.Income_other_sources_after'] || 0);

  const commentary  = d['D81.Other_information_CO_main_reason'] || '';
  const petLumpSum  = num(d['D81.Lump_sum_payable_app'] || d['Petitioner.Lump_sum'] || 0);
  const resLumpSum  = num(d['D81.Lump_sum_payable_res'] || d['Respondent.Lump_sum'] || 0);

  return {
    petName, resName, petFirstName, resFirstName, petOcc, resOcc, caseNumber,
    petDobIso, resDobIso, cohabIso, marriageIso, sepIso, condIso,
    fmhAddress, fmhAddressShort, prop2Address,
    children,
    fmhValue, prop2Value, prop3Value, prop4Value, prop3Cgt,
    petMortTotal, resMortTotal, carParkEach, jointAssets,
    // Row arrays (Negotiation)
    petBankRows, petIsaRows, petInvestRows, petVehRows, petAddRows,
    petPensionRows, petCCRows, petLoanRows,
    resBankRows, resIsaRows, resInvestRows, resVehRows, resAddRows, resBizRows,
    resPensionRows, resCCRows, resLoanRows, resTaxLiabRows,
    // Totals (Assisted)
    petBanks, petIsas, petInvest, petVehs, petAddAss,
    petPenTotal, petCreditCards, petLoans, petVehFinance,
    resBanks, resIsas, resInvest, resVehs, resAddAss, resBiz,
    resPenTotal, resCreditCards, resLoans, resTaxLiab, resVehFinance,
    childBanks, childIsas, childAdd, childBankRows, childIsaRows, childAddRows,
    petSalary, petBenefits, petStatePen, petPenInc, petBankInt, petOtherInc, petRental,
    resSalary, resBenefits, resStatePen, resPenInc, resBankInt, resOtherInc, resRental,
    petOtherCapAfter, resOtherCapAfter,
    petSalaryAfter, petBenAfter, petPenIncAfter, petBankIntAfter, petOtherIncAfter,
    resSalaryAfter, resBenAfter, resPenIncAfter, resBankIntAfter, resOtherIncAfter,
    petLumpSum, resLumpSum,
    commentary,
  };
}

// ─── FILL ASSISTED TEMPLATE ──────────────────────────────────────────────────

async function fillAssistedTemplate(templateBuffer, data) {
  const wb = await XlsxPopulate.fromDataAsync(templateBuffer);
  const ws = wb.sheet('Financial Disclosure');
  if (!ws) throw new Error('Sheet "Financial Disclosure" not found');

  const w  = (ref, val) => { if (val !== null && val !== undefined) ws.cell(ref).value(val); };
  const wd = (ref, iso) => { const d = toDate(iso); if (d) ws.cell(ref).value(d); };

  // Names — first name only (drives =B5/=C5 references throughout)
  w('B5', data.petFirstName);
  w('C5', data.resFirstName);

  // Dates
  wd('H3', data.petDobIso);   w('J3', data.petOcc);
  wd('H4', data.resDobIso);   w('J4', data.resOcc);
  wd('L3', data.cohabIso);
  wd('L4', data.marriageIso);
  wd('L5', data.sepIso);
  wd('L6', data.condIso);

  // Children — name in G5 only (G6/G7 are template labels), DOB in H5-H7
  data.children.slice(0, 3).forEach((child, i) => {
    if (i === 0) w('G5', child.name);
    wd(`H${5 + i}`, child.dob);
  });

  // Properties — write equity (value minus mortgage)
  const fmhEquity   = Math.max(0, data.fmhValue - data.petMortTotal);
  const prop2Equity = Math.max(0, data.prop2Value - data.resMortTotal);
  w('B6', fmhEquity / 2);
  w('C6', fmhEquity / 2);
  w('B9', 0);
  w('C9', prop2Equity);

  // Other assets — fixed rows matching template pre-labelled rows
  w('B16', data.petVehs);                                            // Vehicles
  w('C16', data.resVehs);
  data.petBankRows.slice(0, 3).forEach((b, i) => w(`B${18+i}`, b.value)); // Banks 18-20
  data.resBankRows.slice(0, 3).forEach((b, i) => w(`C${18+i}`, b.value));
  w('B21', data.petIsas);                                            // ISAs
  w('C21', data.resIsas);
  w('B22', data.petInvest);                                          // Shares
  w('C22', data.resInvest);
  w('C23', data.resBiz);                                             // Business
  w('B24', data.petAddAss);                                          // Other assets
  w('C24', data.resAddAss);
  w('B25', data.jointAssets.reduce((t, a) => t + a.petShare, 0));  // Joint assets — pet share
  w('C25', data.jointAssets.reduce((t, a) => t + a.resShare, 0)); // Joint assets — res share

  // Liabilities — fixed rows
  w('B32', data.petCreditCards);
  w('C32', data.resCreditCards);
  w('B33', data.petLoans);
  w('C33', data.resLoans);
  w('C34', data.resTaxLiab);
  w('B36', data.petVehFinance);
  w('C36', data.resVehFinance);

  // Children assets — rows 37-38
  const childAssets = [...data.childBankRows, ...data.childIsaRows, ...data.childAddRows];
  childAssets.slice(0, 2).forEach((a, i) => {
    w(`F${37+i}`, a.label);
    w(`G${37+i}`, a.value);
  });

  // Pensions — fixed rows 44-52
  data.petPensionRows.slice(0, 9).forEach((p, i) => w(`B${44+i}`, p.value));
  data.resPensionRows.slice(0, 9).forEach((p, i) => w(`C${44+i}`, p.value));

  // Income now — fixed rows 59-66
  w('B59', data.petSalary);    w('C59', data.resSalary);
  w('B60', data.petBenefits);  w('C60', data.resBenefits);
  w('B61', data.petStatePen);  w('C61', data.resStatePen);
  w('B62', data.petPenInc);    w('C62', data.resPenInc);
  w('B63', data.petBankInt);   w('C63', data.resBankInt);
  w('B66', data.petRental);    w('C66', data.resRental);

  // Net effect — G72/I72 are the only inputs; all other G/I are template formulas
  w('G72', fmhEquity / 2);
  w('I72', fmhEquity / 2 + prop2Equity);

  return wb.outputAsync();
}

// ─── FILL NEGOTIATION TEMPLATE ────────────────────────────────────────────────

async function fillNegotiationTemplate(templateBuffer, data) {
  const wb = await XlsxPopulate.fromDataAsync(templateBuffer);

  // Find the correct sheet — handles both naming conventions:
  // Old Cattell: "2. Assets, Debts & Net Effect "
  // New template: "Assets, Debts & Net Effect "
  const ws = wb.sheets().find(s => {
    const n = s.name().trim();
    return n.startsWith('2. Assets') || n.startsWith('Assets, Debts');
  });
  if (!ws) {
    const names = wb.sheets().map(s => `"${s.name()}"`).join(', ');
    throw new Error(`Negotiation sheet not found. Available: ${names}. Please upload the correct template.`);
  }

  const w  = (ref, val) => { if (val !== null && val !== undefined) ws.cell(ref).value(val); };
  const wd = (ref, iso) => { const d = toDate(iso); if (d) ws.cell(ref).value(d); };


  // ── Names — first name only (B2/B3 drive all =B2/=B3 refs throughout) ─────
  w('B2', data.petFirstName || data.petName);
  w('B3', data.resFirstName || data.resName);
  w('D3', data.caseNumber ? `${data.caseNumber} - ` : '');

  // ── "Other assets" sub-table headers use first names ─────────────────────
  w('H15', `Other assets - ${data.petFirstName || data.petName}`);
  w('L15', `Other assets - ${data.resFirstName || data.resName}`);

  // ── Property address labels ───────────────────────────────────────────────
  // H6 = FMH short address: "Street. Town" format matching Cattell "9 Patterdale Close. Crewe"
  w('H6', data.fmhAddressShort || data.fmhAddress || 'Family Home');
  // L6: only overwrite if there is actually a second property
  if (data.prop2Value > 0) w('L6', data.prop2Address || 'Property 2');

  // ── Property values ───────────────────────────────────────────────────────
  // J7=FMH value, J8=FMH mortgage, J9=formula in template (leave), J10=blank
  // N7/N8 only if there is a real second property
  w('J7', data.fmhValue);
  w('J8', data.petMortTotal);
  if (data.prop2Value > 0) {
    w('N7', data.prop2Value);
    w('N8', data.resMortTotal);
  }

  // ── Petitioner sole assets — rows 16–29 ──────────────────────────────────
  // Cattell order: banks (J col), investments (J col), add. assets (K+J cols), vehicles (K+J cols)
  let petRow = 16;
  data.petBankRows.forEach(r => {
    if (petRow > 29) return;
    w(`H${petRow}`, r.label);  w(`J${petRow}`, r.value);
    petRow++;
  });
  [...data.petInvestRows, ...data.petIsaRows].forEach(r => {
    if (petRow > 29) return;
    w(`H${petRow}`, r.label);  w(`J${petRow}`, r.value);
    petRow++;
  });
  data.petAddRows.forEach(r => {
    if (petRow > 29) return;
    w(`H${petRow}`, r.label);  w(`K${petRow}`, r.value);  w(`J${petRow}`, r.value);
    petRow++;
  });
  data.petVehRows.forEach(r => {
    if (petRow > 29) return;
    w(`H${petRow}`, r.label);  w(`K${petRow}`, r.value);  w(`J${petRow}`, r.value);
    petRow++;
  });

  // ── Respondent sole assets — rows 16–29 ──────────────────────────────────
  // Cattell order: banks (N col), investments/biz (N col), add. assets (O+N cols), vehicles (O+N cols)
  let resRow = 16;
  data.resBankRows.forEach(r => {
    if (resRow > 29) return;
    w(`L${resRow}`, r.label);  w(`N${resRow}`, r.value);
    resRow++;
  });
  [...data.resInvestRows, ...data.resIsaRows, ...data.resBizRows].forEach(r => {
    if (resRow > 29) return;
    w(`L${resRow}`, r.label);  w(`N${resRow}`, r.value);
    resRow++;
  });
  data.resAddRows.forEach(r => {
    if (resRow > 29) return;
    w(`L${resRow}`, r.label);  w(`O${resRow}`, r.value);  w(`N${resRow}`, r.value);
    resRow++;
  });
  data.resVehRows.forEach(r => {
    if (resRow > 29) return;
    w(`L${resRow}`, r.label);  w(`O${resRow}`, r.value);  w(`N${resRow}`, r.value);
    resRow++;
  });

  // ── Joint assets — rows 34–45 (H+L=label, J=pet share, N=res share) ──────
  data.jointAssets.slice(0, 12).forEach((asset, i) => {
    w(`H${34 + i}`, asset.label);  w(`J${34 + i}`, asset.petShare);
    w(`L${34 + i}`, asset.label);  w(`N${34 + i}`, Math.round(asset.resShare));
  });

  // ── Children assets — P/R cols rows 16–20, only if value > 0 ─────────────
  const childAssets = [...data.childBankRows, ...data.childIsaRows, ...data.childAddRows]
    .filter(a => a.value > 0);
  childAssets.slice(0, 5).forEach((a, i) => {
    w(`P${16 + i}`, a.label);
    w(`R${16 + i}`, a.value);
  });

  // ── Income now — B22/D22=salary, B23/D23=state benefits ─────────────────
  w('B22', data.petSalary);
  w('D22', Math.round(data.resSalary));  // Cattell shows 2998 not 2998.35
  w('B23', data.petBenefits);
  w('D23', data.resBenefits);

  // ── Petitioner liabilities — H=label, J=value, rows 51–58 ───────────────
  let petLiabRow = 51;
  if (data.petVehFinance > 0) {
    w(`H${petLiabRow}`, 'Car finance ');
    w(`J${petLiabRow}`, data.petVehFinance);
    petLiabRow++;
  }
  [...data.petCCRows, ...data.petLoanRows].filter(r => r.value > 0).forEach(liab => {
    if (petLiabRow > 58) return;
    w(`H${petLiabRow}`, liab.label);
    w(`J${petLiabRow}`, liab.value);
    petLiabRow++;
  });

  // ── Respondent liabilities — L=label, N=value, rows 51–58 ───────────────
  // Order: car finance, then credit cards/loans, then tax last
  let resLiabRow = 51;
  if (data.resVehFinance > 0) {
    w(`L${resLiabRow}`, 'Car finance ');
    w(`N${resLiabRow}`, data.resVehFinance);
    resLiabRow++;
  }
  [...data.resCCRows, ...data.resLoanRows].forEach(liab => {
    if (resLiabRow > 58) return;
    w(`L${resLiabRow}`, liab.label);
    w(`N${resLiabRow}`, Math.round(liab.value));
    resLiabRow++;
  });
  data.resTaxLiabRows.forEach(liab => {
    if (resLiabRow > 58) return;
    // Format: "Tax liability Jan 2027 " (extract date from parentheses)
    const dateMatch = liab.label.match(/\(([^)]+)\)/);
    const taxLabel = dateMatch ? `Tax liability ${dateMatch[1]} ` : liab.label;
    w(`L${resLiabRow}`, taxLabel);
    w(`N${resLiabRow}`, Math.round(liab.value));
    resLiabRow++;
  });

  // ── Pensions — H=label, J=value rows 62–68 pet; L=label, N=value rows 62–68 res ──
  data.petPensionRows.slice(0, 7).forEach((pen, i) => {
    w(`H${62 + i}`, pen.label);
    w(`J${62 + i}`, pen.value);
  });
  data.resPensionRows.slice(0, 7).forEach((pen, i) => {
    w(`L${62 + i}`, pen.label);
    w(`N${62 + i}`, Math.round(pen.value));
  });

  // ── Income after (future) — only write if non-zero ────────────────────────
  if (data.petSalaryAfter)  w('B58', data.petSalaryAfter);
  if (data.resSalaryAfter)  w('D58', data.resSalaryAfter);
  if (data.petBenAfter)     w('B59', data.petBenAfter);
  if (data.resBenAfter)     w('D59', data.resBenAfter);
  if (data.petPenIncAfter)  w('B60', data.petPenIncAfter);
  if (data.resPenIncAfter)  w('D60', data.resPenIncAfter);
  if (data.petBankIntAfter) w('B61', data.petBankIntAfter);
  if (data.resBankIntAfter) w('D61', data.resBankIntAfter);

  // ── Dates — rows 77–80 ───────────────────────────────────────────────────
  wd('B77', data.petDobIso);  w('D77', data.petOcc);
  wd('B78', data.resDobIso);  w('D78', data.resOcc);
  w('E77', 'Cohabitation');   wd('F77', data.cohabIso);
  w('E78', 'Marriage ');      wd('F78', data.marriageIso);
  w('E79', 'Separation ');    wd('F79', data.sepIso);
  w('E80', 'CDO application date');
  if (data.condIso) wd('F80', data.condIso);

  // Children — A=name, B=DOB, rows 79–81
  data.children.slice(0, 3).forEach((child, i) => {
    w(`A${79 + i}`, child.name);
    wd(`B${79 + i}`, child.dob);
  });

  if (data.commentary) w('B84', data.commentary);

  return wb.outputAsync();
}

module.exports = { extractData, fillAssistedTemplate, fillNegotiationTemplate };
