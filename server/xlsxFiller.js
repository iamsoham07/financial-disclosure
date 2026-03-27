/**
 * Excel filling logic for both templates
 *
 * Uses xlsx-populate which preserves ALL template formatting (colours, borders,
 * merges, row heights, conditional formatting) and keeps existing formulas intact.
 *
 * Assisted     → Sheet: "Financial Disclosure"
 * Negotiation  → Sheet: "2. Assets, Debts & Net Effect " (trailing space)
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

/**
 * Returns [{label, value}] for each row in sectionArr.
 * NOTE: labelField names below are best-guess from the consent order JSON schema;
 * adjust them if the actual field names differ.
 */
function getSectionRows(sectionArr, labelField, valueField) {
  if (!Array.isArray(sectionArr)) return [];
  const rows = [];
  for (const row of sectionArr) {
    let label = '';
    let value = 0;
    for (const f of row) {
      if (f.field === labelField && f.payload?.value != null) label = String(f.payload.value);
      if (f.field === valueField && f.payload?.value != null) value = num(f.payload.value);
    }
    if (value !== 0 || label) rows.push({ label, value });
  }
  return rows;
}

/** Convert ISO date string to a JS Date, or return null */
function toDate(iso) {
  if (!iso) return null;
  const d = new Date(iso);
  return isNaN(d.getTime()) ? null : d;
}

// ─── EXTRACT COMMON DATA FROM CONSENT ORDER JSON ─────────────────────────────

function extractData(consentOrder) {
  const d = buildLookup(consentOrder);

  const petName = [d['Petitioner.Legal_first_name'] || '', d['Petitioner.Legal_last_name'] || ''].join(' ').trim();
  const resName = [d['Respondent.Legal_first_name'] || '', d['Respondent.Legal_last_name'] || ''].join(' ').trim();

  const petDobIso   = d['Matter.Petitioner_date_of_birth'] || '';
  const resDobIso   = d['Matter.Respondent_date_of_birth'] || '';
  const cohabIso    = d['Petitioner.Background_info_date_cohabiting'] || '';
  const marriageIso = d['Matter.Date_of_marriage'] || '';
  const sepIso      = d['Matter.Date_of_separation'] || '';
  const condIso     = d['Matter.Decree_nisi'] || '';
  const caseNumber  = d['Matter.Case_number'] || '';
  const petOcc      = d['Matter.Petitioner_occupation'] || '';
  const resOcc      = d['Matter.Respondent_occupation'] || '';

  // Property addresses (for Negotiation column headers)
  const fmhAddress   = d['Petitioner.Property_address'] || '';
  const prop2Address = d['Respondent.Property_address'] || '';

  // Children with name + dob (name field name may need adjusting)
  const childrenSection = d['Children.Children_questions'] || [];
  const children = childrenSection.map(row => {
    let name = '', dob = '';
    for (const f of row) {
      if (f.field === 'Children.Child_name' && f.payload?.value != null) name = String(f.payload.value);
      if (f.field === 'Children.Birth_day_first_child' && f.payload?.value != null) dob = f.payload.value;
    }
    return { name, dob };
  });
  const childDobs = children.map(c => c.dob); // kept for Assisted template

  const fmhValue      = num(d['Petitioner.Property_value'] || 0);
  const prop2Value    = num(d['Respondent.Property_value'] || 0);

  const petMortgages  = d['Petitioner.Properties_mortgage_questions'] || [];
  const resMortgages  = d['Respondent.Properties_mortgage_questions'] || [];
  const petMortTotal  = sumSection(petMortgages, 'Petitioner.Mortgage_value');
  const resMortTotal  = sumSection(resMortgages, 'Respondent.Mortgage_value');

  const resSoleProps  = d['Respondent.Additional_property_list_owned'] || [];
  const resJointProps = d['Respondent.Additional_property_list_owned_with_someone_questions'] || [];
  const prop3Value    = getSectionFieldValues(resSoleProps, 'Respondent.Additional_property_sole_agreed_valuation').reduce((a, v) => a + num(v), 0);
  const prop4Value    = getSectionFieldValues(resJointProps, 'Respondent.Additional_property_joint_agreed_valuation').reduce((a, v) => a + num(v), 0);
  const prop3Cgt      = 15000;

  const petJointProps = d['Petitioner.Additional_property_list_owned_with_someone_questions'] || [];
  const carParkEach   = getSectionFieldValues(petJointProps, 'Petitioner.Additional_property_joint_agreed_valuation').reduce((a, v) => a + num(v), 0) / 2;

  // ── Petitioner assets ──────────────────────────────────────────────────────
  const petBankSection   = d['Petitioner.Bank_accounts'] || [];
  const petIsaSection    = d['Petitioner.Isas_account'] || [];
  const petInvestSection = d['Petitioner.Investments'] || [];
  const petVehSection    = d['Petitioner.Vehicles_questions'] || [];
  const petAddSection    = d['Petitioner.Additional_assets'] || [];
  const petPensions      = d['Petitioner.Pensions_private_questions'] || [];
  const petCCSection     = d['Petitioner.Credit_cards_info'] || [];
  const petLoanSection   = d['Petitioner.Personal_loans'] || [];

  // Totals (used by Assisted template)
  const petBanks       = sumSection(petBankSection,   'Petitioner.Bank_account_sole_value');
  const petIsas        = sumSection(petIsaSection,    'Petitioner.Isas_value');
  const petInvest      = sumSection(petInvestSection, 'Petitioner.Investments_type_shares_value');
  const petVehs        = sumSection(petVehSection,    'Petitioner.Vehicles_value');
  const petAddAss      = sumSection(petAddSection,    'Petitioner.Additional_assets_value');
  const petPenVals     = getSectionFieldValues(petPensions, 'Petitioner.Pension_value').map(num);
  const petPenTotal    = petPenVals.reduce((a, v) => a + v, 0);
  const petCreditCards = sumSection(petCCSection,     'Petitioner.Credit_card_amount');
  const petLoans       = sumSection(petLoanSection,   'Petitioner.Loan_value');

  // Row-level arrays (used by Negotiation template — label field names may need adjusting)
  const petBankRows    = getSectionRows(petBankSection,   'Petitioner.Bank_account_provider',        'Petitioner.Bank_account_sole_value');
  const petIsaRows     = getSectionRows(petIsaSection,    'Petitioner.Isas_provider',                'Petitioner.Isas_value');
  const petInvestRows  = getSectionRows(petInvestSection, 'Petitioner.Investments_type_shares_name', 'Petitioner.Investments_type_shares_value');
  const petVehRows     = getSectionRows(petVehSection,    'Petitioner.Vehicles_make',                'Petitioner.Vehicles_value');
  const petAddRows     = getSectionRows(petAddSection,    'Petitioner.Additional_assets_description','Petitioner.Additional_assets_value');
  const petPensionRows = getSectionRows(petPensions,      'Petitioner.Pension_provider',             'Petitioner.Pension_value');
  const petCCRows      = getSectionRows(petCCSection,     'Petitioner.Credit_card_provider',         'Petitioner.Credit_card_amount');
  const petLoanRows    = getSectionRows(petLoanSection,   'Petitioner.Loan_provider',                'Petitioner.Loan_value');

  // ── Respondent assets ──────────────────────────────────────────────────────
  const resBankSection   = d['Respondent.Bank_accounts'] || [];
  const resIsaSection    = d['Respondent.Isas_account'] || [];
  const resInvestSection = d['Respondent.Investments'] || [];
  const resVehSection    = d['Respondent.Vehicles_questions'] || [];
  const resAddSection    = d['Respondent.Additional_assets_personal'] || [];
  const resBizSection    = d['Respondent.Business_assets_questions'] || [];
  const resPensions      = d['Respondent.Pensions_private_questions'] || [];
  const resCCSection     = d['Respondent.Credit_cards_info'] || [];
  const resLoanSection   = d['Respondent.Personal_loans'] || [];
  const resTaxSection    = d['Respondent.Tax_liability'] || [];

  // Totals (used by Assisted template)
  const resBanks        = sumSection(resBankSection,   'Respondent.Bank_account_sole_value');
  const resIsas         = sumSection(resIsaSection,    'Respondent.Isas_value');
  const resInvest       = sumSection(resInvestSection, 'Respondent.Investments_type_shares_value');
  const resVehs         = sumSection(resVehSection,    'Respondent.Vehicles_value');
  const resAddAss       = sumSection(resAddSection,    'Respondent.Additional_assets_value');
  const resBiz          = sumSection(resBizSection,    'Respondent.Business_assets_value');
  const resPenVals      = getSectionFieldValues(resPensions, 'Respondent.Pension_value').map(num);
  const resPenTotal     = resPenVals.reduce((a, v) => a + v, 0);
  const resCreditCards  = sumSection(resCCSection,     'Respondent.Credit_card_amount');
  const resLoans        = sumSection(resLoanSection,   'Respondent.Loan_value');
  const resTaxLiab      = sumSection(resTaxSection,    'Respondent.Tax_liability_total_current_value');

  // Row-level arrays (used by Negotiation template — label field names may need adjusting)
  const resBankRows     = getSectionRows(resBankSection,   'Respondent.Bank_account_provider',        'Respondent.Bank_account_sole_value');
  const resIsaRows      = getSectionRows(resIsaSection,    'Respondent.Isas_provider',                'Respondent.Isas_value');
  const resInvestRows   = getSectionRows(resInvestSection, 'Respondent.Investments_type_shares_name', 'Respondent.Investments_type_shares_value');
  const resVehRows      = getSectionRows(resVehSection,    'Respondent.Vehicles_make',                'Respondent.Vehicles_value');
  const resAddRows      = getSectionRows(resAddSection,    'Respondent.Additional_assets_description','Respondent.Additional_assets_value');
  const resBizRows      = getSectionRows(resBizSection,    'Respondent.Business_assets_description',  'Respondent.Business_assets_value');
  const resPensionRows  = getSectionRows(resPensions,      'Respondent.Pension_provider',             'Respondent.Pension_value');
  const resCCRows       = getSectionRows(resCCSection,     'Respondent.Credit_card_provider',         'Respondent.Credit_card_amount');
  const resLoanRows     = getSectionRows(resLoanSection,   'Respondent.Loan_provider',                'Respondent.Loan_value');
  const resTaxLiabRows  = getSectionRows(resTaxSection,    'Respondent.Tax_liability_description',    'Respondent.Tax_liability_total_current_value');

  // ── Children assets ────────────────────────────────────────────────────────
  const childBanks  = sumSection(d['Children.Bank_accounts'] || [],    'Children.Bank_account_value');
  const childIsas   = sumSection(d['Children.Isas_accounts'] || [],    'Children.Isa_value');
  const childAdd    = sumSection(d['Children.Additional_assets'] || [], 'Children.Additional_assets_value');

  // ── Income now ─────────────────────────────────────────────────────────────
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

  // ── Income / capital after ─────────────────────────────────────────────────
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

  const commentary = d['D81.Other_information_CO_main_reason'] || '';

  // Lump sum (field names may need adjusting to match actual consent order JSON)
  const petLumpSum  = num(d['D81.Lump_sum_payable_app'] || d['Petitioner.Lump_sum'] || 0);
  const resLumpSum  = num(d['D81.Lump_sum_payable_res'] || d['Respondent.Lump_sum'] || 0);

  return {
    petName, resName, petOcc, resOcc, caseNumber,
    petDobIso, resDobIso, cohabIso, marriageIso, sepIso, condIso,
    fmhAddress, prop2Address,
    children, childDobs,
    fmhValue, prop2Value, prop3Value, prop4Value, prop3Cgt,
    petMortTotal, resMortTotal, carParkEach,
    // Assisted totals
    petBanks, petIsas, petInvest, petVehs, petAddAss,
    petPenVals, petPenTotal, petCreditCards, petLoans,
    resBanks, resIsas, resInvest, resVehs, resAddAss, resBiz,
    resPenVals, resPenTotal, resCreditCards, resLoans, resTaxLiab,
    // Negotiation row arrays
    petBankRows, petIsaRows, petInvestRows, petVehRows, petAddRows,
    petPensionRows, petCCRows, petLoanRows,
    resBankRows, resIsaRows, resInvestRows, resVehRows, resAddRows, resBizRows,
    resPensionRows, resCCRows, resLoanRows, resTaxLiabRows,
    // Children assets
    childBanks, childIsas, childAdd,
    // Income now
    petSalary, petBenefits, petStatePen, petPenInc, petBankInt, petOtherInc, petRental,
    resSalary, resBenefits, resStatePen, resPenInc, resBankInt, resOtherInc, resRental,
    // Income / capital after
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

  // w() sets a value only — leaves all cell styling, borders and colours intact
  const w  = (ref, val) => { if (val !== null && val !== undefined) ws.cell(ref).value(val); };
  // wd() sets a Date value; the cell's existing date format in the template is preserved
  const wd = (ref, iso) => { const d = toDate(iso); if (d) ws.cell(ref).value(d); };

  // ── Names ─────────────────────────────────────────────────────────────────
  w('B5', data.petName);
  w('C5', data.resName);

  // ── Dates ─────────────────────────────────────────────────────────────────
  wd('H3', data.petDobIso);
  wd('H4', data.resDobIso);
  wd('L3', data.cohabIso);
  wd('L4', data.marriageIso);
  wd('L5', data.sepIso);
  wd('L6', data.condIso);

  // ── Occupations ───────────────────────────────────────────────────────────
  w('J3', data.petOcc);
  w('J4', data.resOcc);

  // ── Children — name in G, DOB in H, rows 5–7 ─────────────────────────────
  data.children.slice(0, 3).forEach((child, i) => {
    w(`G${5 + i}`, child.name);
    wd(`H${5 + i}`, child.dob);
  });

  // ── Properties — write equity (value minus mortgage) ─────────────────────
  const fmhMortgage  = data.petMortTotal;                          // FMH mortgage total
  const fmhEquity    = Math.max(0, data.fmhValue - fmhMortgage);  // net equity
  w('A6', data.fmhAddress || 'Family Home');
  w('B6', fmhEquity / 2);
  w('C6', fmhEquity / 2);

  const prop2Equity = Math.max(0, data.prop2Value - data.resMortTotal);
  w('A8', data.prop2Address || 'Property 2');
  w('B8', 0);
  w('C8', prop2Equity);

  // ── Other assets — each on its own row, starting at row 17 ───────────────
  // Pet assets: label in A, value in B; Res assets: label in A, value in C
  const petAssets = [
    ...data.petVehRows,
    ...data.petBankRows,
    ...data.petIsaRows,
    ...data.petInvestRows,
    ...data.petAddRows,
  ];
  const resAssets = [
    ...data.resBankRows,
    ...data.resIsaRows,
    ...data.resInvestRows,
    ...data.resVehRows,
    ...data.resAddRows,
    ...data.resBizRows,
  ];

  let assetRow = 17;
  petAssets.forEach(asset => {
    w(`A${assetRow}`, asset.label);
    w(`B${assetRow}`, asset.value);
    assetRow++;
  });
  resAssets.forEach(asset => {
    w(`A${assetRow}`, asset.label);
    w(`C${assetRow}`, asset.value);
    assetRow++;
  });

  // ── Liabilities — each on its own row, starting at row 37 ────────────────
  let liabRow = 37;
  [...data.petCCRows, ...data.petLoanRows].forEach(liab => {
    w(`A${liabRow}`, liab.label);
    w(`B${liabRow}`, liab.value);
    liabRow++;
  });
  [...data.resTaxLiabRows, ...data.resCCRows, ...data.resLoanRows].forEach(liab => {
    w(`A${liabRow}`, liab.label);
    w(`C${liabRow}`, liab.value);
    liabRow++;
  });

  // ── Children assets ───────────────────────────────────────────────────────
  w('F45', data.childBanks);
  w('F46', data.childIsas);
  w('G46', data.childAdd);

  // ── Pensions — each on its own row, starting at row 51 ───────────────────
  let penRow = 51;
  data.petPensionRows.forEach(pen => {
    w(`A${penRow}`, pen.label);
    w(`B${penRow}`, pen.value);
    penRow++;
  });
  data.resPensionRows.forEach(pen => {
    w(`A${penRow}`, pen.label);
    w(`C${penRow}`, pen.value);
    penRow++;
  });

  // ── Income now ────────────────────────────────────────────────────────────
  w('B66', data.petSalary);     w('C66', data.resSalary);
  w('B67', data.petBenefits);   w('C67', data.resBenefits);
  w('B68', data.petStatePen);   w('C68', data.resStatePen);
  w('B69', data.petPenInc);     w('C69', data.resPenInc);
  w('B70', data.petBankInt);    w('C70', data.resBankInt);
  w('B71', data.petOtherInc);   w('C71', data.resOtherInc);
  w('B72', data.petRental);     w('C72', data.resRental);
  if (data.petRental || data.resRental) w('D73', 'Rental Income');

  // ── Net effect ────────────────────────────────────────────────────────────
  // G79/I79: property equity figures
  w('G79', fmhEquity / 2);
  w('I79', fmhEquity / 2 + prop2Equity);
  // G80, G82, G84 are template formulas (=B80 etc.) — left untouched
  // G86/I86: lump sum
  w('G86', data.petLumpSum);
  w('I86', data.resLumpSum);
  // G90/I90: rental income
  w('G90', data.petRental);
  w('I90', data.resRental);

  return wb.outputAsync();
}

// ─── FILL NEGOTIATION TEMPLATE ────────────────────────────────────────────────

async function fillNegotiationTemplate(templateBuffer, data) {
  const wb = await XlsxPopulate.fromDataAsync(templateBuffer);
  const SHEET = '2. Assets, Debts & Net Effect ';
  const ws = wb.sheet(SHEET);
  if (!ws) throw new Error(`Sheet "${SHEET}" not found`);

  const w  = (ref, val) => { if (val !== null && val !== undefined) ws.cell(ref).value(val); };
  const wd = (ref, iso) => { const d = toDate(iso); if (d) ws.cell(ref).value(d); };
  // wf() writes a formula — Excel will calculate it on open
  const wf = (ref, formula) => ws.cell(ref).formula(formula);

  // ── Names & case number ───────────────────────────────────────────────────
  w('B2', data.petName);
  w('B3', data.resName);
  w('D3', data.caseNumber ? `${data.caseNumber} - ` : '');

  // ── Property address labels ───────────────────────────────────────────────
  w('H6', data.fmhAddress || 'Family Home');
  w('L6', data.prop2Address || 'Property 2');

  // ── Property values ───────────────────────────────────────────────────────
  w('J7',  data.fmhValue);    w('J8',  data.petMortTotal); w('J10', 0);
  w('N7',  data.prop2Value);  w('N8',  data.resMortTotal); w('N10', 0);
  w('R7',  data.prop3Value);  w('R8',  0);                 w('R10', data.prop3Cgt);
  w('V7',  data.prop4Value);  w('V8',  0);                 w('V10', 0);

  // ── Petitioner sole assets — label in H, value in J, up to 14 rows from row 16 ──
  const petAssets = [
    ...data.petBankRows,
    ...data.petIsaRows,
    ...data.petInvestRows,
    ...data.petVehRows,
    ...data.petAddRows,
  ].slice(0, 14);
  petAssets.forEach((asset, i) => {
    w(`H${16 + i}`, asset.label);
    w(`J${16 + i}`, asset.value);
  });

  // ── Respondent sole assets — label in L, value in N, up to 14 rows from row 16 ──
  const resAssets = [
    ...data.resBankRows,
    ...data.resIsaRows,
    ...data.resInvestRows,
    ...data.resVehRows,
    ...data.resAddRows,
    ...data.resBizRows,
  ].slice(0, 14);
  resAssets.forEach((asset, i) => {
    w(`L${16 + i}`, asset.label);
    w(`N${16 + i}`, asset.value);
  });

  // ── Joint other assets ────────────────────────────────────────────────────
  w('J31', data.carParkEach);
  w('N31', data.carParkEach);

  // ── Children assets ───────────────────────────────────────────────────────
  w('R13', data.childBanks);
  w('R14', data.childIsas);
  w('R15', data.childAdd);

  // ── Income now ────────────────────────────────────────────────────────────
  w('B22', data.petSalary);    w('D22', data.resSalary);
  w('B23', data.petBenefits);  w('D23', data.resBenefits);
  w('B24', data.petPenInc);    w('D24', data.resPenInc);
  w('B25', data.petBankInt);   w('D25', data.resBankInt);
  w('B26', data.petOtherInc);  w('D26', data.resOtherInc);

  // ── Net effect — property & capital after (rows 40–50) ───────────────────
  w('B40', data.fmhValue);              w('D40', 0);
  w('B41', 0);                          w('D41', data.prop2Value);
  w('B42', 0);                          w('D42', data.prop3Value - data.prop3Cgt);
  w('B43', 0);                          w('D43', data.prop4Value);

  const petOtherNow = data.petBanks + data.petIsas + data.petInvest + data.petVehs + data.petAddAss;
  const resOtherNow = data.resBanks + data.resBiz  + data.resInvest + data.resVehs + data.resAddAss;
  w('B44', data.petOtherCapAfter || petOtherNow);
  w('D44', data.resOtherCapAfter || resOtherNow);

  wf('B45', 'SUM(B40:B44)');     wf('D45', 'SUM(D40:D44)');
  wf('F40', 'B40+D40');          wf('F41', 'B41+D41');
  wf('F42', 'B42+D42');          wf('F43', 'B43+D43');
  wf('F44', 'B44+D44');          wf('F45', 'B45+D45');
  w('B46', data.petCreditCards + data.petLoans);
  w('D46', data.resTaxLiab + data.resCreditCards + data.resLoans);
  wf('F46', 'B46+D46');
  wf('B47', 'B45-B46');          wf('D47', 'D45-D46');
  wf('F47', 'B47+D47');
  w('B48', data.petPenTotal);    w('D48', data.resPenTotal);
  wf('F48', 'B48+D48');
  w('B49', 0);                   w('D49', 0);
  wf('F49', 'B49+D49');
  wf('C49', 'IF(F49=0,0,B49/F49)');
  wf('E49', 'IF(F49=0,0,D49/F49)');
  wf('B50', 'B47+B48+B49');      wf('D50', 'D47+D48+D49');
  wf('F50', 'B50+D50');

  // ── Petitioner liabilities — label in H, value in J, from row 51 ─────────
  let petLiabRow = 51;
  [...data.petCCRows, ...data.petLoanRows].forEach(liab => {
    w(`H${petLiabRow}`, liab.label);
    w(`J${petLiabRow}`, liab.value);
    petLiabRow++;
  });

  // ── Respondent liabilities — label in L, value in N, from row 51 ─────────
  let resLiabRow = 51;
  [...data.resTaxLiabRows, ...data.resCCRows, ...data.resLoanRows].forEach(liab => {
    w(`L${resLiabRow}`, liab.label);
    w(`N${resLiabRow}`, liab.value);
    resLiabRow++;
  });

  // ── Petitioner pensions — label in H, value in J, rows 62–68 ─────────────
  data.petPensionRows.slice(0, 7).forEach((pen, i) => {
    w(`H${62 + i}`, pen.label);
    w(`J${62 + i}`, pen.value);
  });

  // ── Respondent pensions — label in L, value in N, rows 62–68 ─────────────
  data.resPensionRows.slice(0, 7).forEach((pen, i) => {
    w(`L${62 + i}`, pen.label);
    w(`N${62 + i}`, pen.value);
  });

  // ── Income after ──────────────────────────────────────────────────────────
  w('B58', data.petSalaryAfter);    w('D58', data.resSalaryAfter);
  w('B59', data.petBenAfter);       w('D59', data.resBenAfter);
  w('B60', data.petPenIncAfter);    w('D60', data.resPenIncAfter);
  w('B61', data.petBankIntAfter);   w('D61', data.resBankIntAfter);
  w('B62', data.petOtherIncAfter);  w('D62', data.resOtherIncAfter);

  // ── Dates / background ────────────────────────────────────────────────────
  wd('B77', data.petDobIso);  w('D77', data.petOcc);
  wd('B78', data.resDobIso);  w('D78', data.resOcc);
  wd('F77', data.cohabIso);
  wd('F78', data.marriageIso);
  wd('F79', data.sepIso);
  wd('F80', data.condIso);

  // ── Children — name in A, DOB in B, rows 79–81 ───────────────────────────
  data.children.slice(0, 3).forEach((child, i) => {
    w(`A${79 + i}`, child.name);
    wd(`B${79 + i}`, child.dob);
  });

  if (data.commentary) w('B84', data.commentary);

  return wb.outputAsync();
}

module.exports = { extractData, fillAssistedTemplate, fillNegotiationTemplate };
