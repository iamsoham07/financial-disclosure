/**
 * Excel filling logic for both templates
 *
 * Uses xlsx-populate which preserves ALL template formatting (colours, borders,
 * merges, row heights, conditional formatting) and keeps existing formulas intact.
 *
 * Assisted   → Sheet: "Financial Disclosure"
 * Negotiation → Sheet: "Assets, Debts & Net Effect " (trailing space)
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

  const childrenSection = d['Children.Children_questions'] || [];
  const childDobs = getSectionFieldValues(childrenSection, 'Children.Birth_day_first_child');

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

  const petBanks    = sumSection(d['Petitioner.Bank_accounts'] || [], 'Petitioner.Bank_account_sole_value');
  const petIsas     = sumSection(d['Petitioner.Isas_account'] || [], 'Petitioner.Isas_value');
  const petInvest   = sumSection(d['Petitioner.Investments'] || [], 'Petitioner.Investments_type_shares_value');
  const petVehs     = sumSection(d['Petitioner.Vehicles_questions'] || [], 'Petitioner.Vehicles_value');
  const petAddAss   = sumSection(d['Petitioner.Additional_assets'] || [], 'Petitioner.Additional_assets_value');
  const petPensions = d['Petitioner.Pensions_private_questions'] || [];
  const petPenVals  = getSectionFieldValues(petPensions, 'Petitioner.Pension_value').map(num);
  const petPenTotal = petPenVals.reduce((a, v) => a + v, 0);
  const petCreditCards = sumSection(d['Petitioner.Credit_cards_info'] || [], 'Petitioner.Credit_card_amount');
  const petLoans    = sumSection(d['Petitioner.Personal_loans'] || [], 'Petitioner.Loan_value');

  const resBanks    = sumSection(d['Respondent.Bank_accounts'] || [], 'Respondent.Bank_account_sole_value');
  const resIsas     = sumSection(d['Respondent.Isas_account'] || [], 'Respondent.Isas_value');
  const resInvest   = sumSection(d['Respondent.Investments'] || [], 'Respondent.Investments_type_shares_value');
  const resVehs     = sumSection(d['Respondent.Vehicles_questions'] || [], 'Respondent.Vehicles_value');
  const resAddAss   = sumSection(d['Respondent.Additional_assets_personal'] || [], 'Respondent.Additional_assets_value');
  const resBiz      = sumSection(d['Respondent.Business_assets_questions'] || [], 'Respondent.Business_assets_value');
  const resPensions = d['Respondent.Pensions_private_questions'] || [];
  const resPenVals  = getSectionFieldValues(resPensions, 'Respondent.Pension_value').map(num);
  const resPenTotal = resPenVals.reduce((a, v) => a + v, 0);
  const resCreditCards  = sumSection(d['Respondent.Credit_cards_info'] || [], 'Respondent.Credit_card_amount');
  const resLoans    = sumSection(d['Respondent.Personal_loans'] || [], 'Respondent.Loan_value');
  const resTaxLiab  = sumSection(d['Respondent.Tax_liability'] || [], 'Respondent.Tax_liability_total_current_value');

  const childBanks  = sumSection(d['Children.Bank_accounts'] || [], 'Children.Bank_account_value');
  const childIsas   = sumSection(d['Children.Isas_accounts'] || [], 'Children.Isa_value');
  const childAdd    = sumSection(d['Children.Additional_assets'] || [], 'Children.Additional_assets_value');

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

  return {
    petName, resName, petOcc, resOcc, caseNumber,
    petDobIso, resDobIso, cohabIso, marriageIso, sepIso, condIso,
    childDobs,
    fmhValue, prop2Value, prop3Value, prop4Value, prop3Cgt,
    petMortTotal, resMortTotal, carParkEach,
    petBanks, petIsas, petInvest, petVehs, petAddAss,
    petPenVals, petPenTotal, petCreditCards, petLoans,
    resBanks, resIsas, resInvest, resVehs, resAddAss, resBiz,
    resPenVals, resPenTotal, resCreditCards, resLoans, resTaxLiab,
    childBanks, childIsas, childAdd,
    petSalary, petBenefits, petStatePen, petPenInc, petBankInt, petOtherInc, petRental,
    resSalary, resBenefits, resStatePen, resPenInc, resBankInt, resOtherInc, resRental,
    petOtherCapAfter, resOtherCapAfter,
    petSalaryAfter, petBenAfter, petPenIncAfter, petBankIntAfter, petOtherIncAfter,
    resSalaryAfter, resBenAfter, resPenIncAfter, resBankIntAfter, resOtherIncAfter,
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

  // Names
  w('B5', data.petName);
  w('C5', data.resName);

  // Dates
  wd('H3', data.petDobIso);
  wd('H4', data.resDobIso);
  wd('L3', data.cohabIso);
  wd('L4', data.marriageIso);
  wd('L5', data.sepIso);
  wd('L6', data.condIso);
  data.childDobs.slice(0, 3).forEach((iso, i) => wd(`H${5 + i}`, iso));

  // Property (FMH split 50/50)
  w('B6', data.fmhValue / 2);
  w('C6', data.fmhValue / 2);
  w('B7', data.petMortTotal / 2);
  w('C7', data.resMortTotal / 2);
  w('B8', 0);
  w('C8', 0);

  // Property 2
  w('B9',  0);
  w('C9',  data.prop2Value);
  w('B10', 0);
  w('C10', data.resMortTotal);
  w('B11', 0);
  w('C11', 0);

  // Other assets
  w('B16', data.petVehs);
  w('C16', data.resVehs);
  w('B18', data.petBanks);
  w('C18', data.resBanks);
  w('B21', data.petIsas);
  w('C21', data.resIsas);
  w('B22', data.petInvest);
  w('C22', data.resInvest);
  w('B23', 0);
  w('C23', data.resBiz);
  w('B24', data.petAddAss);
  w('C24', data.resAddAss);

  // Liabilities
  w('B32', data.petCreditCards);
  w('C32', data.resCreditCards);
  w('B33', data.petLoans);
  w('C33', data.resLoans);
  w('B34', 0);
  w('C34', data.resTaxLiab);

  // Pensions
  data.petPenVals.slice(0, 9).forEach((val, i) => w(`B${44 + i}`, val));
  data.resPenVals.slice(0, 9).forEach((val, i) => w(`C${44 + i}`, val));

  // Income now
  w('B59', data.petSalary);
  w('C59', data.resSalary);
  w('B60', data.petBenefits);
  w('C60', data.resBenefits);
  w('B61', data.petStatePen);
  w('C61', data.resStatePen);
  w('B62', data.petPenInc);
  w('C62', data.resPenInc);
  w('B63', data.petBankInt);
  w('C63', data.resBankInt);
  w('B66', data.petRental);
  w('C66', data.resRental);

  // Net effect summary
  w('G72', data.fmhValue);
  w('I72', data.prop2Value);

  const petOtherNow = data.petBanks + data.petIsas + data.petInvest + data.petVehs + data.petAddAss;
  const resOtherNow = data.resBanks + data.resBiz  + data.resInvest + data.resVehs + data.resAddAss;
  w('G73', data.petOtherCapAfter || petOtherNow);
  w('I73', data.resOtherCapAfter || resOtherNow);

  w('G75', data.petCreditCards + data.petLoans);
  w('I75', data.resCreditCards + data.resLoans + data.resTaxLiab);
  w('G77', data.petPenTotal);
  w('I77', data.resPenTotal);
  w('G79', data.petSalaryAfter + data.petBenAfter + data.petPenIncAfter + data.petBankIntAfter + data.petOtherIncAfter);
  w('I79', data.resSalaryAfter + data.resBenAfter + data.resPenIncAfter + data.resBankIntAfter + data.resOtherIncAfter);

  return wb.outputAsync();
}

// ─── FILL NEGOTIATION TEMPLATE ────────────────────────────────────────────────

async function fillNegotiationTemplate(templateBuffer, data) {
  const wb = await XlsxPopulate.fromDataAsync(templateBuffer);
  const SHEET = 'Assets, Debts & Net Effect ';
  const ws = wb.sheet(SHEET);
  if (!ws) throw new Error(`Sheet "${SHEET}" not found`);

  const w  = (ref, val) => { if (val !== null && val !== undefined) ws.cell(ref).value(val); };
  const wd = (ref, iso) => { const d = toDate(iso); if (d) ws.cell(ref).value(d); };
  // wf() writes a formula — Excel will calculate it on open
  const wf = (ref, formula) => ws.cell(ref).formula(formula);

  // Names
  w('B4', data.petName);
  w('D4', data.resName);

  // Dates / background
  wd('B74', data.petDobIso);  w('D74', data.petOcc);
  wd('B75', data.resDobIso);  w('D75', data.resOcc);
  wd('F74', data.cohabIso);
  wd('F75', data.marriageIso);
  wd('F76', data.sepIso);
  wd('F77', data.condIso);
  w('B78', data.caseNumber);

  // Property current
  w('J5', data.fmhValue);   w('J6', data.petMortTotal);  w('J8', 0);
  w('N5', data.prop2Value); w('N6', data.resMortTotal);   w('N8', 0);
  w('R5', data.prop3Value); w('R6', 0);                   w('R8', data.prop3Cgt);
  w('V5', data.prop4Value); w('V6', 0);                   w('V8', 0);

  // Petitioner sole assets
  w('J13', data.petBanks);
  w('J14', data.petIsas);
  w('J15', data.petInvest);
  w('J16', data.petVehs);
  w('J17', data.petAddAss);

  // Respondent sole assets
  w('N13', data.resBanks);
  w('N14', data.resBiz);
  w('N15', data.resInvest);
  w('N16', data.resAddAss);
  w('N17', data.resVehs);

  // Joint other assets
  w('J31', data.carParkEach);
  w('N31', data.carParkEach);

  // Children assets
  w('R13', data.childBanks);
  w('R14', data.childIsas);
  w('R15', data.childAdd);

  // Liabilities current
  w('J48', data.petCreditCards + data.petLoans);
  w('N48', data.resTaxLiab);
  w('N49', data.resCreditCards + data.resLoans);

  // Pensions current
  data.petPenVals.slice(0, 7).forEach((val, i) => w(`J${59 + i}`, val));
  data.resPenVals.slice(0, 7).forEach((val, i) => w(`N${59 + i}`, val));

  // Income now
  w('B19', data.petSalary);    w('D19', data.resSalary);
  w('B20', data.petBenefits);  w('D20', data.resBenefits);
  w('B21', data.petPenInc);    w('D21', data.resPenInc);
  w('B22', data.petBankInt);   w('D22', data.resBankInt);
  w('B23', data.petOtherInc);  w('D23', data.resOtherInc);

  // Net effect — property after
  w('B40', data.fmhValue);              w('D40', 0);
  w('B41', 0);                          w('D41', data.prop2Value);
  w('B42', 0);                          w('D42', data.prop3Value - data.prop3Cgt);
  w('B43', 0);                          w('D43', data.prop4Value);

  // Net effect — other capital after
  const petOtherNow = data.petBanks + data.petIsas + data.petInvest + data.petVehs + data.petAddAss;
  const resOtherNow = data.resBanks + data.resBiz  + data.resInvest + data.resVehs + data.resAddAss;
  w('B44', data.petOtherCapAfter || petOtherNow);
  w('D44', data.resOtherCapAfter || resOtherNow);

  // Calculation formulas — Excel will evaluate these on open
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

  // Income after
  w('B55', data.petSalaryAfter);    w('D55', data.resSalaryAfter);
  w('B56', data.petBenAfter);       w('D56', data.resBenAfter);
  w('B57', data.petPenIncAfter);    w('D57', data.resPenIncAfter);
  w('B58', data.petBankIntAfter);   w('D58', data.resBankIntAfter);
  w('B59', data.petOtherIncAfter);  w('D59', data.resOtherIncAfter);

  if (data.commentary) w('B81', data.commentary);

  return wb.outputAsync();
}

module.exports = { extractData, fillAssistedTemplate, fillNegotiationTemplate };
