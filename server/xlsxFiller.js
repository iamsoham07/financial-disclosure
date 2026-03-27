/**
 * Excel filling logic for both templates
 *
 * Assisted   → Consent-Order---Summary-of-Financial-Disclosure.xlsx
 *              Sheet: "Financial Disclosure"
 *              Input cells: B5,C5 (names), B6,C6 (prop values), B7,C7 (mortgages),
 *              B9,C9 (prop2), B10,C10 (mort2), B16-B26/C16-C26 (other assets),
 *              B32-B38/C32-C38 (liabilities), B44-B52/C44-C52 (pensions),
 *              B59-B67/C59-C67 (income), G72,I72 (net effect property after),
 *              G73,I73 (other assets after), G75,I75 (liabilities after),
 *              G77,I77 (pensions after), G79,I79 (income after),
 *              H3,H4 (DOBs), L3,L4 (cohab/marriage dates), H5,H6,H7 (child DOBs)
 *
 * Negotiation → Financial_Disclosure_and_Net_Effect_Table_Template.xlsx
 *               Sheet: "Assets, Debts & Net Effect "
 *               Input cells: B4,D4 (names), J5,N5,R5,V5 (prop values),
 *               J6,N6,R6,V6 (mortgages), J8,N8,R8,V8 (CGT),
 *               J13-J26 / N13-N26 (other assets), J31/N31 (joint),
 *               R13-R15 (children), J48/N48 (liabilities),
 *               J59-J65/N59-N65 (pensions), B19-B23/D19-D23 (income now),
 *               B40-B50/D40-D50 (net effect after), B55-B59/D55-D59 (income after),
 *               B74,B75 (DOBs), F74-F77 (dates), B78 (case number), B81 (commentary)
 */

const XLSX = require('xlsx');

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

function isoToSerial(iso) {
  if (!iso) return null;
  const dt = new Date(iso);
  if (isNaN(dt)) return null;
  return (dt - new Date(Date.UTC(1899, 11, 30))) / 86400000;
}

/** Write value preserving all existing cell style metadata */
function writeCell(ws, ref, value) {
  if (value === null || value === undefined) return;
  const existing = ws[ref] || {};
  if (typeof value === 'string' && value.startsWith('=')) {
    ws[ref] = { ...existing, t: 'n', f: value.slice(1), v: 0 };
  } else if (typeof value === 'number') {
    ws[ref] = { ...existing, t: 'n', v: value };
  } else {
    ws[ref] = { ...existing, t: 's', v: String(value) };
  }
}

function writeDateCell(ws, ref, iso) {
  const serial = isoToSerial(iso);
  if (serial === null) return;
  const existing = ws[ref] || {};
  ws[ref] = { ...existing, t: 'n', v: serial, z: existing.z || 'dd/mm/yyyy' };
}

// ─── EXTRACT COMMON DATA FROM CONSENT ORDER JSON ─────────────────────────────

function extractData(consentOrder) {
  const d = buildLookup(consentOrder);

  // Names
  const petName = [d['Petitioner.Legal_first_name'] || '', d['Petitioner.Legal_last_name'] || ''].join(' ').trim();
  const resName = [d['Respondent.Legal_first_name'] || '', d['Respondent.Legal_last_name'] || ''].join(' ').trim();

  // Dates
  const petDobIso   = d['Matter.Petitioner_date_of_birth'] || '';
  const resDobIso   = d['Matter.Respondent_date_of_birth'] || '';
  const cohabIso    = d['Petitioner.Background_info_date_cohabiting'] || '';
  const marriageIso = d['Matter.Date_of_marriage'] || '';
  const sepIso      = d['Matter.Date_of_separation'] || '';
  const condIso     = d['Matter.Decree_nisi'] || '';
  const caseNumber  = d['Matter.Case_number'] || '';
  const petOcc      = d['Matter.Petitioner_occupation'] || '';
  const resOcc      = d['Matter.Respondent_occupation'] || '';

  // Children DOBs
  const childrenSection = d['Children.Children_questions'] || [];
  const childDobs = getSectionFieldValues(childrenSection, 'Children.Birth_day_first_child');

  // Properties
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

  // Petitioner assets
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

  // Respondent assets
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

  // Children assets
  const childBanks  = sumSection(d['Children.Bank_accounts'] || [], 'Children.Bank_account_value');
  const childIsas   = sumSection(d['Children.Isas_accounts'] || [], 'Children.Isa_value');
  const childAdd    = sumSection(d['Children.Additional_assets'] || [], 'Children.Additional_assets_value');

  // Income now
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

  // After settlement
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
// Sheet: "Financial Disclosure"
// B5/C5 = names, B6/C6 = FMH value each, B7/C7 = mortgage each
// B9/C9 = prop2, B10/C10 = mortgage2, B16-B26/C16-C26 = other assets
// B32-B38/C32-C38 = liabilities, B44-B52/C44-C52 = pensions
// B59-B67/C59-C67 = income now
// G72/I72 = net effect property, G73/I73 = other assets, G75/I75 = liabilities
// G77/I77 = pensions, G79/I79 = income after
// H3/H4 = DOBs, L3=cohab, L4=marriage, L5=separation, L6=conditional order
// H5/H6/H7 = child DOBs

function fillAssistedTemplate(templateBuffer, data) {
  const wb = XLSX.read(templateBuffer, { type: 'buffer', cellStyles: true, cellFormula: true, cellDates: false });
  const ws = wb.Sheets['Financial Disclosure'];
  if (!ws) throw new Error('Sheet "Financial Disclosure" not found');

  const w  = (ref, val) => writeCell(ws, ref, val);
  const wd = (ref, iso) => writeDateCell(ws, ref, iso);

  // Names (B5=App1, C5=App2 — these drive =B5/=C5 headers throughout)
  w('B5', data.petName);
  w('C5', data.resName);

  // Dates section (right side)
  wd('H3', data.petDobIso);        // App1 DOB
  wd('H4', data.resDobIso);        // App2 DOB
  wd('L3', data.cohabIso);         // Cohabitation date
  wd('L4', data.marriageIso);      // Marriage date
  wd('L5', data.sepIso);           // Separation date
  wd('L6', data.condIso);          // Conditional order date

  // Children DOBs (H5, H6, H7)
  data.childDobs.slice(0, 3).forEach((iso, i) => wd(`H${5 + i}`, iso));

  // Property (Address 1 = FMH, split 50/50)
  w('B6', data.fmhValue / 2);       // App1 share of FMH
  w('C6', data.fmhValue / 2);       // App2 share of FMH
  w('B7', data.petMortTotal / 2);   // App1 mortgage share
  w('C7', data.resMortTotal / 2);   // App2 mortgage share
  w('B8', 0);                        // Costs of sale
  w('C8', 0);

  // Property 2 (Address 2 — Respondent's sole home)
  w('B9', 0);
  w('C9', data.prop2Value);
  w('B10', 0);
  w('C10', data.resMortTotal);
  w('B11', 0);
  w('C11', 0);

  // Other Assets (B16:B26 / C16:C26, summed by template B27/C27)
  w('B16', data.petVehs);           // Vehicles
  w('C16', data.resVehs);
  w('B18', data.petBanks);          // Bank 1 (total banks)
  w('C18', data.resBanks);
  w('B21', data.petIsas);           // ISAs/Savings
  w('C21', data.resIsas);
  w('B22', data.petInvest);         // Shares/Investments
  w('C22', data.resInvest);
  w('B23', 0);                       // Business (petitioner)
  w('C23', data.resBiz);
  w('B24', data.petAddAss);         // Other Assets
  w('C24', data.resAddAss);

  // Liabilities (B32:B38 / C32:C38, summed by template B39/C39)
  w('B32', data.petCreditCards);    // Credit Cards
  w('C32', data.resCreditCards);
  w('B33', data.petLoans);          // Loans
  w('C33', data.resLoans);
  w('B34', 0);                       // Tax Liabilities (petitioner)
  w('C34', data.resTaxLiab);        // Tax Liabilities (respondent)

  // Pensions (B44:B52 / C44:C52, summed by template B53/C53)
  data.petPenVals.slice(0, 9).forEach((val, i) => w(`B${44 + i}`, val));
  data.resPenVals.slice(0, 9).forEach((val, i) => w(`C${44 + i}`, val));

  // Income now (B59:B67 / C59:C67, summed by template B68/C68)
  w('B59', data.petSalary);         // Earned income
  w('C59', data.resSalary);
  w('B60', data.petBenefits);       // State benefits
  w('C60', data.resBenefits);
  w('B61', data.petStatePen);       // State pension
  w('C61', data.resStatePen);
  w('B62', data.petPenInc);         // Other pension income
  w('C62', data.resPenInc);
  w('B63', data.petBankInt);        // Interest
  w('C63', data.resBankInt);
  w('B66', data.petRental);         // Other income (rental)
  w('C66', data.resRental);

  // Net Effect Summary (right side G/I columns)
  // G = App1 after, I = App2 after
  // Property after
  w('G72', data.fmhValue);          // App1 gets FMH
  w('I72', data.prop2Value);        // App2 keeps his home

  // Other assets after — use D81 after values if present, else calculated
  const petOtherNow = data.petBanks + data.petIsas + data.petInvest + data.petVehs + data.petAddAss;
  const resOtherNow = data.resBanks + data.resBiz + data.resInvest + data.resVehs + data.resAddAss;
  w('G73', data.petOtherCapAfter || petOtherNow);
  w('I73', data.resOtherCapAfter || resOtherNow);

  // Liabilities after
  w('G75', data.petCreditCards + data.petLoans);
  w('I75', data.resCreditCards + data.resLoans + data.resTaxLiab);

  // Pensions after (unchanged)
  w('G77', data.petPenTotal);
  w('I77', data.resPenTotal);

  // Income after
  w('G79', data.petSalaryAfter + data.petBenAfter + data.petPenIncAfter + data.petBankIntAfter + data.petOtherIncAfter);
  w('I79', data.resSalaryAfter + data.resBenAfter + data.resPenIncAfter + data.resBankIntAfter + data.resOtherIncAfter);

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
}

// ─── FILL NEGOTIATION TEMPLATE ────────────────────────────────────────────────
// Sheet: "Assets, Debts & Net Effect " (trailing space intentional)

function fillNegotiationTemplate(templateBuffer, data) {
  const wb = XLSX.read(templateBuffer, { type: 'buffer', cellStyles: true, cellFormula: true, cellDates: false });
  const SHEET = 'Assets, Debts & Net Effect ';
  const ws = wb.Sheets[SHEET];
  if (!ws) throw new Error(`Sheet "${SHEET}" not found`);

  const w  = (ref, val) => writeCell(ws, ref, val);
  const wd = (ref, iso) => writeDateCell(ws, ref, iso);

  // Names
  w('B4', data.petName);
  w('D4', data.resName);

  // Dates / background
  wd('B74', data.petDobIso);   w('D74', data.petOcc);
  wd('B75', data.resDobIso);   w('D75', data.resOcc);
  wd('F74', data.cohabIso);
  wd('F75', data.marriageIso);
  wd('F76', data.sepIso);
  wd('F77', data.condIso);
  w('B78', data.caseNumber);

  // Property current
  w('J5', data.fmhValue);    w('J6', data.petMortTotal);  w('J8', 0);
  w('N5', data.prop2Value);  w('N6', data.resMortTotal);  w('N8', 0);
  w('R5', data.prop3Value);  w('R6', 0);                  w('R8', data.prop3Cgt);
  w('V5', data.prop4Value);  w('V6', 0);                  w('V8', 0);

  // Petitioner sole assets (J13:J26 → J27 → J45 → B9)
  w('J13', data.petBanks);
  w('J14', data.petIsas);
  w('J15', data.petInvest);
  w('J16', data.petVehs);
  w('J17', data.petAddAss);

  // Respondent sole assets (N13:N26 → N27 → N45 → D9)
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

  // Liabilities current (J48:J55 → J56 → B11 | N48:N55 → N56 → D11)
  w('J48', data.petCreditCards + data.petLoans);
  w('N48', data.resTaxLiab);
  w('N49', data.resCreditCards + data.resLoans);

  // Pensions current (J59:J65 → J66 → B13 | N59:N65 → N66 → D13)
  data.petPenVals.slice(0, 7).forEach((val, i) => w(`J${59 + i}`, val));
  data.resPenVals.slice(0, 7).forEach((val, i) => w(`N${59 + i}`, val));

  // Income now
  w('B19', data.petSalary);   w('D19', data.resSalary);
  w('B20', data.petBenefits); w('D20', data.resBenefits);
  w('B21', data.petPenInc);   w('D21', data.resPenInc);
  w('B22', data.petBankInt);  w('D22', data.resBankInt);
  w('B23', data.petOtherInc); w('D23', data.resOtherInc);

  // Net effect — property after
  w('B40', data.fmhValue);             w('D40', 0);
  w('B41', 0);                         w('D41', data.prop2Value);
  w('B42', 0);                         w('D42', data.prop3Value - data.prop3Cgt);
  w('B43', 0);                         w('D43', data.prop4Value);

  // Net effect — other capital after
  const petOtherNow = data.petBanks + data.petIsas + data.petInvest + data.petVehs + data.petAddAss;
  const resOtherNow = data.resBanks + data.resBiz + data.resInvest + data.resVehs + data.resAddAss;
  w('B44', data.petOtherCapAfter || petOtherNow);
  w('D44', data.resOtherCapAfter || resOtherNow);

  // Fix template's hardcoded 0s in rows 45-50 with proper formulas
  w('B45', '=SUM(B40:B44)');    w('D45', '=SUM(D40:D44)');
  w('F40', '=B40+D40');         w('F41', '=B41+D41');
  w('F42', '=B42+D42');         w('F43', '=B43+D43');
  w('F44', '=B44+D44');         w('F45', '=B45+D45');
  w('B46', data.petCreditCards + data.petLoans);
  w('D46', data.resTaxLiab + data.resCreditCards + data.resLoans);
  w('F46', '=B46+D46');
  w('B47', '=B45-B46');         w('D47', '=D45-D46');
  w('F47', '=B47+D47');
  w('B48', data.petPenTotal);   w('D48', data.resPenTotal);
  w('F48', '=B48+D48');
  w('B49', 0);                  w('D49', 0);
  w('F49', '=B49+D49');
  w('C49', '=IF(F49=0,0,B49/F49)');
  w('E49', '=IF(F49=0,0,D49/F49)');
  w('B50', '=B47+B48+B49');     w('D50', '=D47+D48+D49');
  w('F50', '=B50+D50');

  // Income after
  w('B55', data.petSalaryAfter);   w('D55', data.resSalaryAfter);
  w('B56', data.petBenAfter);      w('D56', data.resBenAfter);
  w('B57', data.petPenIncAfter);   w('D57', data.resPenIncAfter);
  w('B58', data.petBankIntAfter);  w('D58', data.resBankIntAfter);
  w('B59', data.petOtherIncAfter); w('D59', data.resOtherIncAfter);

  if (data.commentary) w('B81', data.commentary);

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx', cellStyles: true });
}

module.exports = { extractData, fillAssistedTemplate, fillNegotiationTemplate };
