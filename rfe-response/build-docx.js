const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Header, Footer,
  PageNumber, Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle, convertInchesToTwip,
} = require('docx');

const FONT = 'Times New Roman';
const SZ = 22;            // 11pt half-points
const SZ_BODY = 23;       // 11.5pt

// --- inline markup: **bold**, __italic__ -----------------------------------
function rt(text, base = {}) {
  const out = [];
  const re = /(\*\*[^*]+\*\*|__[^_]+__)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), font: FONT, size: SZ_BODY, ...base }));
    const tok = m[0];
    if (tok.startsWith('**')) out.push(new TextRun({ text: tok.slice(2, -2), font: FONT, size: SZ_BODY, bold: true, ...base }));
    else out.push(new TextRun({ text: tok.slice(2, -2), font: FONT, size: SZ_BODY, italics: true, ...base }));
    last = re.lastIndex;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), font: FONT, size: SZ_BODY, ...base }));
  return out;
}

const P = (text, opts = {}) => new Paragraph({
  children: rt(text, opts.base || {}),
  spacing: { after: opts.after !== undefined ? opts.after : 160, line: 276 },
  alignment: opts.align,
  indent: opts.indent,
  ...(opts.extra || {}),
});

const BLANK = (after = 120) => new Paragraph({ text: '', spacing: { after } });

const H = (text, level) => new Paragraph({
  children: [new TextRun({ text, font: FONT, size: level === 1 ? 24 : SZ_BODY, bold: true })],
  spacing: { before: 320, after: 160 },
  keepNext: true,
});

const QUOTE = (text) => new Paragraph({
  children: rt(text, { italics: true }),
  indent: { left: convertInchesToTwip(0.5), right: convertInchesToTwip(0.5) },
  spacing: { before: 120, after: 200, line: 260 },
});

const BULLET = (text, level = 0) => new Paragraph({
  children: rt(text),
  bullet: { level },
  spacing: { after: 120, line: 276 },
});

const NUM = (text, ref) => new Paragraph({
  children: rt(text),
  numbering: { reference: ref, level: 0 },
  spacing: { after: 120, line: 276 },
});

// --- exhibit table ---------------------------------------------------------
const COLS = [convertInchesToTwip(0.55), convertInchesToTwip(4.15), convertInchesToTwip(1.8)];
const cell = (text, i, opts = {}) => new TableCell({
  width: { size: COLS[i], type: WidthType.DXA },
  shading: opts.head ? { type: ShadingType.CLEAR, fill: 'EFEFEF' } : undefined,
  margins: { top: 80, bottom: 80, left: 100, right: 100 },
  children: [new Paragraph({
    children: [new TextRun({ text, font: FONT, size: 21, bold: !!opts.head })],
    spacing: { after: 0, line: 240 },
  })],
});
const exRow = (a, b, c, head = false) => new TableRow({
  tableHeader: head,
  children: [cell(a, 0, { head }), cell(b, 1, { head }), cell(c, 2, { head })],
});

const EXHIBITS = [
  ['A', 'Board of Directors resolution deferring cash compensation, adopted [DATE], certified by the Secretary', '[PERIOD]'],
  ['B', 'Declaration of [DIRECTOR / OFFICER NAME], [TITLE], Octet Inc., attesting to the Beneficiary’s continuous full-time service as CEO', 'Sept 10 – Nov 10, 2025'],
  ['C', 'Minutes of Board meetings held [DATES]', 'Within RFE period'],
  ['D', 'Internal leadership meeting records / calendar records', 'Sept–Nov 2025'],
  ['E', 'Executed financing documents ([TERM SHEET / SAFE / STOCK PURCHASE AGREEMENT]) signed by the Beneficiary as CEO', '[DATES]'],
  ['F', 'Investor correspondence and diligence materials authored by the Beneficiary', 'Sept–Nov 2025'],
  ['G', 'Commercial agreements executed by the Beneficiary as Octet’s authorized signatory', '[DATES]'],
  ['H', 'Third-party materials identifying the Beneficiary as Octet’s CEO ([PRESS / EVENT / PARTNER MATERIALS])', '[DATES]'],
  ['I', 'Product and technical work product and release records', 'Sept–Nov 2025'],
  ['J', 'Octet Inc. financial statements showing accrued compensation liability owed to the Beneficiary', 'Q3–Q4 2025 [and current]'],
  ['K', 'Payroll-provider records ([PROVIDER]) showing active employee enrollment and payroll register', 'Sept 10 – Nov 10, 2025'],
  ['L', 'Health insurance enrollment confirmation and carrier invoices showing employer-paid premiums', 'Sept, Oct, Nov 2025'],
  ['M', 'Payroll / earnings statements for the requested period [IF ANY EXIST]', 'Sept 10 – Nov 10, 2025'],
  ['N', 'Form W-2 (2025) and Forms 941 (Q3, Q4 2025)', '2025'],
  ['O', 'Evidence of payment of deferred compensation [IF ANY HAS BEEN PAID]', '[DATE]'],
  ['P', 'Employment agreement and any deferral addendum signed by the Beneficiary', '[DATE]'],
  ['Q', 'Declaration of Amal Devani', '—'],
];

const body = [];
const add = (...x) => body.push(...x);

// Letterhead
add(new Paragraph({
  children: [new TextRun({ text: 'OCTET INC.', font: FONT, size: 30, bold: true, characterSpacing: 20 })],
  spacing: { after: 40 },
}));
add(new Paragraph({
  children: [new TextRun({ text: '1115 Broadway, Floor 11  ·  New York, NY 10010', font: FONT, size: 20 })],
  spacing: { after: 80 },
  border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000', space: 6 } },
}));
add(BLANK(240));

add(P('September [__], 2026', { after: 280 }));

add(P('U.S. Citizenship and Immigration Services', { after: 0 }));
add(P('Attn: RFE/NOIT/NOIR/NOID Response', { after: 0 }));
add(P('6046 N. Belt Line Rd., Suite 111', { after: 0 }));
add(P('Irving, TX 75038-0011', { after: 280 }));

add(P('**RE:  Response to Request for Evidence**', { after: 60 }));
add(P('**Petitioner:**  Octet Inc.', { after: 0, indent: { left: convertInchesToTwip(0.4) } }));
add(P('**Beneficiary:**  Amal Devani', { after: 0, indent: { left: convertInchesToTwip(0.4) } }));
add(P('**Form:**  I-129, Petition for a Nonimmigrant Worker (O-1A)', { after: 0, indent: { left: convertInchesToTwip(0.4) } }));
add(P('**Receipt No.:**  IOE8795688197', { after: 0, indent: { left: convertInchesToTwip(0.4) } }));
add(P('**RFE Issued:**  August 28, 2026        **Response Due:**  September 30, 2026', { after: 280, indent: { left: convertInchesToTwip(0.4) } }));

add(P('Dear Officer:', { after: 200 }));

add(P('Octet Inc. (“Octet” or “the Petitioner”) submits this response to the Request for Evidence dated August 28, 2026, concerning its Form I-129 petition seeking to extend the O-1A classification of Amal Devani (the “Beneficiary”) in the position of Founder and Chief Executive Officer.'));
add(P('The RFE raises a single issue. Under the heading **“Maintaining Status,”** USCIS requests:'));
add(QUOTE('“Submit copies of the beneficiary’s pay records (leave and earnings statements, pay stubs, etc.) from September 10, 2025, to November 10, 2025, to show that the beneficiary has maintained nonimmigrant status.”'));
add(P('This response addresses that request directly. As the RFE expressly permits, the Petitioner “submit[s] one, some, or all of these items” and additionally “submit[s] other evidence to satisfy the request.”'));

// I
add(H('I.  Summary of Response', 1));
add(P('Throughout the period September 10, 2025 through November 10, 2025 — and continuously before and after it — Mr. Devani **remained employed full-time by Octet Inc., the petitioner through which his O-1 status was obtained, in the identical capacity described in the approved petition: Founder and Chief Executive Officer.** He performed the executive duties set out in that petition without interruption.'));
add(P('During that same period, and pursuant to a **resolution of the Board of Directors adopted [DATE OF BOARD RESOLUTION]**, Octet temporarily **deferred the cash disbursement** of Mr. Devani’s salary in order to preserve runway during an active financing process. The deferral changed the **timing of payment only**. It did not reduce, forgive, waive, or extinguish the compensation owed. **The full amount continues to be owed to Mr. Devani, has been carried on Octet’s books as an accrued liability since it was incurred, and remains payable in full** under the terms of the resolution. [IF PARTIALLY OR FULLY REPAID: As of [DATE], Octet has paid $[AMOUNT] of the deferred balance; the remaining $[AMOUNT] is scheduled for payment upon [TRIGGER].]'));
add(P('Mr. Devani also continued to receive **non-cash remuneration** from Octet throughout the period, including **employer-sponsored health insurance coverage with premiums paid by the company** — an economic benefit conferred by the employer in consideration of his continued service.'));
add(P('The Petitioner respectfully submits that these facts establish maintenance of status. As set out in Part III below, valid O-1 status turns on **whether the beneficiary continued to work for the petitioning employer in the capacity specified in the approved petition** — not on the timing or amount of cash compensation. USCIS’s own published guidance confirms that for the O classification “**no particular wage structure is required.**”'));

// II
add(H('II.  Statement of Facts', 1));
add(H('A.  The approved petition and the position', 2));
add(P('Octet Inc. is a [BRIEF DESCRIPTION OF THE COMPANY — sector, product, stage], incorporated in [STATE] and headquartered at 1115 Broadway, Floor 11, New York, NY 10010. Mr. Devani’s O-1A classification was approved under Receipt No. [PRIOR RECEIPT NUMBER], valid from [DATE] to [DATE], for the position of Founder and Chief Executive Officer. The instant petition was filed on December 10, 2025 to extend that classification in the same position, with the same employer, for continuation of the same business activity.'));

add(H('B.  Continuous full-time employment in the same capacity', 2));
add(P('Mr. Devani has served as Octet’s Founder and Chief Executive Officer since [DATE], and did so without interruption from September 10, 2025 through November 10, 2025. He worked full-time, [HOURS] hours per week, at the company’s New York offices and remotely on company business. His duties during the RFE period were the duties described in the approved petition, including:'));
add(NUM('**Executive leadership and management.** Setting company strategy; directing [NUMBER] employees and [NUMBER] contractors; chairing [WEEKLY / BIWEEKLY] leadership meetings; and reporting to the Board of Directors. (Exhibits [C], [D] — board minutes and internal leadership meeting records dated within the RFE period.)', 'duties'));
add(NUM('**Capital formation.** Leading Octet’s [SEED / SERIES A / BRIDGE] financing — preparing materials, running the investor process, negotiating terms, and executing definitive documents. During the RFE period specifically, Mr. Devani [DESCRIBE: e.g., “conducted investor meetings with [N] firms, negotiated and signed a term sheet dated [DATE] with [INVESTOR], and closed $[AMOUNT] in financing on [DATE]”]. (Exhibits [E], [F] — executed financing documents and dated investor correspondence.)', 'duties'));
add(NUM('**Representing the company externally.** Acting as Octet’s authorized signatory and public representative — executing [CUSTOMER / VENDOR / PARTNERSHIP] agreements on the company’s behalf, and representing Octet at [CONFERENCE / PRESS / INDUSTRY EVENT] on [DATE]. (Exhibits [G], [H] — countersigned agreements bearing his signature as CEO within the RFE period, and third-party materials identifying him as Octet’s CEO.)', 'duties'));
add(NUM('**Product and technical direction.** [DESCRIBE, IF APPLICABLE — releases shipped, roadmap decisions, technical leadership during the period.] (Exhibit [I].)', 'duties'));
add(P('None of this work stopped, paused, or diminished at any point in the RFE period. Mr. Devani took no leave of absence and held no other employment.'));

add(H('C.  The Board resolution deferring cash compensation', 2));
add(P('On [DATE], Octet’s Board of Directors adopted a written resolution (the “Deferral Resolution”) addressing executive cash compensation in light of the company’s cash position during an active financing process. In relevant part, the Deferral Resolution provides that:'));
add(BULLET('Mr. Devani’s employment as Chief Executive Officer **continues in full force**, with no change to his title, duties, authority, reporting line, or full-time commitment;'));
add(BULLET('his annual base salary of **$[AMOUNT]** remains **fully earned and owed** for services rendered;'));
add(BULLET('**cash disbursement** of that salary is **deferred**, beginning [START DATE] and continuing until [END DATE OR TRIGGERING EVENT — e.g., “the closing of the Company’s next equity financing of not less than $[AMOUNT]”];'));
add(BULLET('the deferred amounts **accrue as a liability of the Company** payable to Mr. Devani [WITH INTEREST AT [RATE] / WITHOUT INTEREST]; and'));
add(BULLET('nothing in the resolution reduces, waives, or forgives any portion of the compensation owed.'));
add(P('(Exhibit [A] — certified copy of the Deferral Resolution, with signature page.)'));

add(H('D.  The deferred compensation is a real, recorded, outstanding obligation', 2));
add(P('The deferred amounts are not a fiction of paperwork. They appear as **accrued compensation on Octet’s balance sheet** for each period since the deferral began, and in the company’s payroll records maintained by [PAYROLL PROVIDER], with which Mr. Devani remained an actively enrolled W-2 employee throughout the RFE period. As of [DATE], the outstanding accrued balance owed to Mr. Devani is **$[AMOUNT]**.'));
add(P('(Exhibits [J], [K] — financial statements showing the accrued compensation liability for the periods covering September–November 2025, and payroll-provider records for the same period showing Mr. Devani’s active employee enrollment.)'));

add(H('E.  Non-cash remuneration during the RFE period', 2));
add(P('Octet provided, and paid the premiums for, **employer-sponsored health insurance** covering Mr. Devani continuously through the RFE period, under plan [PLAN / CARRIER NAME], at a company cost of $[AMOUNT] per month. (Exhibit [L] — enrollment confirmation and carrier invoices showing premiums paid by Octet for September, October, and November 2025.) [ADD ANY OTHER BENEFITS ACTUALLY PROVIDED: equity vesting under his [DATE] stock agreement, expense reimbursement, device or technology allowance, professional liability coverage.]'));

add(H('F.  The specific records requested', 2));
add(P('__[Select the option below that is accurate and delete the others.]__'));
add(BULLET('__[If pay statements exist for the window:]__ Enclosed at Exhibit [M] are Mr. Devani’s payroll statements for each pay period from September 10, 2025 through November 10, 2025. Consistent with the Deferral Resolution, these statements reflect [$0 / reduced] cash disbursement while showing his continued active employment and the accrual of compensation owed.'));
add(BULLET('__[If no pay statements were generated:]__ Because no cash wages were disbursed to Mr. Devani during September 10, 2025 through November 10, 2025 pursuant to the Deferral Resolution, no earnings statements were generated for those pay periods. In place of statements that do not exist, the Petitioner submits the payroll register and employee-status records from [PAYROLL PROVIDER] for that period (Exhibit [K]), the accrued-compensation entries from the company’s books (Exhibit [J]), and the contemporaneous work evidence described above.'));
add(BULLET('__[If some cash was paid:]__ Enclosed at Exhibit [M] are pay statements reflecting the $[AMOUNT] in cash compensation actually disbursed to Mr. Devani during the RFE period, the balance having been deferred under the Deferral Resolution.'));
add(P('The Petitioner also encloses **Form W-2 for tax year 2025** and **IRS Form 941 quarterly returns for Q3 and Q4 2025** (Exhibit [N]), and [IF APPLICABLE] **evidence of subsequent payment of deferred amounts** (Exhibit [O]).'));

// III
add(H('III.  Legal Argument: Deferred Cash Compensation Does Not Affect Maintenance of O-1 Status', 1));
add(H('A.  The elements of maintaining O-1 status do not include receipt of cash wages', 2));
add(P('An O-1 nonimmigrant maintains status by remaining within the terms and conditions of the classification in which he was admitted. See INA § 237(a)(1)(C)(i), 8 U.S.C. § 1227(a)(1)(C)(i). For an O-1 beneficiary those terms are:'));
add(NUM('**Employment only by the petitioning employer.** “An alien in this status may be employed only by the petitioner through whom the status was obtained.” 8 C.F.R. § 274a.12(b)(13).', 'elements'));
add(NUM('**Employment only in the approved capacity and activity.** 8 C.F.R. § 214.2(o)(8)(iii)(A)(1) makes revocation appropriate where “[t]he beneficiary is no longer employed by the petitioner **in the capacity specified in the petition**.”', 'elements'));
add(NUM('**Employment only during the petition validity period.** “The beneficiary may only engage in employment during the validity period of the petition.” 8 C.F.R. § 214.2(o)(10).', 'elements'));
add(NUM('**No unauthorized employment.** “A nonimmigrant who is permitted to engage in employment may engage only in such employment as has been authorized.” 8 C.F.R. § 214.1(e).', 'elements'));
add(P('Mr. Devani satisfied every one of these conditions during September 10 – November 10, 2025. He worked for Octet Inc. and no one else; he worked as Founder and Chief Executive Officer, the capacity specified in the approved petition; he worked within the petition’s validity period; and he engaged in no other employment.'));
add(P('**Conspicuously absent from these conditions is any requirement that the beneficiary receive cash compensation, in any amount or on any schedule.** Neither INA § 101(a)(15)(O)(i), nor 8 C.F.R. § 214.2(o), nor 8 C.F.R. § 214.1 conditions the maintenance of O-1 status on payment. The statutory definition asks whether the beneficiary “seeks to enter the United States **to continue work** in the area of extraordinary ability,” INA § 101(a)(15)(O)(i), and the regulation describes an O-1 beneficiary as one authorized “**to perform services** relating to an event or events,” 8 C.F.R. § 214.2(o)(1)(i). The touchstone is the performance of services.'));

add(H('B.  USCIS has expressly stated that no wage structure is required for the O classification', 2));
add(P('USCIS’s published guidance, __O Nonimmigrant Classifications: Question and Answers__, states without qualification:'));
add(QUOTE('“The regulations do not contain a prevailing wage requirement. Furthermore, no particular wage structure is required.”'));
add(P('The same guidance confirms that the compensation term is whatever the parties agreed to: “A detailed description of the wage offered or fee structure and that the wage offered/fee structure was agreed upon may satisfy this requirement.” Consistent with that framework, 8 C.F.R. § 214.2(o)(2)(ii)(B) requires only “[c]opies of any written contracts between the petitioner and the alien beneficiary or, if there is no written contract, a summary of the terms of the oral agreement under which the alien will be employed,” and the USCIS Policy Manual explains that the summary need only set out “[t]he terms offered by the petitioner (employer)” and “[t]he terms accepted by the beneficiary (employee).” 2 USCIS-PM M.7.'));
add(P('Here, the terms of employment — including the timing of payment — were set by the Board of Directors and accepted by Mr. Devani, and are documented in the Deferral Resolution at Exhibit [A]. The arrangement __is__ the agreed term, not a departure from it.'));

add(H('C.  The contrast with the H-1B program confirms the point', 2));
add(P('Where Congress and the Department of Labor intended to guarantee payment to a nonimmigrant worker, they said so explicitly. The H-1B program requires the employer to pay the required wage even when the worker is in “nonproductive status,” including “because of lack of assigned work … or any other reason.” 20 C.F.R. § 655.731(c)(7)(i). That obligation flows from the H-1B labor condition application regime under INA § 212(n).'));
add(P('**The O classification has no labor condition application, no prevailing wage requirement, and no counterpart to § 655.731(c)(7).** The three references to compensation anywhere in 8 C.F.R. § 214.2(o) confirm this: a disclosure requirement for agent-filed petitions (§ 214.2(o)(2)(iv)(E)(1)); an **optional** evidentiary criterion regarding high salary, one of eight (§ 214.2(o)(3)(iii)(B)(8)); and a labor-dispute provision (§ 214.2(o)(14)). None imposes a duty to disburse cash on any schedule.'));

add(H('D.  Nothing about the deferral removed Mr. Devani from the capacity specified in the petition', 2));
add(P('The one compensation-adjacent risk to O-1 status would be a change so fundamental that the beneficiary is “no longer employed by the petitioner in the capacity specified in the petition,” 8 C.F.R. § 214.2(o)(8)(iii)(A)(1), or a “material change[] in the terms and conditions of employment” requiring an amended petition, 8 C.F.R. § 214.2(o)(2)(iv)(D). Neither occurred:'));
add(BULLET('**His capacity did not change.** Title, duties, authority, reporting relationship, worksite, and full-time commitment were identical before, during, and after the deferral. The evidence at Exhibits [C]–[I] shows him performing the CEO function continuously through the RFE period — chairing the board, leading a financing to a signed [TERM SHEET / CLOSING], and executing agreements as the company’s authorized officer.'));
add(BULLET('**His compensation was not reduced.** The salary was earned in full and remains owed in full. What changed was the date of disbursement. The obligation is recorded on the company’s books as a liability (Exhibit [J]) and remains outstanding [/ has since been paid in the amount of $[AMOUNT], Exhibit [O]].'));
add(BULLET('**He continued to receive remuneration in kind.** Employer-paid health insurance premiums (Exhibit [L]) were conferred on him as an employee throughout the period, which is itself inconsistent with any suggestion that the employment relationship lapsed.'));
add(BULLET('**The employment relationship never ceased.** He remained an actively enrolled W-2 employee of Octet on the company’s payroll system throughout (Exhibit [K]) and was reported as such on the company’s employment tax filings (Exhibit [N]).'));

add(H('E.  Deferring founder cash compensation during a financing is ordinary and expected in this industry', 2));
add(P('Octet is an early-stage venture-backed company. Preserving cash runway by deferring the chief executive’s salary — while the chief executive continues to work full-time and the obligation remains on the books — is a routine and prudent exercise of the Board’s fiduciary duty, not an abandonment of employment. Indeed, the deferral existed **because** Mr. Devani was working: the compensation accrued precisely because he continued to render services as CEO. A reading that treated deferral as a status violation would penalize the beneficiary for the very work the petition describes, and would find no support in the statute, the regulations, or USCIS guidance.'));

// IV
add(H('IV.  Response to the “Prior Approvals” Discussion', 1));
add(P('The RFE notes that USCIS will not defer to a prior approval where there has been “a material change in circumstances or eligibility requirements.” The Petitioner respectfully submits that there has been **no material change** in Mr. Devani’s employment or eligibility: he holds the same position, with the same employer, performing the same duties, in furtherance of the same business activity described in the approved petition. The record submitted with the petition and supplemented here independently establishes his eligibility under 8 C.F.R. § 214.2(o)(3)(iii), without reliance on the prior approval, and the Petitioner acknowledges and accepts its burden of proof under INA § 291.'));
add(P('[IF THE PETITION RELIED ON THE HIGH-SALARY CRITERION AT § 214.2(o)(3)(iii)(B)(8), ADD: Moreover, the criteria satisfied in this petition are [LIST], each independently established in the record. The beneficiary’s total remuneration — including the full accrued salary owed, employer-paid benefits, and his equity in Octet valued at $[AMOUNT] per the company’s [DATE] 409A valuation — further supports eligibility. See 8 C.F.R. § 214.2(o)(3)(iii)(C) (comparable evidence).]'));

// V
add(H('V.  Index of Enclosed Evidence', 1));
add(new Table({
  columnWidths: COLS,
  width: { size: COLS.reduce((a, b) => a + b, 0), type: WidthType.DXA },
  rows: [exRow('Ex.', 'Document', 'Period covered', true), ...EXHIBITS.map(r => exRow(r[0], r[1], r[2]))],
}));
add(BLANK(200));

// VI
add(H('VI.  Conclusion', 1));
add(P('The record establishes that from September 10, 2025 through November 10, 2025, Mr. Devani was employed full-time by Octet Inc. — the petitioner through which his O-1 status was obtained — in the identical capacity described in the approved petition, performing the executive duties of Founder and Chief Executive Officer, receiving employer-paid benefits, and earning compensation that Octet recorded as an obligation and remains bound to pay in full. The Board’s decision to defer the timing of cash disbursement changed neither the employment relationship nor the compensation owed, and no provision of the Act or the regulations conditions O-1 status on the receipt of cash wages on a particular schedule.'));
add(P('The Petitioner respectfully requests that USCIS find that the Beneficiary maintained valid O-1 nonimmigrant status and approve the petition. Should any additional documentation assist in the adjudication, the Petitioner will provide it promptly upon request.'));
add(BLANK(160));
add(P('Respectfully submitted,', { after: 700 }));
add(new Paragraph({
  children: [new TextRun({ text: '', font: FONT, size: SZ_BODY })],
  border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '000000', space: 2 } },
  spacing: { after: 80 },
  indent: { right: convertInchesToTwip(3.6) },
}));
add(P('[NAME], [TITLE]', { after: 0 }));
add(P('Octet Inc.', { after: 0 }));
add(P('[PHONE]  ·  [EMAIL]', { after: 260 }));
add(P('**Enclosures:** RFE coversheet; original RFE notice; Exhibits A–Q'));

const doc = new Document({
  creator: 'Octet Inc.',
  title: 'RFE Response — Octet Inc. / Amal Devani (IOE8795688197)',
  description: 'Response to Request for Evidence dated August 28, 2026',
  numbering: {
    config: [
      { reference: 'duties', levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) } } } }] },
      { reference: 'elements', levels: [{ level: 0, format: 'decimal', text: '%1.', alignment: AlignmentType.START, style: { paragraph: { indent: { left: convertInchesToTwip(0.5), hanging: convertInchesToTwip(0.25) } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: convertInchesToTwip(1), bottom: convertInchesToTwip(1), left: convertInchesToTwip(1), right: convertInchesToTwip(1) },
      },
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          spacing: { after: 200 },
          children: [new TextRun({ text: 'DEVANI, Amal  ·  Receipt No. IOE8795688197  ·  RFE Response', font: FONT, size: 18, color: '444444' })],
        })],
      }),
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: ['Page ', PageNumber.CURRENT, ' of ', PageNumber.TOTAL_PAGES], font: FONT, size: 18, color: '444444' })],
        })],
      }),
    },
    children: body,
  }],
});

Packer.toBuffer(doc).then(buf => {
  const out = '/home/user/EB1A/rfe-response/Octet-RFE-Response-IOE8795688197.docx';
  fs.writeFileSync(out, buf);
  console.log('wrote', out, buf.length, 'bytes');
});
