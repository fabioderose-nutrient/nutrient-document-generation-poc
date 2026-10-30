// ═══════════════════════════════════════════════════════════════════
//  PIXEL-PERFECT DEMO TEMPLATES (generated client-side with `docx`)
//  Each template contains {{Field}} placeholders matching a dataset.
// ═══════════════════════════════════════════════════════════════════
import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, AlignmentType, VerticalAlign, ShadingType,
} from "docx";

const NAVY = "1F2937";
const BLUE = "2563EB";
const GRAY = "868E96";
const LIGHT = "F1F3F5";
const FONT = "Arial";

const thin = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const none = { style: BorderStyle.NONE, size: 0, color: "auto" };
const allThin = { top: thin, bottom: thin, left: thin, right: thin };
const noBorders = { top: none, bottom: none, left: none, right: none };

function run(text, opts = {}) {
  return new TextRun({ text, font: FONT, size: opts.size ?? 20, bold: opts.bold ?? false, color: opts.color ?? "212529", italics: opts.italics ?? false });
}
function p(children, opts = {}) {
  return new Paragraph({
    children: Array.isArray(children) ? children : [children],
    alignment: opts.align,
    spacing: { before: opts.before ?? 0, after: opts.after ?? 120 },
  });
}
function title(text, size = 32) {
  return p(run(text, { size, bold: true, color: NAVY }), { align: AlignmentType.CENTER, after: 60 });
}
function subtitle(text) {
  return p(run(text, { size: 20, color: GRAY }), { align: AlignmentType.CENTER, after: 240 });
}
// Usable width inside 0.75" margins on US Letter = 10080 twips
const TABLE_W = 10080;
const LABEL_W = 3530;
const VALUE_W = TABLE_W - LABEL_W;

function labelValueRow(label, placeholder, shaded = false) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: LABEL_W, type: WidthType.DXA },
        borders: allThin,
        shading: { type: ShadingType.CLEAR, fill: LIGHT },
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [p(run(label, { bold: true, size: 18, color: NAVY }), { after: 0 })],
      }),
      new TableCell({
        width: { size: VALUE_W, type: WidthType.DXA },
        borders: allThin,
        shading: shaded ? { type: ShadingType.CLEAR, fill: "EFF6FF" } : undefined,
        verticalAlign: VerticalAlign.CENTER,
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [p(run(placeholder, { size: 20, bold: shaded }), { after: 0 })],
      }),
    ],
  });
}
function infoTable(rows) {
  return new Table({
    width: { size: TABLE_W, type: WidthType.DXA },
    columnWidths: [LABEL_W, VALUE_W],
    borders: { ...allThin, insideHorizontal: thin, insideVertical: thin },
    rows,
  });
}
function spacer(h = 120) {
  return new Paragraph({ children: [], spacing: { after: h } });
}
function sectionHeading(text) {
  return p(run(text.toUpperCase(), { size: 18, bold: true, color: BLUE }), { before: 200, after: 100 });
}
function signatureBlock(role) {
  return [
    spacer(400),
    p(run("_________________________________", { size: 20 })),
    p([run(role, { bold: true, size: 18, color: NAVY })], { after: 0 }),
    p(run("This document was generated electronically and is valid without a manual signature.", { size: 14, italics: true, color: GRAY })),
  ];
}
function docWrap(children) {
  return new Document({
    styles: { default: { document: { run: { font: FONT, size: 20 } } } },
    sections: [{ properties: { page: { margin: { top: 1080, bottom: 1080, left: 1080, right: 1080 } } }, children }],
  });
}

function certificateOfTaxes() {
  return docWrap([
    title("COUNTY TREASURER"),
    subtitle("Office of the Treasurer — Property Tax Division"),
    title("CERTIFICATE OF TAXES DUE", 26),
    p([run("Certificate No: ", { bold: true, size: 18 }), run("{{CertificateNumber}}", { size: 18 }), run("      Date of Issue: ", { bold: true, size: 18 }), run("{{IssueDate}}", { size: 18 })], { align: AlignmentType.CENTER, after: 240 }),
    sectionHeading("Property Information"),
    infoTable([
      labelValueRow("Parcel Number", "{{ParcelNumber}}"),
      labelValueRow("Owner of Record", "{{OwnerName}}"),
      labelValueRow("Owner Mailing Address", "{{OwnerAddress}}"),
      labelValueRow("Property Address", "{{PropertyAddress}}"),
      labelValueRow("Legal Description", "{{LegalDescription}}"),
    ]),
    sectionHeading("Tax Summary"),
    infoTable([
      labelValueRow("Tax Year", "{{TaxYear}}"),
      labelValueRow("Assessed Value", "{{AssessedValue}}"),
      labelValueRow("Tax Amount", "{{TaxAmount}}"),
      labelValueRow("Interest", "{{InterestAmount}}"),
      labelValueRow("Certificate Fees", "{{FeesAmount}}"),
      labelValueRow("TOTAL DUE", "{{TotalDue}}", true),
    ]),
    spacer(200),
    p(run("I, the undersigned County Treasurer, do hereby certify that the amount of taxes, interest and fees due upon the real property described above, as shown by the records of this office, is as stated herein. This certificate is issued pursuant to applicable state statutes.", { size: 18 })),
    ...signatureBlock("County Treasurer"),
  ]);
}

function businessLicense() {
  return docWrap([
    title("CITY CLERK'S OFFICE"),
    subtitle("Business Licensing Division"),
    title("BUSINESS LICENSE CERTIFICATE", 26),
    p([run("License No: ", { bold: true, size: 18 }), run("{{LicenseNumber}}", { size: 18 }), run("      Status: ", { bold: true, size: 18 }), run("{{Status}}", { size: 18 })], { align: AlignmentType.CENTER, after: 240 }),
    sectionHeading("Licensee Information"),
    infoTable([
      labelValueRow("Business Name", "{{BusinessName}}"),
      labelValueRow("Owner", "{{OwnerName}}"),
      labelValueRow("Business Address", "{{BusinessAddress}}"),
      labelValueRow("License Type", "{{LicenseType}}"),
    ]),
    sectionHeading("License Terms"),
    infoTable([
      labelValueRow("Date of Issue", "{{IssueDate}}"),
      labelValueRow("Expiration Date", "{{ExpirationDate}}"),
      labelValueRow("License Fee Paid", "{{FeeAmount}}", true),
    ]),
    spacer(200),
    p(run("This license is granted to the above-named business subject to the ordinances of the City and the laws of the State. It must be displayed conspicuously at the place of business and is not transferable.", { size: 18 })),
    ...signatureBlock("City Clerk"),
  ]);
}

function w2Statement() {
  return docWrap([
    title("WAGE AND TAX STATEMENT"),
    subtitle("Tax Year {{TaxYear}} — Employer Copy (Demonstration Document)"),
    sectionHeading("Employer"),
    infoTable([
      labelValueRow("Employer Name", "{{EmployerName}}"),
      labelValueRow("Employer EIN", "{{EmployerEIN}}"),
      labelValueRow("Employer Address", "{{EmployerAddress}}"),
    ]),
    sectionHeading("Employee"),
    infoTable([
      labelValueRow("Employee Name", "{{EmployeeName}}"),
      labelValueRow("Social Security Number", "{{EmployeeSSN}}"),
      labelValueRow("Employee Address", "{{EmployeeAddress}}"),
    ]),
    sectionHeading("Wages and Withholding"),
    infoTable([
      labelValueRow("1. Wages, tips, other compensation", "{{WagesTips}}"),
      labelValueRow("2. Federal income tax withheld", "{{FederalTaxWithheld}}"),
      labelValueRow("3. Social security wages", "{{SocialSecurityWages}}"),
      labelValueRow("4. Social security tax withheld", "{{SocialSecurityTax}}"),
      labelValueRow("5. Medicare wages and tips", "{{MedicareWages}}"),
      labelValueRow("6. Medicare tax withheld", "{{MedicareTax}}"),
      labelValueRow("15. State", "{{State}}"),
      labelValueRow("16. State wages, tips, etc.", "{{StateWages}}"),
      labelValueRow("17. State income tax", "{{StateTax}}", true),
    ]),
    spacer(200),
    p(run("This is a demonstration layout for the Nutrient Document Authoring proof of concept. It is not an official IRS form.", { size: 14, italics: true, color: GRAY })),
  ]);
}

const BUILDERS = {
  certificateOfTaxes,
  businessLicense,
  w2Statement,
};

export const TEMPLATES = [
  { id: "certificateOfTaxes", name: "Certificate of Taxes Due", description: "Pixel-perfect tax certificate with merge fields", datasetId: "taxCertificates" },
  { id: "businessLicense", name: "Business License Certificate", description: "License certificate with merge fields", datasetId: "businessLicenses" },
  { id: "w2Statement", name: "Wage & Tax Statement (W-2 style)", description: "Compliance-style wage statement", datasetId: "w2Statements" },
];

const cache = {};
export async function buildTemplate(id) {
  if (!cache[id]) {
    const doc = BUILDERS[id]();
    cache[id] = await Packer.toBlob(doc);
  }
  return cache[id];
}
