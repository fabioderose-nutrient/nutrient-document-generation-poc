// Writes samples/sample-upload.csv: five fictitious property tax records whose
// columns match the {{fields}} of the "Certificate of Taxes Due" template.
// Upload it in the Data panel to switch the data source to a CSV file.
//
//   npm run samples
//
import { mkdir, writeFile } from "node:fs/promises";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const OUT = join(dirname(fileURLToPath(import.meta.url)), "..", "samples", "sample-upload.csv");

const PEOPLE = [
  ["Jordan", "Avery"],
  ["Casey", "Brooks"],
  ["Morgan", "Ellis"],
  ["Riley", "Parker"],
  ["Quinn", "Hayes"],
];
const STREETS = ["Example Ct", "Sample Ave", "Demo Rd", "Placeholder Ln", "Test Blvd"];
const CITY = "Springfield, OH";

const money = (n) => "$" + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
const csvCell = (v) => (/[",\n]/.test(v) ? `"${String(v).replace(/"/g, '""')}"` : String(v));

const header = [
  "CertificateNumber", "IssueDate", "ParcelNumber", "OwnerName", "OwnerAddress", "PropertyAddress",
  "LegalDescription", "TaxYear", "AssessedValue", "TaxAmount", "InterestAmount", "FeesAmount", "TotalDue", "Email",
];

const rows = PEOPLE.map(([first, last], i) => {
  const assessed = 250000 + i * 41250;
  const tax = assessed * 0.0112;
  const interest = tax * 0.015;
  const fees = 25;
  const zip = String(43010 + i * 7).padStart(5, "0");
  return {
    CertificateNumber: `CTD-2026-${String(90001 + i).padStart(5, "0")}`,
    IssueDate: "August 20, 2026",
    ParcelNumber: `R${String(9100001 + i).padStart(7, "0")}`,
    OwnerName: `${first} ${last}`,
    OwnerAddress: `${10 + i * 11} ${STREETS[i]}, ${CITY} ${zip}`,
    PropertyAddress: `${200 + i * 23} ${STREETS[(i + 2) % STREETS.length]}, ${CITY} ${zip}`,
    LegalDescription: `LOT ${i + 3}, BLOCK ${i + 1}, SAMPLE ESTATES ADDITION`,
    TaxYear: "2025",
    AssessedValue: money(assessed),
    TaxAmount: money(tax),
    InterestAmount: money(interest),
    FeesAmount: money(fees),
    TotalDue: money(tax + interest + fees),
    Email: `${first.toLowerCase()}.${last.toLowerCase()}@example.com`,
  };
});

const csv = [header.join(","), ...rows.map((r) => header.map((h) => csvCell(r[h])).join(","))].join("\n") + "\n";
await mkdir(dirname(OUT), { recursive: true });
await writeFile(OUT, csv);
console.log(`Wrote ${rows.length} records to samples/sample-upload.csv`);
