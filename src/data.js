// ═══════════════════════════════════════════════════════════════════
//  SAMPLE DATASETS + CSV PARSING
//  All data is synthetic demo data.
// ═══════════════════════════════════════════════════════════════════

const FIRST = ["James", "Maria", "Robert", "Linda", "Michael", "Patricia", "David", "Jennifer", "William", "Elizabeth", "Richard", "Barbara", "Joseph", "Susan", "Thomas", "Jessica", "Charles", "Sarah", "Daniel", "Karen", "Matthew", "Nancy", "Anthony", "Lisa"];
const LAST = ["Smith", "Johnson", "Williams", "Brown", "Jones", "Garcia", "Miller", "Davis", "Rodriguez", "Martinez", "Hernandez", "Lopez", "Gonzalez", "Wilson", "Anderson", "Thomas", "Taylor", "Moore", "Jackson", "Martin", "Lee", "Perez", "Thompson", "White"];
const STREETS = ["Main St", "Oak Ave", "Maple Dr", "Cedar Ln", "Washington Blvd", "Park Rd", "Lake View Dr", "Hillcrest Ave", "River Rd", "Sunset Blvd", "2nd Ave", "Highland Dr"];
const CITIES = ["Springfield", "Riverton", "Fairview", "Greenville", "Franklin", "Clinton", "Madison", "Georgetown", "Salem", "Ashland", "Milford", "Oakdale"];
const BUSINESS = ["Riverside Café", "Summit Auto Repair", "Golden Gate Bakery", "Pine Hill Landscaping", "Bluebird Books", "Cornerstone Plumbing", "Lakeside Fitness", "Maple Leaf Daycare", "Ironworks Gym", "Harbor Light Seafood", "Cedar Creek Brewing", "Twin Oaks Dental"];
const LICENSE_TYPES = ["Retail Food Establishment", "General Contractor", "Professional Services", "Automotive Repair", "Child Care Facility", "Alcoholic Beverage — On Premise"];

function money(n) {
  return "$" + n.toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
}
function pad(n, w) {
  return String(n).padStart(w, "0");
}

function buildTaxCertificates(count = 24) {
  const records = [];
  for (let i = 0; i < count; i++) {
    const first = FIRST[i % FIRST.length], last = LAST[(i * 7 + 3) % LAST.length];
    const city = CITIES[i % CITIES.length];
    const assessed = 180000 + i * 13750;
    const tax = assessed * 0.0112;
    const interest = tax * 0.015 * ((i % 4) + 1);
    const fees = 25;
    records.push({
      CertificateNumber: `CTD-2026-${pad(1000 + i, 5)}`,
      IssueDate: "August 20, 2026",
      ParcelNumber: `R${pad(430000 + i * 137, 7)}`,
      OwnerName: `${first} ${last}`,
      OwnerAddress: `${100 + i * 12} ${STREETS[i % STREETS.length]}, ${city}, OH ${pad(43000 + i * 13, 5)}`,
      PropertyAddress: `${200 + i * 9} ${STREETS[(i + 4) % STREETS.length]}, ${city}, OH ${pad(43000 + i * 13, 5)}`,
      LegalDescription: `LOT ${(i % 22) + 1}, BLOCK ${(i % 9) + 1}, ${city.toUpperCase()} HEIGHTS ADDITION, SECTION ${(i % 3) + 1}`,
      TaxYear: "2025",
      AssessedValue: money(assessed),
      TaxAmount: money(tax),
      InterestAmount: money(interest),
      FeesAmount: money(fees),
      TotalDue: money(tax + interest + fees),
      Email: `${first.toLowerCase()}.${last.toLowerCase()}@example.com`,
    });
  }
  return records;
}

function buildBusinessLicenses(count = 24) {
  const records = [];
  for (let i = 0; i < count; i++) {
    const first = FIRST[(i * 5 + 2) % FIRST.length], last = LAST[i % LAST.length];
    const city = CITIES[(i + 3) % CITIES.length];
    records.push({
      LicenseNumber: `BL-2026-${pad(200 + i * 3, 5)}`,
      BusinessName: `${BUSINESS[i % BUSINESS.length]}${i >= BUSINESS.length ? " " + ["II", "North", "South", "Express"][i % 4] : ""}`,
      OwnerName: `${first} ${last}`,
      BusinessAddress: `${400 + i * 7} ${STREETS[(i + 2) % STREETS.length]}, ${city}, OH ${pad(43100 + i * 11, 5)}`,
      LicenseType: LICENSE_TYPES[i % LICENSE_TYPES.length],
      IssueDate: "January 15, 2026",
      ExpirationDate: "January 14, 2027",
      Status: i % 7 === 5 ? "Conditional" : "Active",
      FeeAmount: money(150 + (i % 5) * 75),
      Email: `${first.toLowerCase()}.${last.toLowerCase()}@example.com`,
    });
  }
  return records;
}

function buildW2Records(count = 24) {
  const records = [];
  for (let i = 0; i < count; i++) {
    const first = FIRST[(i * 3 + 1) % FIRST.length], last = LAST[(i * 11 + 5) % LAST.length];
    const city = CITIES[(i + 6) % CITIES.length];
    const wages = 42000 + i * 3875;
    records.push({
      TaxYear: "2025",
      EmployeeName: `${first} ${last}`,
      EmployeeSSN: `XXX-XX-${pad(1000 + i * 41, 4)}`,
      EmployeeAddress: `${300 + i * 5} ${STREETS[(i + 1) % STREETS.length]}, ${city}, OH ${pad(43200 + i * 17, 5)}`,
      EmployerName: "City of Riverton — Public Works Department",
      EmployerEIN: "00-0000000",
      EmployerAddress: "500 Municipal Plaza, Riverton, OH 43001",
      WagesTips: money(wages),
      FederalTaxWithheld: money(wages * 0.148),
      SocialSecurityWages: money(wages),
      SocialSecurityTax: money(wages * 0.062),
      MedicareWages: money(wages),
      MedicareTax: money(wages * 0.0145),
      State: "OH",
      StateWages: money(wages),
      StateTax: money(0),
      Email: `${first.toLowerCase()}.${last.toLowerCase()}@example.com`,
    });
  }
  return records;
}

export const DATASETS = {
  taxCertificates: {
    id: "taxCertificates",
    name: "Certificates of Taxes Due",
    description: "24 property tax certificate records (synthetic)",
    records: buildTaxCertificates(),
  },
  businessLicenses: {
    id: "businessLicenses",
    name: "Business Licenses",
    description: "24 business license records (synthetic)",
    records: buildBusinessLicenses(),
  },
  w2Statements: {
    id: "w2Statements",
    name: "Wage & Tax Statements (W-2 style)",
    description: "24 employee wage statement records (synthetic)",
    records: buildW2Records(),
  },
};

export function datasetFields(dataset) {
  return dataset && dataset.records.length ? Object.keys(dataset.records[0]) : [];
}

// Minimal CSV parser: handles quoted fields with commas and escaped quotes.
export function parseCSV(text) {
  const rows = [];
  let row = [], field = "", inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { field += '"'; i++; }
        else inQuotes = false;
      } else field += ch;
    } else if (ch === '"') {
      inQuotes = true;
    } else if (ch === ",") {
      row.push(field); field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field); field = "";
      if (row.length > 1 || row[0] !== "") rows.push(row);
      row = [];
    } else field += ch;
  }
  if (field !== "" || row.length) { row.push(field); rows.push(row); }
  if (rows.length < 2) return null;
  const headers = rows[0].map((h) => h.trim());
  const records = rows.slice(1).map((r) => {
    const rec = {};
    headers.forEach((h, i) => { rec[h] = (r[i] ?? "").trim(); });
    return rec;
  });
  return { headers, records };
}
