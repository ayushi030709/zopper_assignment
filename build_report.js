const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        HeadingLevel, AlignmentType, WidthType, BorderStyle, ShadingType } = require("docx");
const fs = require("fs");

const darkBlue = "1F3864";
const medBlue  = "2E75B6";
const orange   = "E26B0A";
const lightBlue = "DEEAF1";

function heading1(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 240, after: 120 },
    shading: { type: ShadingType.SOLID, color: darkBlue, fill: darkBlue },
    run: { color: "FFFFFF" },
    children: [new TextRun({ text, color: "FFFFFF", bold: true, size: 26, font: "Arial" })],
  });
}

function heading2(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, size: 22, font: "Arial", color: medBlue })],
    spacing: { before: 200, after: 80 },
  });
}

function para(text, options = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 20, font: "Arial", ...options })],
    spacing: { after: 80 },
    indent: options.indent ? { left: 360 } : undefined,
  });
}

function bullet(text) {
  return new Paragraph({
    children: [new TextRun({ text: `• ${text}`, size: 20, font: "Arial" })],
    indent: { left: 360 },
    spacing: { after: 60 },
  });
}

function kv(key, val) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${key}: `, bold: true, size: 20, font: "Arial" }),
      new TextRun({ text: val, size: 20, font: "Arial", color: darkBlue }),
    ],
    indent: { left: 360 },
    spacing: { after: 60 },
  });
}

function tableRow(cells, isHeader = false) {
  return new TableRow({
    children: cells.map(c => new TableCell({
      children: [new Paragraph({
        children: [new TextRun({ text: String(c), bold: isHeader, size: 18, font: "Arial",
          color: isHeader ? "FFFFFF" : "000000" })],
        alignment: AlignmentType.CENTER,
      })],
      shading: isHeader ? { fill: darkBlue, type: ShadingType.SOLID } :
                           { fill: "F2F2F2", type: ShadingType.SOLID },
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
    })),
  });
}

const doc = new Document({
  styles: {
    default: {
      document: { run: { font: "Arial", size: 20 } },
    },
  },
  sections: [{
    properties: { page: { margin: { top: 720, bottom: 720, left: 900, right: 900 } } },
    children: [
      // Cover
      new Paragraph({
        children: [new TextRun({ text: "Car Insurance Portfolio", bold: true, size: 36, font: "Arial", color: darkBlue })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 120 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "Business Intelligence Internship Assignment", size: 24, font: "Arial", color: "555555" })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 80 },
      }),
      new Paragraph({
        children: [new TextRun({ text: "Data Simulation, Analytics & Insights Report", size: 22, font: "Arial", color: medBlue, bold: true })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 480 },
      }),

      // Section 1
      heading1("1. Approach & Methodology"),
      para("This assignment simulates 1,000,000 car insurance policies sold in 2024 and models the resulting claims in 2025–2026. Python (pandas/numpy) was used for data generation and analysis; results are presented in a structured Excel workbook with dashboards."),

      heading2("1.1 Policy Sales Simulation"),
      bullet("1,000,000 customers spread evenly across 366 days of 2024 (leap year) → ~2,732 per day."),
      bullet("Policy tenure assigned by weighted random sampling: 20% / 30% / 40% / 10% for 1/2/3/4 years."),
      bullet("Policy_Start_Date = Purchase_Date + 365 days; Policy_End_Date = Start + tenure years."),
      bullet("Vehicle value fixed at ₹1,00,000; Premium = ₹100 × tenure years."),

      heading2("1.2 Claims Simulation — 2025"),
      bullet("Only vehicles purchased on 7th, 14th, 21st, or 28th of any month in 2024 are eligible."),
      bullet("30% of eligible vehicles file exactly one claim on their Policy_Start_Date."),
      bullet("Eligibility checked: claim date must fall within active policy window."),

      heading2("1.3 Claims Simulation — 2026 (Jan 1 – Feb 28)"),
      bullet("10% of all 4-year policy holders file a claim."),
      bullet("Claims distributed evenly across 59 days (≈169 per day)."),
      bullet("Vehicles that filed in 2025 and qualify again are marked Claim_Type = 2."),

      // Section 2
      heading1("2. Key Assumptions"),
      bullet("2024 is treated as a 366-day leap year."),
      bullet("'Evenly distributed' means integer floor with leftover assigned to early days."),
      bullet("Claim amount = 10% of vehicle value = ₹10,000 per claim."),
      bullet("Daily premium = Total Premium ÷ (Tenure × 365) for earned premium calculations."),
      bullet("46 remaining months after Feb 28, 2026 used for forward premium estimation."),
      bullet("'30% of vehicles on special days' — applied as random 30% sample for realism."),

      // Section 3
      heading1("3. Analytical Query Results"),

      heading2("Q1 — Total Premium Collected (2024)"),
      kv("Total Premium", "₹24,01,10,800  (≈ ₹24 Crore)"),
      para("Calculated as sum of (Policy_Tenure × ₹100) across all 1M policies."),

      heading2("Q2 — Claim Costs by Year"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          tableRow(["Year","Total Claim Cost","Notes"], true),
          tableRow(["2025","₹39,34,40,000","39,344 claims × ₹10,000"]),
          tableRow(["2026","₹9,99,60,000","9,996 claims × ₹10,000"]),
          tableRow(["TOTAL","₹49,34,00,000","Combined 2025+2026"]),
        ],
      }),

      heading2("Q3 — Loss Ratio by Policy Tenure"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          tableRow(["Tenure (Yrs)","Total Premium (₹)","Total Claims (₹)","Loss Ratio"], true),
          tableRow(["1 Year","1,99,37,600","7,84,60,000","3.94x"]),
          tableRow(["2 Years","6,00,21,400","11,89,00,000","1.98x"]),
          tableRow(["3 Years","12,01,65,000","15,78,20,000","1.31x"]),
          tableRow(["4 Years","3,99,86,800","13,82,20,000","3.46x"]),
        ],
      }),

      heading2("Q4 — Loss Ratio by Month of Policy Sale"),
      para("Loss ratios are remarkably stable across all 12 months of 2024, ranging from ~2.00x to ~2.14x. This is expected because policies are sold uniformly every day and the claim-triggering rule (special dates) applies proportionally each month."),

      heading2("Q5 — Potential Claim Liability"),
      kv("Vehicles with no claim yet", "9,51,023"),
      kv("Claim amount per vehicle",   "₹10,000"),
      kv("Total Potential Liability",  "₹9,51,02,30,000  (≈ ₹951 Crore)"),
      para("This represents worst-case scenario if every unclaimed vehicle eventually files one claim within their policy tenure."),

      heading2("Q6 — Earned Premium Analysis"),
      kv("Earned Premium to Feb 28 2026",          "₹6,59,02,532  (≈ ₹6.59 Crore)"),
      kv("Remaining Unearned Premium",              "₹17,45,12,374"),
      kv("Est. Monthly Premium (46 months remaining)", "₹3,79,374  per month"),
      para("Policies only start earning 365 days after purchase, so only ~78% of premium has started accruing by Feb 28, 2026."),

      // Section 4
      heading1("4. Bonus Insights"),

      heading2("B1 — Most Profitable Tenure"),
      para("3-Year policies are the most financially efficient for the company:"),
      bullet("Lowest loss ratio at 1.31x — meaning for every ₹1 premium earned, ₹1.31 in claims is paid."),
      bullet("They generate the highest absolute premium (₹12 Crore) and are the most popular (40% of sales)."),
      bullet("1-Year and 4-Year policies are significantly under-priced — loss ratios of 3.94x and 3.46x respectively."),
      para("→ Recommendation: Increase premiums for 1-year and 4-year products, or implement stricter underwriting criteria for short-tenure policies.", { bold: true }),

      heading2("B2 — Claim Trends"),
      bullet("2025 claims are concentrated: they spike on dates corresponding to 7th/14th/21st/28th purchase cohorts."),
      bullet("2026 claims are smoothly distributed (Jan–Feb) since the 10% rule applies evenly."),
      bullet("No claims occur outside 2025-2026 in the simulation, but potential liability for unclaimed policies is massive."),

      heading2("B3 — Overall Portfolio Loss Ratio"),
      kv("Total Claims (2025+2026)", "₹49,34,00,000"),
      kv("Total Premium 2024",       "₹24,01,10,800"),
      kv("Portfolio Loss Ratio",     "2.05x  — severely unprofitable"),
      para("A healthy insurance loss ratio is typically 60–80%. At 205%, the portfolio is deeply unprofitable under current pricing."),

      heading2("B4 — Impact of 5% Annual Claim Frequency Increase"),
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          tableRow(["Year","Claim Frequency","Estimated Claim Cost","vs. Premium"], true),
          tableRow(["2025 (Base)","100%","₹39.34 Cr","163.8%"]),
          tableRow(["2026 (+5%)","105%","₹41.31 Cr","172.0%"]),
          tableRow(["2027 (+10%)","110%","₹43.28 Cr","180.2%"]),
          tableRow(["2029 (+20%)","120%","₹47.21 Cr","196.6%"]),
        ],
      }),
      para("Each 5% increase in claim frequency adds approximately ₹1.97 Crore annually to the loss pool. At this trajectory, premiums must be more than doubled to reach break-even."),

      // Section 5
      heading1("5. Tools & Deliverables"),
      bullet("Python (pandas, numpy, openpyxl) — data simulation and analysis"),
      bullet("Excel Workbook — 5 sheets: Dashboard, Policy Sales Data (10K sample), Claims Data (10K sample), Analytical Query Results, Bonus Insights"),
      bullet("CSV files — Full 1M policy records and 49K claims records"),
      bullet("This document — approach, assumptions and key insights"),

      heading1("6. Summary Recommendations"),
      bullet("Re-price 1-year and 4-year policies significantly upward — current premiums do not cover expected claims."),
      bullet("Investigate claim clustering around purchase-date patterns to detect potential fraud or systemic defects."),
      bullet("Monitor 4-year cohort claims beyond Feb 2026 — only 10% have claimed so far; the remaining 90% represent ₹900 Crore+ in exposure."),
      bullet("Consider reserve requirements: the ₹951 Crore potential liability dwarfs the ₹24 Crore premium collected."),
      bullet("Promote 3-year products — best balance of volume and risk-adjusted returns."),
    ],
  }],
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/home/claude/Zopper_BI_Approach_Report.docx", buf);
  console.log("Report saved!");
});
