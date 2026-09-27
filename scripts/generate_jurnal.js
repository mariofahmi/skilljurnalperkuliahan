/**
 * generate_jurnal.js — Membuat dokumen JURNAL PERKULIAHAN UNIROW (.docx)
 *
 * Tabel: No | Hari | Tanggal | Jam | Makul | Pokok Bahasan
 * Desain presisi tinggi (KOP 3-kolom simetris, identitas tabular laser-aligned, 
 * kolom pokok bahasan proporsional luas, kelas dikosongkan dengan titik-titik isian).
 *
 * Pemakaian:
 *   node generate_jurnal.js data.json output.docx [logo.png]
 */
const fs = require("fs");
const path = require("path");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, BorderStyle, ShadingType,
  ImageRun, VerticalAlign, PageBreak, UnderlineType,
} = require("docx");

const [, , dataPath, outPath, logoPathArg] = process.argv;
if (!dataPath || !outPath) {
  console.error("Usage: node generate_jurnal.js data.json output.docx [logo.png]");
  process.exit(1);
}
const data = JSON.parse(fs.readFileSync(dataPath, "utf-8"));
const logoPath = logoPathArg || path.join(__dirname, "..", "assets", "logo_unirow.png");
const logoBuffer = fs.existsSync(logoPath) ? fs.readFileSync(logoPath) : null;

const FONT = "Times New Roman";
const jumlahPertemuan = data.jumlah_pertemuan || 16;
const baris = data.baris || [];

const thinBorder = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const allBorders = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const noBorders = {
  top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
  left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
};

function cellText(text, { bold = false, size = 17, align = AlignmentType.LEFT, underline = false } = {}) {
  return new Paragraph({
    alignment: align,
    spacing: { before: 0, after: 0, line: 250 },
    children: [new TextRun({ text: String(text ?? ""), bold, font: FONT, size, underline: underline ? { type: UnderlineType.SINGLE } : undefined })],
  });
}

function headerCell(text, { width } = {}) {
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    borders: allBorders,
    verticalAlign: VerticalAlign.CENTER,
    shading: { type: ShadingType.CLEAR, fill: "E2E2E2" },
    margins: { top: 60, bottom: 60, left: 50, right: 50 },
    children: [cellText(text, { bold: true, size: 19, align: AlignmentType.CENTER })],
  });
}

function dataCell(text, { width, align = AlignmentType.CENTER } = {}) {
  return new TableCell({
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    borders: allBorders,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 35, bottom: 35, left: 50, right: 50 },
    children: [cellText(text, { align, size: 17 })],
  });
}

// ---- Lebar kolom (DXA) — halaman A4 portrait (Margin kiri/kanan 1080 dxa, Lebar konten 9746 dxa) ----
const W_NO = 450;
const W_HARI = 950;
const W_TANGGAL = 1250;
const W_JAM = 1050;
const W_MAKUL = 1950;
const W_POKOK = 4096;
const colWidths = [W_NO, W_HARI, W_TANGGAL, W_JAM, W_MAKUL, W_POKOK];
const tableWidth = colWidths.reduce((a, b) => a + b, 0); // 9746 DXA

// ---- Header ----
const headerRow = new TableRow({
  tableHeader: true,
  cantSplit: true,
  children: [
    headerCell("No", { width: W_NO }),
    headerCell("Hari", { width: W_HARI }),
    headerCell("Tanggal", { width: W_TANGGAL }),
    headerCell("Jam", { width: W_JAM }),
    headerCell("Makul", { width: W_MAKUL }),
    headerCell("Pokok Bahasan", { width: W_POKOK }),
  ],
});

// ---- Baris data ----
const dataRows = [];
for (let i = 1; i <= jumlahPertemuan; i++) {
  const row = baris.find((b) => Number(b.no) === i) || {};
  dataRows.push(new TableRow({
    cantSplit: true,
    children: [
      dataCell(row.no ?? i, { width: W_NO }),
      dataCell(row.hari ?? "", { width: W_HARI }),
      dataCell(row.tanggal ?? "", { width: W_TANGGAL }),
      dataCell(row.jam ?? "", { width: W_JAM }),
      dataCell(row.makul ?? "", { width: W_MAKUL, align: AlignmentType.LEFT }),
      dataCell(row.pokok_bahasan ?? "", { width: W_POKOK, align: AlignmentType.LEFT }),
    ],
  }));
}

const jurnalTable = new Table({
  width: { size: tableWidth, type: WidthType.DXA },
  columnWidths: colWidths,
  rows: [headerRow, ...dataRows],
});

// ---- Kop surat 3-Kolom Simetris (Logo 1200, Teks 7346, Spacer 1200) ----
const kopParagraphs = [
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 260 },
    children: [new TextRun({ text: "JURNAL PERKULIAHAN", bold: true, font: FONT, size: 30 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 260 },
    children: [new TextRun({ text: `TAHUN AKADEMIK ${data.tahun_akademik || "20... / 20..."}`.toUpperCase(), bold: true, font: FONT, size: 22 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 260 },
    children: [new TextRun({ text: "UNIVERSITAS PGRI RONGGOLAWE TUBAN", bold: true, font: FONT, size: 28 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 240 },
    children: [new TextRun({ text: "Jl. Manunggal 61 Tuban, Telp. (0356) 322233, Fax (0356) 331578", font: FONT, size: 18 })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 240 },
    children: [new TextRun({ text: "Website: www.unirow.ac.id, Email: prospective@unirow.ac.id", font: FONT, size: 18 })],
  }),
];

const logoImageChildren = logoBuffer
  ? [new ImageRun({ data: logoBuffer, type: "png", transformation: { width: 59, height: 60 } })]
  : [];

const kopTable = new Table({
  width: { size: tableWidth, type: WidthType.DXA },
  columnWidths: [1200, 7346, 1200],
  borders: noBorders,
  rows: [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          width: { size: 1200, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          borders: noBorders,
          margins: { top: 0, bottom: 0, left: 0, right: 30 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: logoImageChildren })],
        }),
        new TableCell({
          width: { size: 7346, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          borders: noBorders,
          margins: { top: 0, bottom: 0, left: 20, right: 20 },
          children: kopParagraphs,
        }),
        new TableCell({
          width: { size: 1200, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          borders: noBorders,
          children: [new Paragraph({ children: [] })],
        }),
      ],
    }),
  ],
});

// Garis batas kop (double line)
const kopDividerParagraph = new Paragraph({
  spacing: { before: 40, after: 120 },
  border: { bottom: { style: BorderStyle.DOUBLE, size: 12, color: "000000", space: 4 } },
  children: [new TextRun({ text: "", font: FONT })],
});

// ---- Normalisasi Semester/Kelas: Kosongkan bagian kelas (titik-titik isian) ----
let semKelasVal = data.semester_kelas;
if (!semKelasVal) {
  semKelasVal = data.semester ? `${data.semester} / ..........` : "1 / ..........";
} else if (/\/\s*[a-zA-Z0-9]+$/.test(semKelasVal)) {
  semKelasVal = semKelasVal.replace(/\/\s*[a-zA-Z0-9]+$/, "/ ..........");
}

// ---- Blok identitas (Tabular Laser-Aligned tanpa border: Label 1900, Colon 200, Value 7646) ----
function idRow(label, value) {
  return new TableRow({
    cantSplit: true,
    children: [
      new TableCell({
        width: { size: 1900, type: WidthType.DXA },
        borders: noBorders,
        margins: { top: 15, bottom: 15, left: 0, right: 0 },
        children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, font: FONT, size: 20 })] })],
      }),
      new TableCell({
        width: { size: 200, type: WidthType.DXA },
        borders: noBorders,
        margins: { top: 15, bottom: 15, left: 0, right: 0 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: ":", bold: true, font: FONT, size: 20 })] })],
      }),
      new TableCell({
        width: { size: 7646, type: WidthType.DXA },
        borders: noBorders,
        margins: { top: 15, bottom: 15, left: 0, right: 0 },
        children: [new Paragraph({ children: [new TextRun({ text: ` ${value || ""}`, font: FONT, size: 20 })] })],
      }),
    ],
  });
}

const identityTable = new Table({
  width: { size: tableWidth, type: WidthType.DXA },
  columnWidths: [1900, 200, 7646],
  borders: noBorders,
  rows: [
    idRow("Fakultas", "Keguruan dan Ilmu Pendidikan (FKIP)"),
    idRow("Program Studi", data.prodi),
    idRow("Dosen Pengampu", data.dosen),
    idRow("Semester/Kelas", semKelasVal),
  ],
});

// ---- Blok tanda tangan penutup ----
const tglPenutup = data.tanggal_penutup || ".........................";
const kaprodiName = data.kaprodi || "Mario Fahmi Syahrial, M.Pd.";
const dosenName = data.dosen || "........................................";

const signatureTable = new Table({
  width: { size: tableWidth, type: WidthType.DXA },
  columnWidths: [Math.floor(tableWidth / 2), tableWidth - Math.floor(tableWidth / 2)],
  borders: noBorders,
  rows: [
    new TableRow({
      cantSplit: true,
      children: [
        new TableCell({
          borders: noBorders,
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 260 }, children: [new TextRun({ text: "Mengetahui,\nKetua Program Studi", font: FONT, size: 20 })] }),
            new Paragraph({ spacing: { before: 800, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: kaprodiName, bold: true, underline: { type: UnderlineType.SINGLE }, font: FONT, size: 20 })] }),
          ],
        }),
        new TableCell({
          borders: noBorders,
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, spacing: { line: 260 }, children: [new TextRun({ text: `Tuban, ${tglPenutup}\nDosen Pengampu,`, font: FONT, size: 20 })] }),
            new Paragraph({ spacing: { before: 800, after: 0 }, alignment: AlignmentType.CENTER, children: [new TextRun({ text: dosenName, bold: true, underline: { type: UnderlineType.SINGLE }, font: FONT, size: 20 })] }),
          ],
        }),
      ],
    }),
  ],
});

const doc = new Document({
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 850, right: 1080, bottom: 850, left: 1080 },
        },
      },
      children: [
        kopTable,
        kopDividerParagraph,
        identityTable,
        new Paragraph({ spacing: { before: 0, after: 80 }, children: [] }),
        jurnalTable,
        new Paragraph({ spacing: { before: 160, after: 0 }, children: [] }),
        signatureTable,
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync(outPath, buffer);
  console.log(`[OK] Berhasil: ${outPath}`);
});
