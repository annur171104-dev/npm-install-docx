const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  VerticalAlign, PageOrientation, LevelFormat
} = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "AAAAAA" };
const borders = { top: border, bottom: border, left: border, right: border };
const thickBorder = { style: BorderStyle.SINGLE, size: 4, color: "2E6DA4" };
const thickBorders = { top: thickBorder, bottom: thickBorder, left: thickBorder, right: thickBorder };

const cellMargin = { top: 100, bottom: 100, left: 120, right: 120 };

function headerCell(text, widthDxa, shade = "2E6DA4") {
  return new TableCell({
    borders: thickBorders,
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: cellMargin,
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold: true, color: "FFFFFF", size: 20, font: "Arial" })]
    })]
  });
}

function dataCell(text, widthDxa, shade = "FFFFFF", isCenter = false) {
  return new TableCell({
    borders,
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { fill: shade, type: ShadingType.CLEAR },
    margins: cellMargin,
    verticalAlign: VerticalAlign.TOP,
    children: [new Paragraph({
      alignment: isCenter ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, size: 18, font: "Arial" })]
    })]
  });
}

function noteCell(text, widthDxa) {
  return new TableCell({
    borders,
    width: { size: widthDxa, type: WidthType.DXA },
    shading: { fill: "FFF9E6", type: ShadingType.CLEAR },
    margins: cellMargin,
    verticalAlign: VerticalAlign.TOP,
    children: [new Paragraph({
      children: [new TextRun({ text, size: 18, font: "Arial", italics: true, color: "CC6600" })]
    })]
  });
}

function makePara(text, bold = false, size = 22, color = "222222", spacing = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80, ...spacing },
    children: [new TextRun({ text, bold, size, font: "Arial", color })]
  });
}

function sectionTitle(text) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, bold: true, size: 26, font: "Arial", color: "2E6DA4" })],
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E6DA4", space: 1 } }
  });
}

// DATA SETS - Teacher Interview Questions
const guruRows = [
  {
    no: "1",
    kompetensi: "Pedagogis",
    indikatorKomp: "Pemahaman karakteristik peserta didik",
    indikatorRapor: "Guru memahami latar belakang, karakteristik, dan kebutuhan peserta didik",
    pertanyaan: "Bagaimana cara Anda memahami karakteristik dan kebutuhan peserta didik?",
    shade: "F5F5F5"
  },
  {
    no: "2",
    kompetensi: "Pedagogis",
    indikatorKomp: "Merancang pembelajaran yang mendidik",
    indikatorRapor: "Guru merancang rencana pembelajaran yang sesuai dengan kurikulum",
    pertanyaan: "Jelaskan proses perencanaan pembelajaran Anda untuk satu unit pembelajaran",
    shade: "FFFFFF"
  },
  {
    no: "3",
    kompetensi: "Pedagogis",
    indikatorKomp: "Implementasi pembelajaran yang mendidik",
    indikatorRapor: "Guru melaksanakan proses pembelajaran yang bermutu",
    pertanyaan: "Bagaimana Anda menciptakan lingkungan belajar yang kondusif?",
    shade: "F5F5F5"
  },
  {
    no: "4",
    kompetensi: "Pedagogis",
    indikatorKomp: "Pemanfaatan TIK dalam pembelajaran",
    indikatorRapor: "Guru menggunakan teknologi untuk meningkatkan kualitas pembelajaran",
    pertanyaan: "Teknologi apa saja yang Anda gunakan dalam pembelajaran?",
    shade: "FFFFFF"
  },
  {
    no: "5",
    kompetensi: "Pedagogis",
    indikatorKomp: "Evaluasi hasil belajar",
    indikatorRapor: "Guru melakukan penilaian pembelajaran secara komprehensif",
    pertanyaan: "Bagaimana strategi penilaian yang Anda gunakan?",
    shade: "F5F5F5"
  },
  {
    no: "6",
    kompetensi: "Profesional",
    indikatorKomp: "Penguasaan materi pembelajaran",
    indikatorRapor: "Guru menguasai materi pelajaran dengan baik",
    pertanyaan: "Bagaimana Anda tetap update dengan perkembangan materi pelajaran?",
    shade: "FFFFFF"
  },
  {
    no: "7",
    kompetensi: "Profesional",
    indikatorKomp: "Pengembangan keprofesionalan berkelanjutan",
    indikatorRapor: "Guru mengikuti pengembangan profesional secara berkelanjutan",
    pertanyaan: "Kegiatan pengembangan profesional apa yang telah Anda ikuti?",
    shade: "F5F5F5"
  },
  {
    no: "8",
    kompetensi: "Kepribadian",
    indikatorKomp: "Integritas kepribadian",
    indikatorRapor: "Guru menunjukkan integritas kepribadian yang kuat",
    pertanyaan: "Bagaimana Anda menjadi teladan bagi peserta didik?",
    shade: "FFFFFF"
  },
  {
    no: "9",
    kompetensi: "Kepribadian",
    indikatorKomp: "Kepribadian yang matang dan mandiri",
    indikatorRapor: "Guru menunjukkan kepribadian yang matang",
    pertanyaan: "Bagaimana Anda menangani situasi yang sulit?",
    shade: "F5F5F5"
  },
  {
    no: "10",
    kompetensi: "Sosial",
    indikatorKomp: "Berkomunikasi efektif dengan peserta didik",
    indikatorRapor: "Guru berkomunikasi dengan efektif dan empatik",
    pertanyaan: "Bagaimana strategi komunikasi Anda dengan peserta didik?",
    shade: "FFFFFF"
  },
  {
    no: "11",
    kompetensi: "Sosial",
    indikatorKomp: "Berkomunikasi dengan orang tua/wali",
    indikatorRapor: "Guru membangun kerjasama dengan orang tua/wali",
    pertanyaan: "Bagaimana Anda melibatkan orang tua dalam proses pembelajaran?",
    shade: "F5F5F5"
  },
  {
    no: "12",
    kompetensi: "Sosial",
    indikatorKomp: "Kerjasama dengan sesama guru",
    indikatorRapor: "Guru berkolaborasi dengan rekan kerja",
    pertanyaan: "Bagaimana Anda berkolaborasi dengan guru lain?",
    shade: "FFFFFF"
  },
  {
    no: "13",
    kompetensi: "Kompetensi Tambahan",
    indikatorKomp: "Adaptasi terhadap perubahan kurikulum",
    indikatorRapor: "Guru dapat beradaptasi dengan perubahan kurikulum",
    pertanyaan: "Bagaimana pengalaman Anda menghadapi perubahan kurikulum?",
    shade: "F5F5F5"
  },
  {
    no: "14",
    kompetensi: "Kompetensi Tambahan",
    indikatorKomp: "Inovasi dalam pembelajaran",
    indikatorRapor: "Guru mengembangkan inovasi pembelajaran",
    pertanyaan: "Inovasi apa yang pernah Anda coba dalam pembelajaran?",
    shade: "FFFFFF"
  },
  {
    no: "15",
    kompetensi: "Kompetensi Tambahan",
    indikatorKomp: "Keterlibatan dalam ekosistem sekolah",
    indikatorRapor: "Guru terlibat aktif dalam kehidupan sekolah",
    pertanyaan: "Kegiatan sekolah apa yang Anda ikuti di luar pembelajaran?",
    shade: "F5F5F5"
  }
];

// DATA SETS - Principal Interview Questions
const kepsekRows = [
  {
    no: "1",
    kompetensi: "Manajerial",
    indikatorKomp: "Perencanaan strategis sekolah",
    indikatorRapor: "Kepala sekolah merencanakan pengembangan sekolah secara strategis",
    pertanyaan: "Bagaimana visi dan misi sekolah dalam pengembangan pendidikan?",
    shade: "F5F5F5"
  },
  {
    no: "2",
    kompetensi: "Manajerial",
    indikatorKomp: "Pengelolaan sarana dan prasarana",
    indikatorRapor: "Kepala sekolah mengelola sarana dan prasarana dengan baik",
    pertanyaan: "Bagaimana strategi Anda dalam pengelolaan infrastruktur sekolah?",
    shade: "FFFFFF"
  },
  {
    no: "3",
    kompetensi: "Manajerial",
    indikatorKomp: "Pengelolaan tenaga kependidikan",
    indikatorRapor: "Kepala sekolah membina dan mengembangkan tenaga pendidik",
    pertanyaan: "Program apa yang Anda tawarkan untuk pengembangan guru?",
    shade: "F5F5F5"
  },
  {
    no: "4",
    kompetensi: "Kewirausahaan",
    indikatorKomp: "Penciptaan inovasi sekolah",
    indikatorRapor: "Kepala sekolah menciptakan inovasi yang memberikan nilai tambah",
    pertanyaan: "Inovasi apa yang telah Anda implementasikan di sekolah ini?",
    shade: "FFFFFF"
  },
  {
    no: "5",
    kompetensi: "Kewirausahaan",
    indikatorKomp: "Kemitraan dengan stakeholder",
    indikatorRapor: "Kepala sekolah membangun kemitraan yang produktif",
    pertanyaan: "Bagaimana Anda membangun kerjasama dengan masyarakat dan industri?",
    shade: "F5F5F5"
  },
  {
    no: "6",
    kompetensi: "Supervisi",
    indikatorKomp: "Pemantauan pembelajaran",
    indikatorRapor: "Kepala sekolah melakukan monitoring terhadap proses pembelajaran",
    pertanyaan: "Bagaimana sistem supervisi pembelajaran Anda?",
    shade: "FFFFFF"
  },
  {
    no: "7",
    kompetensi: "Supervisi",
    indikatorKomp: "Pembinaan guru",
    indikatorRapor: "Kepala sekolah membina guru dalam meningkatkan kualitas mengajar",
    pertanyaan: "Program pembinaan guru apa yang telah Anda lakukan?",
    shade: "F5F5F5"
  },
  {
    no: "8",
    kompetensi: "Supervisi",
    indikatorKomp: "Evaluasi program sekolah",
    indikatorRapor: "Kepala sekolah melakukan evaluasi program secara berkala",
    pertanyaan: "Bagaimana mekanisme evaluasi program sekolah Anda?",
    shade: "FFFFFF"
  },
  {
    no: "9",
    kompetensi: "Sosial",
    indikatorKomp: "Komunikasi dengan stakeholder",
    indikatorRapor: "Kepala sekolah berkomunikasi dengan baik ke semua pihak",
    pertanyaan: "Bagaimana Anda membangun komunikasi efektif dengan orang tua?",
    shade: "F5F5F5"
  },
  {
    no: "10",
    kompetensi: "Sosial",
    indikatorKomp: "Kepemimpinan yang kolaboratif",
    indikatorRapor: "Kepala sekolah menunjukkan kepemimpinan yang inklusif",
    pertanyaan: "Bagaimana gaya kepemimpinan Anda melibatkan semua pihak?",
    shade: "FFFFFF"
  },
  {
    no: "11",
    kompetensi: "Pribadi",
    indikatorKomp: "Integritas dan akuntabilitas",
    indikatorRapor: "Kepala sekolah menunjukkan integritas dan bertanggung jawab",
    pertanyaan: "Bagaimana Anda mempertanggungjawabkan kinerja sekolah?",
    shade: "F5F5F5"
  },
  {
    no: "12",
    kompetensi: "Pribadi",
    indikatorKomp: "Kepemimpinan strategis",
    indikatorRapor: "Kepala sekolah menunjukkan visi yang jelas dalam memimpin",
    pertanyaan: "Apa tantangan terbesar dalam memimpin sekolah ini?",
    shade: "FFFFFF"
  },
  {
    no: "13",
    kompetensi: "Administrasi",
    indikatorKomp: "Pengelolaan keuangan sekolah",
    indikatorRapor: "Kepala sekolah mengelola keuangan secara transparan dan akuntabel",
    pertanyaan: "Bagaimana sistem pengelolaan keuangan sekolah Anda?",
    shade: "F5F5F5"
  },
  {
    no: "14",
    kompetensi: "Administrasi",
    indikatorKomp: "Dokumentasi dan pelaporan",
    indikatorRapor: "Kepala sekolah menyediakan dokumentasi lengkap dan pelaporan tepat waktu",
    pertanyaan: "Bagaimana sistem dokumentasi dan pelaporan di sekolah?",
    shade: "FFFFFF"
  },
  {
    no: "15",
    kompetensi: "Pengembangan Berkelanjutan",
    indikatorKomp: "Pembelajaran organisasi",
    indikatorRapor: "Kepala sekolah terus belajar dan berkembang",
    pertanyaan: "Program pengembangan profesional apa yang Anda ikuti?",
    shade: "F5F5F5"
  }
];

const COL = [500, 1300, 1900, 2200, 3126];

function makeTableRows(data) {
  const rows = [
    new TableRow({
      tableHeader: true,
      children: [
        headerCell("No", COL[0]),
        headerCell("Kompetensi", COL[1]),
        headerCell("Indikator Kompetensi", COL[2]),
        headerCell("Indikator dalam Rapor Pendidikan", COL[3]),
        headerCell("Pertanyaan Wawancara", COL[4]),
      ]
    })
  ];

  data.forEach((row) => {
    rows.push(new TableRow({
      children: [
        dataCell(row.no, COL[0], row.shade, true),
        dataCell(row.kompetensi, COL[1], row.shade),
        dataCell(row.indikatorKomp, COL[2], row.shade),
        noteCell(row.indikatorRapor, COL[3]),
        dataCell(row.pertanyaan, COL[4], row.shade),
      ]
    }));
  });
  return rows;
}

function makeTable(data) {
  return new Table({
    width: { size: 9026, type: WidthType.DXA },
    columnWidths: COL,
    rows: makeTableRows(data),
  });
}

const doc = new Document({
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
        }
      },
      children: [
        // Cover Page
        new Paragraph({ text: "", spacing: { after: 240 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: "INSTRUMEN WAWANCARA", bold: true, size: 32, font: "Arial", color: "2E6DA4" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: "KOMPETENSI GURU DAN KEPALA SEKOLAH", bold: true, size: 28, font: "Arial", color: "2E6DA4" })]
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 360 },
          children: [new TextRun({ text: "SDN Mertan 04", size: 24, font: "Arial" })]
        }),
        new Paragraph({ text: "", spacing: { after: 480 } }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "Tahun 2026", size: 22, font: "Arial" })]
        }),

        // Teacher Section
        new Paragraph({ text: "\n\n\n", spacing: { after: 120 } }),
        sectionTitle("I. WAWANCARA KOMPETENSI GURU"),
        makePara("Bagian ini berisi instrumen wawancara untuk mengukur kompetensi guru yang meliputi: kompetensi pedagogis, profesional, kepribadian, dan sosial.", false, 22),

        makeTable(guruRows),

        // Principal Section
        new Paragraph({ text: "\n\n\n", spacing: { after: 120 } }),
        sectionTitle("II. WAWANCARA KOMPETENSI KEPALA SEKOLAH"),
        makePara("Bagian ini berisi instrumen wawancara untuk mengukur kompetensi kepala sekolah yang meliputi: kompetensi manajerial, kewirausahaan, supervisi, sosial, pribadi, dan administrasi.", false, 22),

        makeTable(kepsekRows),

        // Notes Section
        new Paragraph({ text: "\n\n", spacing: { after: 120 } }),
        sectionTitle("III. CATATAN PENTING"),
        makePara("1. Wawancara dilakukan oleh pewawancara yang telah dilatih", false, 22),
        makePara("2. Setiap pertanyaan dirancang untuk menggali informasi mendalam tentang kompetensi", false, 22),
        makePara("3. Responden diberikan kebebasan untuk menjawab dengan jujur dan lengkap", false, 22),
        makePara("4. Hasil wawancara dicatat dengan teliti untuk analisis lebih lanjut", false, 22),
      ]
    }
  ]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("Instrumen_Wawancara_Kompetensi_Guru_SDN_Mertan04.docx", buffer);
  console.log("Document generated successfully!");
});