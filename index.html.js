const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageNumber, TabStopType, TabStopPosition,
  LevelFormat, ExternalHyperlink
} = require('docx');
const fs = require('fs');

// ── Palette ────────────────────────────────────────────────────────────────
const NAVY   = "1A2F5C";
const BLUE   = "2E5FA3";
const TEAL   = "1B7A6B";
const LIGHT  = "EEF3FB";
const BORDER_COLOR = "C5D3E8";
const WHITE  = "FFFFFF";
const GRAY   = "F4F6FA";

// ── Helpers ────────────────────────────────────────────────────────────────
function hr(color = BLUE) {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 1 } },
    spacing: { before: 80, after: 80 },
    children: []
  });
}

function thinHr() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: BORDER_COLOR, space: 1 } },
    spacing: { before: 40, after: 40 },
    children: []
  });
}

function sectionHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 100 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLUE, space: 4 } },
    children: [new TextRun({ text: text.toUpperCase(), bold: true, color: NAVY, size: 26, font: "Arial" })]
  });
}

function jobHeader(title, org, location, dates) {
  // Two-column: title+org left, location+dates right via tab stop
  return new Paragraph({
    spacing: { before: 180, after: 40 },
    tabStops: [{ type: TabStopType.RIGHT, position: 9260 }],
    children: [
      new TextRun({ text: title, bold: true, size: 22, font: "Arial", color: NAVY }),
      new TextRun({ text: "  |  ", size: 22, font: "Arial", color: "888888" }),
      new TextRun({ text: org, bold: true, size: 22, font: "Arial", color: BLUE }),
      new TextRun({ text: "\t", size: 22 }),
      new TextRun({ text: `${location}  ·  ${dates}`, size: 20, font: "Arial", color: "777777", italics: true }),
    ]
  });
}

function bullet(text, boldPart) {
  const children = [];
  if (boldPart) {
    const idx = text.indexOf(boldPart);
    if (idx >= 0) {
      children.push(new TextRun({ text: text.slice(0, idx), size: 20, font: "Arial", color: "333333" }));
      children.push(new TextRun({ text: boldPart, bold: true, size: 20, font: "Arial", color: "222222" }));
      children.push(new TextRun({ text: text.slice(idx + boldPart.length), size: 20, font: "Arial", color: "333333" }));
    } else {
      children.push(new TextRun({ text, size: 20, font: "Arial", color: "333333" }));
    }
  } else {
    children.push(new TextRun({ text, size: 20, font: "Arial", color: "333333" }));
  }
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 20, after: 20 },
    children
  });
}

function projRow(emoji, name, domain, desc, metric1, metric2, tags) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
  const borders = { top: border, bottom: border, left: border, right: border };
  const cellMargins = { top: 100, bottom: 100, left: 140, right: 140 };

  return new TableRow({
    children: [
      // Emoji + name
      new TableCell({
        borders, width: { size: 1400, type: WidthType.DXA },
        margins: cellMargins,
        verticalAlign: VerticalAlign.TOP,
        shading: { fill: LIGHT, type: ShadingType.CLEAR },
        children: [
          new Paragraph({ spacing: { before: 0, after: 20 }, children: [new TextRun({ text: emoji, size: 32, font: "Segoe UI Emoji" })] }),
          new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: name, bold: true, size: 18, font: "Arial", color: NAVY })] }),
          new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: domain, size: 16, font: "Arial", color: TEAL, italics: true })] }),
        ]
      }),
      // Description
      new TableCell({
        borders, width: { size: 4600, type: WidthType.DXA },
        margins: cellMargins,
        verticalAlign: VerticalAlign.TOP,
        children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: desc, size: 18, font: "Arial", color: "444444" })] })]
      }),
      // Metrics
      new TableCell({
        borders, width: { size: 1600, type: WidthType.DXA },
        margins: cellMargins,
        verticalAlign: VerticalAlign.TOP,
        shading: { fill: GRAY, type: ShadingType.CLEAR },
        children: [
          new Paragraph({ spacing: { before: 0, after: 8 }, children: [new TextRun({ text: metric1, bold: true, size: 20, font: "Arial", color: TEAL })] }),
          new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: metric2, size: 18, font: "Arial", color: "666666" })] }),
        ]
      }),
      // Tags
      new TableCell({
        borders, width: { size: 1760, type: WidthType.DXA },
        margins: cellMargins,
        verticalAlign: VerticalAlign.TOP,
        children: [new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: tags, size: 17, font: "Arial", color: BLUE })] })]
      }),
    ]
  });
}

function skillRow(category, items) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: BORDER_COLOR };
  const borders = { top: border, bottom: border, left: border, right: border };
  const margin = { top: 80, bottom: 80, left: 120, right: 120 };
  return new TableRow({
    children: [
      new TableCell({
        borders, width: { size: 2400, type: WidthType.DXA },
        margins: margin,
        shading: { fill: LIGHT, type: ShadingType.CLEAR },
        children: [new Paragraph({ children: [new TextRun({ text: category, bold: true, size: 19, font: "Arial", color: NAVY })] })]
      }),
      new TableCell({
        borders, width: { size: 6960, type: WidthType.DXA },
        margins: margin,
        children: [new Paragraph({ children: [new TextRun({ text: items, size: 19, font: "Arial", color: "333333" })] })]
      }),
    ]
  });
}

// ── Document ───────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 500, hanging: 280 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 20 } } },
    paragraphStyles: [
      {
        id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: "Arial", color: NAVY },
        paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 0 }
      },
      {
        id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "Arial", color: BLUE },
        paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 1 }
      },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: BLUE, space: 4 } },
          spacing: { after: 80 },
          tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
          children: [
            new TextRun({ text: "Tonye Maxwell Kentebe", bold: true, size: 20, font: "Arial", color: NAVY }),
            new TextRun({ text: "\t", size: 20 }),
            new TextRun({ text: "Computer Engineering Portfolio", size: 18, font: "Arial", color: "888888", italics: true }),
          ]
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: BORDER_COLOR, space: 4 } },
          spacing: { before: 80 },
          tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
          alignment: AlignmentType.LEFT,
          children: [
            new TextRun({ text: "University of Waterloo  ·  BASc Computer Engineering  ·  GPA 3.9/4.0", size: 17, font: "Arial", color: "888888" }),
            new TextRun({ text: "\t", size: 17 }),
            new TextRun({ text: "Page ", size: 17, font: "Arial", color: "888888" }),
            new TextRun({ children: [PageNumber.CURRENT], size: 17, font: "Arial", color: "888888" }),
          ]
        })]
      })
    },
    children: [

      // ── HERO HEADER ─────────────────────────────────────────────────────
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 60 },
        children: [new TextRun({ text: "TONYE MAXWELL KENTEBE", bold: true, size: 52, font: "Arial", color: NAVY })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 40 },
        children: [new TextRun({ text: "Computer Engineering  ·  University of Waterloo  ·  Class of 2030", size: 24, font: "Arial", color: BLUE })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
        children: [
          new TextRun({ text: "Brampton, Canada  ·  (647)-965-5159  ·  ", size: 20, font: "Arial", color: "555555" }),
          new ExternalHyperlink({
            link: "mailto:tmkenteb@uwaterloo.ca",
            children: [new TextRun({ text: "tmkenteb@uwaterloo.ca", size: 20, font: "Arial", style: "Hyperlink" })]
          }),
        ]
      }),
      hr(NAVY),

      // Award badges row
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3120, 3120, 3120],
        rows: [new TableRow({ children: [
          new TableCell({
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            width: { size: 3120, type: WidthType.DXA },
            shading: { fill: LIGHT, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
              new TextRun({ text: "🏆 Velocity Co-op Problem Award", bold: true, size: 19, font: "Arial", color: NAVY })
            ]})]
          }),
          new TableCell({
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            width: { size: 3120, type: WidthType.DXA },
            shading: { fill: LIGHT, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
              new TextRun({ text: "🎖 President's Scholarship", bold: true, size: 19, font: "Arial", color: NAVY })
            ]})]
          }),
          new TableCell({
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            width: { size: 3120, type: WidthType.DXA },
            shading: { fill: LIGHT, type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 160, right: 160 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
              new TextRun({ text: "📊 GPA: 3.9 / 4.0", bold: true, size: 19, font: "Arial", color: NAVY })
            ]})]
          }),
        ]})]
      }),

      new Paragraph({ spacing: { before: 160, after: 0 }, children: [] }),

      // ── PROFILE ─────────────────────────────────────────────────────────
      sectionHeading("Professional Profile"),
      new Paragraph({
        spacing: { before: 80, after: 80 },
        children: [new TextRun({
          text: "Second-year Computer Engineering student at the University of Waterloo with a strong record of delivering across full-stack software engineering, artificial intelligence, embedded hardware, and data analysis. Awarded the University of Waterloo Velocity Co-op Problem Award — one of only three students recognised in the term — for developing an innovative AI-powered environmental reporting system. Experienced across Python, C++, JavaScript, cloud platforms (GCP, AWS, Azure), and modern ML frameworks, with hands-on work in startup environments and multidisciplinary engineering teams.",
          size: 20, font: "Arial", color: "333333"
        })]
      }),

      // ── EDUCATION ───────────────────────────────────────────────────────
      sectionHeading("Education"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [800, 8560],
        rows: [new TableRow({ children: [
          new TableCell({
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            width: { size: 800, type: WidthType.DXA },
            shading: { fill: BLUE, type: ShadingType.CLEAR },
            margins: { top: 120, bottom: 120, left: 160, right: 160 },
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "🎓", size: 36, font: "Segoe UI Emoji" })] })]
          }),
          new TableCell({
            borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
            width: { size: 8560, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 200, right: 80 },
            children: [
              new Paragraph({ spacing: { before: 0, after: 20 }, tabStops: [{ type: TabStopType.RIGHT, position: 8360 }],
                children: [
                  new TextRun({ text: "BASc in Computer Engineering", bold: true, size: 26, font: "Arial", color: NAVY }),
                  new TextRun({ text: "\t", size: 22 }),
                  new TextRun({ text: "2025 – 2030", size: 20, font: "Arial", color: "777777", italics: true }),
                ]}),
              new Paragraph({ spacing: { before: 0, after: 20 }, children: [new TextRun({ text: "University of Waterloo  ·  Waterloo, Canada  ·  Excellent Standing", size: 21, font: "Arial", color: BLUE })] }),
              new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: "Cumulative GPA: 3.9 / 4.0  ·  Velocity Co-op Problem Award  ·  President's Scholarship  ·  Harvard CS50 Coursework", size: 19, font: "Arial", color: "555555" })] }),
            ]
          }),
        ]})]
      }),

      // ── WORK EXPERIENCE ─────────────────────────────────────────────────
      sectionHeading("Work Experience"),

      jobHeader("Software Engineer", "Bimo Technologies", "Remote / International", "Apr 2024 – Aug 2024"),
      bullet("Developed and maintained full-stack software systems using Python, JavaScript, SQL, and backend development practices", "full-stack software systems"),
      bullet("Built and optimized application features, APIs, and technical workflows supporting scalable digital platform operations", "application features, APIs, and technical workflows"),
      bullet("Collaborated with cross-functional engineering teams on debugging, deployment coordination, and iterative feature development", "cross-functional engineering teams"),
      bullet("Supported infrastructure, technical documentation, and system optimization in fast-paced startup environments", "system optimization"),
      thinHr(),

      jobHeader("Lead Web Developer", "Synsyma Technologies", "Brampton, Canada", "Jan 2025 – Apr 2025"),
      bullet("Developed responsive web applications and technical platforms using HTML, CSS, JavaScript, Python, and database-integrated workflows", "responsive web applications"),
      bullet("Assisted with frontend development, backend support, workflow automation, and technical system optimization", "frontend development, backend support, workflow automation"),
      bullet("Participated in software troubleshooting, feature implementation, and engineering-focused technical problem solving under project deadlines", "engineering-focused technical problem solving"),
      thinHr(),

      jobHeader("EHG AI Waste Analysis — Research Initiative", "Paterson Group", "Mississauga, Ontario", "Sep 2025 – Dec 2025"),
      bullet("Conducted a comprehensive Environmental Health & Governance (EHG) analysis focused on waste tracking and reporting practices across laboratory and field operations", "EHG analysis"),
      bullet("Applied software design principles, data structures, data mining, LLM-supported workflows, prompt engineering, OpenAI API, and RAG-oriented research concepts to identify inefficiencies and improve environmental reporting systems", "LLM-supported workflows, prompt engineering, OpenAI API, and RAG-oriented research"),
      new Paragraph({
        numbering: { reference: "bullets", level: 0 },
        spacing: { before: 20, after: 20 },
        children: [
          new TextRun({ text: "Awarded the ", size: 20, font: "Arial", color: "333333" }),
          new TextRun({ text: "University of Waterloo Velocity Co-op Problem Award", bold: true, size: 20, font: "Arial", color: TEAL }),
          new TextRun({ text: " — one of only three students recognised in the term for project impact and innovation", size: 20, font: "Arial", color: "333333" }),
        ]
      }),
      thinHr(),

      jobHeader("Geotechnical Engineering Assistant", "Paterson Group", "Mississauga, Ontario", "Fall 2025"),
      bullet("Used AutoCAD, Microsoft Office, and technical reporting tools to support engineering documentation and project work", "AutoCAD"),
      bullet("Worked collaboratively with team members and co-workers to support project coordination and problem-solving", "project coordination"),
      thinHr(),

      jobHeader("Hardware Team Member", "Midnight Sun Solar Car Team — UWaterloo", "Waterloo, Canada", "Mar 2025 – Present"),
      bullet("Developed and tested embedded and hardware-integrated systems for a high-performance solar vehicle", "embedded and hardware-integrated systems"),
      bullet("Applied Python and C++ for diagnostics, backend logic, technical tooling, and engineering workflow support", "Python and C++"),
      bullet("Supported sensor integration, embedded systems tasks, and system optimization in multidisciplinary engineering environments", "sensor integration"),
      bullet("Participated in testing, debugging, and iterative technical development in fast-paced project settings"),

      // ── PROJECTS ────────────────────────────────────────────────────────
      sectionHeading("ML & Engineering Projects"),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1400, 4600, 1600, 1760],
        rows: [
          // Header row
          new TableRow({
            tableHeader: true,
            children: ["Project", "Description", "Key Metric", "Technologies"].map((h, i) => {
              const widths = [1400, 4600, 1600, 1760];
              return new TableCell({
                width: { size: widths[i], type: WidthType.DXA },
                shading: { fill: NAVY, type: ShadingType.CLEAR },
                margins: { top: 80, bottom: 80, left: 140, right: 140 },
                borders: { top: { style: BorderStyle.SINGLE, size: 1, color: NAVY }, bottom: { style: BorderStyle.SINGLE, size: 1, color: NAVY }, left: { style: BorderStyle.SINGLE, size: 1, color: NAVY }, right: { style: BorderStyle.SINGLE, size: 1, color: NAVY } },
                children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 19, font: "Arial", color: WHITE })] })]
              });
            })
          }),
          projRow("🏠", "House Price Prediction", "Regression", "End-to-end regression pipeline on California Housing data. Feature engineering, outlier handling, and five models compared.", "R² ~0.82", "GBM best model", "scikit-learn\nGradient Boosting\nFeature Eng."),
          projRow("📉", "Customer Churn", "Classification", "Binary churn classifier with class imbalance handling, ROC-AUC analysis, and business-aware threshold tuning.", "AUC ~0.94", "GBM + class weights", "XGBoost\nSMOTE\nROC-AUC"),
          projRow("🧩", "Customer Segmentation", "Unsupervised ML", "RFM-based segmentation using K-Means and DBSCAN. Elbow method + silhouette analysis. PCA for visualisation.", "k=4 optimal", "PCA 90%+ var.", "K-Means\nDBSCAN\nPCA"),
          projRow("💬", "Sentiment Analysis", "NLP", "Movie review classifier using TF-IDF pipelines, bigram and character-level n-grams. Error analysis on misclassifications.", "98% accuracy", "AUC ~0.99", "TF-IDF\nLogistic Reg.\nN-grams"),
          projRow("🔍", "Fraud Detection", "Anomaly Detection", "Fraud detection on 50K transactions at 0.5% fraud rate. Isolation Forest vs. supervised models, cost-sensitive threshold tuning.", "AP ~0.92", "Cost-aware thresholds", "Isolation Forest\nPR Curves\nGBM"),
          projRow("🖼️", "Image Classification CNN", "Deep Learning", "Custom 3-block CNN on CIFAR-10 (60K images, 10 classes) with BatchNorm, Dropout, and augmentation. MobileNetV2 transfer learning.", "~88% accuracy", "MobileNetV2 TL", "TensorFlow\nKeras\nData Aug."),
        ]
      }),

      // ── SKILLS ──────────────────────────────────────────────────────────
      sectionHeading("Skills & Technologies"),

      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2400, 6960],
        rows: [
          skillRow("Programming Languages", "Python · C++ · JavaScript · SQL · HTML · CSS"),
          skillRow("Frameworks & Web", "React · REST APIs · Backend Development · Frontend Development · Database Management · Git"),
          skillRow("Cloud & Infrastructure", "GCP · AWS · Azure · Workflow Automation · Enterprise Platform Integration"),
          skillRow("AI / Machine Learning", "Machine Learning · LLMs · OpenAI API · RAG · Prompt Engineering · Data Mining · scikit-learn · TensorFlow · Keras"),
          skillRow("Data & Analytics", "Power BI · Tableau · Data Processing · Data Visualization · Technical Reporting · GIS"),
          skillRow("Engineering Tools", "AutoCAD · Autodesk Civil 3D · Tinkercad · Blender · SolidWorks · InfraWorks · Blender"),
          skillRow("Embedded & Hardware", "Embedded Systems · Sensor Integration · System Optimization · C++ Diagnostics · Hardware Testing"),
          skillRow("Professional Skills", "Cross-Functional Collaboration · Agile Development · Technical Documentation · Fast-Paced Startups · Problem Solving · Time Management"),
        ]
      }),

      // ── CERTIFICATIONS & ADDITIONAL ──────────────────────────────────────
      sectionHeading("Certifications & Additional Information"),
      bullet("Harvard CS50 Coursework — foundational computer science and programming", "Harvard CS50"),
      bullet("Self-directed C++ development for embedded and systems programming", "Self-directed C++"),
      bullet("Full-Stack Software Engineering projects encompassing frontend, backend, API integration, and database design"),
      bullet("Field qualifications: Valid Ontario G Driver's License · Own Transportation · Strong technical communication and documentation"),
      bullet("Languages: English (native)"),

      // ── CLOSING ──────────────────────────────────────────────────────────
      new Paragraph({ spacing: { before: 300, after: 0 }, children: [] }),
      hr(TEAL),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 0 },
        children: [new TextRun({ text: "References available upon request  ·  Portfolio: All 6 ML projects available as Jupyter Notebooks", size: 17, font: "Arial", color: "888888", italics: true })]
      }),
    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync("/sessions/clever-vigilant-meitner/mnt/outputs/tonye_kentebe_portfolio.docx", buf);
  console.log("✅ Portfolio created: tonye_kentebe_portfolio.docx");
}).catch(e => { console.error("Error:", e); process.exit(1); });
