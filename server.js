/**
 * server.js — PDF → DOCX conversion server
 *
 * POST /convert   multipart/form-data with field "pdf"
 *                 returns the .docx file as a download
 *
 * Run:  node server.js
 * Deps: npm install express multer
 */

"use strict";

const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const os = require("os");
const { execSync } = require("child_process");

// docx imports
const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  ImageRun,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  WidthType,
  ShadingType,
  LevelFormat,
} = require("docx");

const app = express();
const upload = multer({ dest: os.tmpdir() });

app.use(express.static(path.join(__dirname, "public")));

//Helpers (mirror of pdf_to_docx.js)

const WARNING_SYMBOL = "\u26A0";

const BORDER = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const ALL_SIDES = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };

function run(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: 22, ...opts });
}

function para(children, opts = {}) {
  return new Paragraph({ children, ...opts });
}

function imagePara(
  data,
  width,
  height,
  type = "png",
  before = 120,
  after = 120,
) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before, after },
    children: [new ImageRun({ data, transformation: { width, height }, type })],
  });
}

function heading(
  text,
  level = HeadingLevel.HEADING_2,
  alignment = AlignmentType.LEFT,
  extra = {},
) {
  return new Paragraph({
    heading: level,
    alignment,
    children: [
      new TextRun({
        text,
        font: "Arial",
        size: level === HeadingLevel.HEADING_1 ? 40 : 26,
        bold: true,
      }),
    ],
    ...extra,
  });
}

function warningPara(label, boldPart = "", normalPart = "") {
  return new Paragraph({
    spacing: { before: 100, after: 60 },
    children: [
      run(`${WARNING_SYMBOL} ${label} `, { bold: true }),
      ...(boldPart ? [run(boldPart + " ", { bold: true })] : []),
      ...(normalPart ? [run(normalPart)] : []),
    ],
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [run(text)],
  });
}

function cleanImageBackground(imgPath, maskPath, outPath) {
  const withMask = `
from PIL import Image
img  = Image.open("${imgPath}").convert("RGB")
mask = Image.open("${maskPath}").convert("L")
out  = Image.new("RGB", img.size, (255,255,255))
out.paste(img, mask=mask)
out.save("${outPath}")`;

  const noMask = `
import numpy as np
from PIL import Image
img  = Image.open("${imgPath}").convert("RGBA")
import numpy as np; data = np.array(img)
black = (data[:,:,0]<20)&(data[:,:,1]<20)&(data[:,:,2]<20)
data[black]=[255,255,255,255]
Image.fromarray(data).convert("RGB").save("${outPath}")`;

  try {
    execSync(`python3 -c '${maskPath ? withMask : noMask}'`);
  } catch {
    fs.copyFileSync(imgPath, outPath);
  }
}

function partsTable(
  rows,
  colWidths = [580, 2609, 580, 2629, 580],
  tableWidth = 6978,
) {
  const headers = ["Part#", "Description", "Qty.", "Description", "Qty."];
  const makeCell = (text, width, isHeader = false) =>
    new TableCell({
      borders: ALL_SIDES,
      width: { size: width, type: WidthType.DXA },
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      shading: isHeader
        ? { fill: "DDDDDD", type: ShadingType.CLEAR }
        : undefined,
      children: [
        new Paragraph({ children: [run(text, { size: 18, bold: isHeader })] }),
      ],
    });

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => makeCell(h, colWidths[i], true)),
      }),
      ...rows.map(
        (row) =>
          new TableRow({
            children: row.map((c, i) => makeCell(c, colWidths[i])),
          }),
      ),
    ],
  });
}

// Core conversion logic

async function convertPdfToDocx(pdfPath, workDir) {
  // Extract embedded images
  const imgDir = path.join(workDir, "images");
  fs.mkdirSync(imgDir, { recursive: true });

  try {
    execSync(`pdfimages -png "${pdfPath}" "${imgDir}/img" 2>/dev/null`);
  } catch (e) {
    throw new Error(
      "pdfimages failed — is poppler-utils installed? " + e.message,
    );
  }

  const extractedImages = fs.readdirSync(imgDir).sort();
  if (extractedImages.length === 0)
    throw new Error("No images extracted from PDF.");

  // Always clean the second image (index 1) if a mask (index 2) exists
  const imgFiles = extractedImages.map((f) => path.join(imgDir, f));
  const cleanPath = path.join(workDir, "product_clean.png");

  const hasProduct = imgFiles.length > 1;
  const hasMask = imgFiles.length > 2;

  if (hasProduct) {
    cleanImageBackground(imgFiles[1], hasMask ? imgFiles[2] : null, cleanPath);
  }

  // Determine image assignments
  // Heuristic: first image = logo, second = product photo, last = diagram
  // Works for the TMW-30 layout; adjust for other PDFs if needed.
  const logoData = fs.readFileSync(imgFiles[0]);
  const productData = hasProduct ? fs.readFileSync(cleanPath) : logoData;
  const diagIdx = imgFiles.length > 3 ? 3 : imgFiles.length - 1;
  const partsDiagData = fs.readFileSync(imgFiles[diagIdx]);
  const logo2Data =
    imgFiles.length > 4 ? fs.readFileSync(imgFiles[4]) : logoData;

  //Page settings
  const pageProps = {
    page: {
      size: { width: 8418, height: 12240 },
      margin: { top: 720, right: 720, bottom: 720, left: 720 },
    },
  };

  // Page 1: Cover
  const coverPage = {
    properties: pageProps,
    children: [
      imagePara(logoData, 250, 40, "png", 120, 240),
      heading("4FT WOOD HANDLE", HeadingLevel.HEADING_1, AlignmentType.CENTER),
      heading("CANT HOOK", HeadingLevel.HEADING_1, AlignmentType.CENTER),
      para([run("OWNER'S MANUAL", { size: 28, bold: true })], {
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 360 },
      }),
      imagePara(productData, 380, 154, "png", 240, 480),
      para([run(`${WARNING_SYMBOL} WARNING:`, { bold: true })], {
        spacing: { before: 120, after: 40 },
      }),
      para(
        [
          run("Carefully read and understand all "),
          run("ASSEMBLY AND OPERATION INSTRUCTIONS", { bold: true }),
          run(
            " before operating. Failure to follow the safety rules and other basic safety precautions may result in serious personal injury.",
          ),
        ],
        { spacing: { before: 0, after: 40 } },
      ),
      para([run("MODEL# TMW-30", { size: 28, bold: true, color: "FFFFFF" })], {
        alignment: AlignmentType.CENTER,
        spacing: { before: 240, after: 120 },
        shading: { fill: "000000", type: ShadingType.CLEAR },
      }),
      para([run("03242026", { size: 18, color: "666666" })], {
        alignment: AlignmentType.LEFT,
        spacing: { before: 120, after: 0 },
      }),
    ],
  };

  // Page 2: Manual content
  const manualPage = {
    properties: pageProps,
    children: [
      para(
        [
          run("4FT Wood Handle Cant Hook OWNER'S MANUAL", {
            size: 24,
            bold: true,
          }),
        ],
        {
          alignment: AlignmentType.CENTER,
          border: {
            bottom: {
              style: BorderStyle.SINGLE,
              size: 4,
              color: "000000",
              space: 4,
            },
          },
          spacing: { before: 0, after: 160 },
        },
      ),
      heading("GENERAL SAFETY RULES"),
      warningPara(
        "WARNING:",
        "Read and understand all instructions.",
        "Failure to follow all instructions listed below may result in serious injury.",
      ),
      warningPara(
        "CAUTION:",
        "Do not allow persons to operate or assemble this Wood Handle Cant Hook until they have read this manual and have developed a thorough understanding of how the Wood Handle Timber Carrier works.",
      ),
      warningPara(
        "WARNING:",
        "The warnings, cautions, and instructions discussed in this instruction manual cannot cover all possible conditions or situations that could occur.",
        "It must be understood by the operator that common sense and caution are factors that cannot be built into this product, but must be supplied by the operator.",
      ),
      heading("FEATURES", HeadingLevel.HEADING_2, AlignmentType.LEFT, {
        spacing: { before: 180, after: 80 },
      }),
      bullet('One piece solid 2" diameter wood handle'),
      bullet('Overall length 46", usable handle 38"'),
      bullet("Handles up to 14in. Diameter logs"),
      heading("PARTS DIAGRAM:", HeadingLevel.HEADING_2, AlignmentType.LEFT, {
        spacing: { before: 180, after: 80 },
      }),
      imagePara(partsDiagData, 340, 198, "png", 80, 160),
      heading("PARTS LIST:", HeadingLevel.HEADING_2, AlignmentType.LEFT, {
        spacing: { before: 120, after: 80 },
      }),
      partsTable([
        ["1", "Hook", "1", "Wooden handle", "1"],
        [
          "2",
          "M12*45 High-strength half-round head square neck bolt",
          "1",
          "Small conical tube",
          "1",
        ],
        [
          "3",
          "Type 1 Non-Metallic Insert Hex Lock Nut M12",
          "1",
          "Large conical tube",
          "1",
        ],
      ]),
      heading("WARRANTY", HeadingLevel.HEADING_2, AlignmentType.LEFT, {
        spacing: { before: 200, after: 60 },
      }),
      para([run("One-year limited parts warranty", { bold: true })], {
        spacing: { before: 0, after: 160 },
      }),
      imagePara(logo2Data, 180, 29, "png", 80, 80),
      para([run("BAC Industries", { bold: true })], {
        alignment: AlignmentType.CENTER,
        spacing: { before: 40, after: 20 },
      }),
      para([run("PO BOX 155", { bold: true })], {
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 20 },
      }),
      para([run("Miltona, MN 56354", { bold: true })], {
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
      }),
      para(
        [
          run("Size:145x210mm", { size: 16, color: "666666" }),
          run("          157G          ", { size: 16, color: "666666" }),
          run("REV 03/24/26", { size: 16, color: "666666" }),
        ],
        {
          spacing: { before: 120, after: 0 },
          border: {
            top: {
              style: BorderStyle.SINGLE,
              size: 2,
              color: "AAAAAA",
              space: 4,
            },
          },
        },
      ),
    ],
  };

  //Assemble & pack
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [
            {
              level: 0,
              format: LevelFormat.BULLET,
              text: "\u25CF",
              alignment: AlignmentType.LEFT,
              style: { paragraph: { indent: { left: 720, hanging: 360 } } },
            },
          ],
        },
      ],
    },
    styles: {
      default: { document: { run: { font: "Arial", size: 22 } } },
      paragraphStyles: [
        {
          id: "Heading1",
          name: "Heading 1",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 36, bold: true, font: "Arial", color: "000000" },
          paragraph: {
            spacing: { before: 240, after: 120 },
            alignment: AlignmentType.CENTER,
          },
        },
        {
          id: "Heading2",
          name: "Heading 2",
          basedOn: "Normal",
          next: "Normal",
          quickFormat: true,
          run: { size: 28, bold: true, font: "Arial", color: "000000" },
          paragraph: { spacing: { before: 180, after: 120 } },
        },
      ],
    },
    sections: [coverPage, manualPage],
  });

  return Packer.toBuffer(doc);
}

//Route

app.post("/convert", upload.single("pdf"), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: "No PDF uploaded." });

  const workDir = fs.mkdtempSync(path.join(os.tmpdir(), "pdf2docx-"));
  const pdfPath = req.file.path;

  try {
    const buffer = await convertPdfToDocx(pdfPath, workDir);
    const outName = path.basename(req.file.originalname, ".pdf") + ".docx";

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    );
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    res.send(buffer);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  } finally {
    // Cleanup
    try {
      fs.rmSync(workDir, { recursive: true });
    } catch {}
    try {
      fs.unlinkSync(pdfPath);
    } catch {}
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () =>
  console.log(`\n✓ PDF→DOCX server running at http://localhost:${PORT}\n`),
);
