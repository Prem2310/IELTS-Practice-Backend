const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const { PDFDocument, StandardFonts } = require("pdf-lib");
const ExcelJS = require("exceljs");

const app = express();
app.use(cors());
app.use(bodyParser.json());

const PORT = process.env.PORT || 5000; // Default to 5000 if the PORT environment variable is not set

// Route to generate a filled PDF
app.post("/generate-pdf", async (req, res) => {
  const { candidateName, testNumber, testDate, answers } = req.body;

  const pdfDoc = await PDFDocument.create();
  const page = pdfDoc.addPage([600, 800]);
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  page.drawText("IELTS Listening Answer Sheet", {
    x: 200,
    y: 750,
    size: 18,
    font,
  });
  page.drawText(`Candidate Name: ${candidateName || "N/A"}`, {
    x: 50,
    y: 720,
    size: 12,
    font,
  });
  page.drawText(`Test Number: ${testNumber || "N/A"}`, {
    x: 50,
    y: 700,
    size: 12,
    font,
  });
  page.drawText(`Test Date: ${testDate || "N/A"}`, {
    x: 50,
    y: 680,
    size: 12,
    font,
  });

  // Layout for answers (2 columns)
  const leftColumnX = 50;
  const rightColumnX = 300;
  let yPosition = 640;

  answers.forEach((answer, index) => {
    const xPosition = index < 20 ? leftColumnX : rightColumnX;
    if (index === 20) yPosition = 640; // Reset for right column

    page.drawText(`${index + 1}. ${answer || ""}`, {
      x: xPosition,
      y: yPosition,
      size: 12,
      font,
    });
    yPosition -= 20;
  });

  const pdfBytes = await pdfDoc.save();

  // Dynamic file name
  const fileName = `${candidateName || "Candidate"}_${
    testNumber || "Test"
  }.pdf`;
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", `attachment; filename=${fileName}`);
  res.send(Buffer.from(pdfBytes));
});

// Route to generate a filled Excel file
app.post("/generate-excel", async (req, res) => {
  try {
    const { candidateName, candidateNumber, testDate, answers } = req.body;

    // Create a new workbook and worksheet
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("IELTS Answers");

    // Add headers
    worksheet.addRow(["Candidate Name", candidateName || "N/A"]);
    worksheet.addRow(["Candidate Number", candidateNumber || "N/A"]);
    worksheet.addRow(["Test Date", testDate || "N/A"]);
    worksheet.addRow([]);
    worksheet.addRow(["Question", "Answer"]);

    // Add answers
    answers.forEach((answer, index) => {
      worksheet.addRow([index + 1, answer || ""]);
    });

    // Write the Excel file to a buffer
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", "attachment; filename=answers.xlsx");
    res.send(buffer);
  } catch (err) {
    res.status(500).send("Error generating Excel: " + err.message);
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
