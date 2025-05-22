import mammoth from "mammoth";
import puppeteer from "puppeteer";
import { writeFile, readFile } from "fs/promises";
import { join } from "path";
import { tmpdir } from "os";
import { randomUUID } from "crypto";

async function convertDocxToPdf(docxBuffer) {
  try {
    const { value: html } = await mammoth.convertToHtml(
      { buffer: docxBuffer },
      {
        styleMap: [
          "p[style-name='Heading 1'] => h1:fresh",
          "p[style-name='Heading 2'] => h2:fresh",
          "p[style-name='Heading 3'] => h3:fresh",
          "p[style-name='Title'] => h1.title:fresh",
          "table => table.docx-table",
          "r[style-name='Strong'] => strong",
        ],
      }
    );

    const fullHtml = `
      <!DOCTYPE html>
      <html>
        <head>
          <meta charset="utf-8">
          <style>
            body {
              font-family: Arial, sans-serif;
              line-height: 1.5;
              margin: 40px;
            }
            table.docx-table {
              border-collapse: collapse;
              width: 100%;
              margin: 15px 0;
            }
            table.docx-table td, table.docx-table th {
              border: 1px solid #ddd;
              padding: 8px;
            }
            h1, h2, h3, h4, h5, h6 {
              margin-top: 20px;
              margin-bottom: 10px;
              font-weight: bold;
            }
            h1 { font-size: 24pt; }
            h2 { font-size: 18pt; }
            h3 { font-size: 14pt; }
            p { margin: 10px 0; }
            img { max-width: 100%; }
            @page {
              margin: 1cm;
            }
          </style>
        </head>
        <body>
          ${html}
        </body>
      </html>
    `;

    const tempDir = tmpdir();
    const tempHtmlPath = join(tempDir, `${randomUUID()}.html`);
    await writeFile(tempHtmlPath, fullHtml);

    const browser = await puppeteer.launch({
      headless: "new",
      args: ["--no-sandbox", "--disable-setuid-sandbox"],
    });
    const page = await browser.newPage();
    await page.goto(`file://${tempHtmlPath}`, { waitUntil: "networkidle0" });

    const pdfBuffer = await page.pdf({
      format: "A4",
      printBackground: true,
      margin: {
        top: "1cm",
        right: "1cm",
        bottom: "1cm",
        left: "1cm",
      },
    });

    await browser.close();

    return pdfBuffer;
  } catch (error) {
    console.error("Error converting DOCX to PDF:", error);
    throw error;
  }
}

async function main() {
  try {
    const samplePath = "./sample.docx";
    console.log(`Reading sample DOCX from ${samplePath}`);

    try {
      const docxBuffer = await readFile(samplePath);
      console.log("DOCX file loaded successfully");

      console.log("Converting DOCX to PDF...");
      const pdfBuffer = await convertDocxToPdf(docxBuffer);
      console.log(`PDF generated successfully (${pdfBuffer.length} bytes)`);

      const outputPath = "./output.pdf";
      await writeFile(outputPath, pdfBuffer);
      console.log(`PDF saved to ${outputPath}`);

      return pdfBuffer;
    } catch (readError) {
      console.error("Error reading sample file:", readError.message);
      console.log("Using a demo function instead...");
    }
  } catch (error) {
    console.error("Error in main function:", error);
  }
}

main();
