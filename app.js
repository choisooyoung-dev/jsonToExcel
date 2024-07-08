const express = require("express");
const xl = require("excel4node");
const fs = require("fs");
const path = require("path");

const app = express();
const port = 3000;

// JSON 파일 경로
const jsonFilePath = path.join(__dirname, "data.json");

// JSON 데이터를 읽어오는 함수
function readJsonData() {
  const rawData = fs.readFileSync(jsonFilePath);
  return JSON.parse(rawData);
}

// 엑셀 파일 생성 함수
function createExcelFile(jsonData, filePath) {
  const wb = new xl.Workbook();

  jsonData.forEach((sheetData, index) => {
    const ws = wb.addWorksheet(`Sheet${index + 1}`);

    const headers = Object.keys(sheetData[0]);

    headers.forEach((header, colIndex) => {
      ws.cell(1, colIndex + 1).string(header);
    });

    sheetData.forEach((row, rowIndex) => {
      headers.forEach((header, colIndex) => {
        ws.cell(rowIndex + 2, colIndex + 1).string(row[header].toString());
      });
    });
  });

  // 엑셀 파일 저장
  return new Promise((resolve, reject) => {
    wb.write(filePath, (err, stats) => {
      if (err) {
        reject(err);
      } else {
        resolve(stats);
      }
    });
  });
}

// 엑셀 파일 다운로드 엔드포인트
app.get("/download-excel", async (req, res) => {
  const jsonData = readJsonData();
  const filePath = path.join(__dirname, "output.xlsx");

  try {
    await createExcelFile(jsonData, filePath);

    // 파일이 생성되었는지 확인
    if (fs.existsSync(filePath)) {
      res.download(filePath, "output.xlsx", (err) => {
        if (err) {
          console.error("Error downloading file:", err);
          res.status(500).send("Error downloading file");
        }
      });
    } else {
      console.error("File not found:", filePath);
      res.status(404).send("File not found");
    }
  } catch (err) {
    console.error("Error creating file:", err);
    res.status(500).send("Error creating file");
  }
});

// HTML 페이지 제공
app.get("/", (req, res) => {
  res.send(`
    <!DOCTYPE html>
    <html>
      <body>
        <button onclick="window.location.href='/download-excel'">Download Excel</button>
      </body>
    </html>
  `);
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
