const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const port = 3020;

// Konfigurasi multer untuk upload file
const upload = multer({ dest: 'uploads/' });

// Mapping kolom dengan alias
const columnMappings = [
  { key: 'nopol', aliases: ['licenseplate', 'nopolisi', 'nopol', 'plate', 'vehicleplate'], required: true },
  { key: 'mobil', aliases: ['unit', 'assettype', 'merk', 'type', 'jeniskendaraan', 'mobil', 'jenis', 'typeunit', 'jeniskendaraan', 'vehicle'] },
  { key: 'lesing', aliases: ['lesing', 'leasing', 'lesng', 'finance', 'financing'] },
  { key: 'ovd', aliases: ['overdue', 'ovd', 'daysoverdue', 'overdu', 'hari', 'keterlambatan', 'dayslate'] },
  { key: 'saldo', aliases: ['saldo', 'credit', 'balance', 'amount', 'remaining'] },
  { key: 'cabang', aliases: ['branchfullname', 'cabang', 'branch', 'office', 'location'] },
  { key: 'nama', aliases: ['nama', 'name', 'fullname', 'customername', 'owner'] },
  { key: 'noka', aliases: ['chasisno', 'nomorrangka', 'norangka', 'no.rangka', 'noka', 'chassis', 'frame'] },
  { key: 'nosin', aliases: ['nomesin', 'nomormesin', 'no.mesin', 'nosin', 'engine', 'engineno'] },
];

// Urutan output yang diinginkan
const outputOrder = ['nopol', 'mobil', 'lesing', 'ovd', 'saldo', 'cabang', 'nama', 'noka', 'nosin'];

function processExcelFile(filePath) {
  try {
    const workbook = XLSX.readFile(filePath, { cellDates: true });
    const result = [];

    workbook.SheetNames.forEach(sheetName => {
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: false });

      if (rows.length < 5) return;

      // Cari header dalam 10 baris pertama
      let headerMap = new Map();
      let headerRowIdx = -1;

      for (let i = 0; i < Math.min(10, rows.length); i++) {
        const row = rows[i];
        row.forEach((cell, colIdx) => {
          if (!cell || typeof cell !== 'string') return;
          const cleanedCell = cell.toString().toLowerCase().replace(/\s/g, '');
          columnMappings.forEach(mapping => {
            if (mapping.aliases.some(alias => cleanedCell.includes(alias))) {
              headerMap.set(mapping.key, colIdx);
              headerRowIdx = i;
            }
          });
        });
        if (headerMap.size > 0) break;
      }

      if (headerRowIdx === -1) return;

      // Proses data rows
      for (let i = headerRowIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        const record = {};

        headerMap.forEach((colIdx, key) => {
          if (colIdx < row.length && row[colIdx] !== undefined && row[colIdx] !== '') {
            let value = row[colIdx].toString().trim().replace(/,/g, '.');
            switch (key) {
              case 'saldo':
                const num = parseFloat(value);
                record[key] = isNaN(num) ? value : Math.round(num).toString();
                break;
              case 'nopol':
                record[key] = value.replace(/\s/g, '');
                break;
              default:
                record[key] = value;
            }
          }
        });

        // Hanya masukkan record yang memiliki nopol (required)
        if (record['nopol']) {
          const rowData = outputOrder.map(key => record[key] || '');
          result.push(rowData);
        }
      }
    });

    return result;
  } catch (error) {
    throw new Error(`Error processing Excel file: ${error.message}`);
  }
}

// Endpoint API untuk upload file
app.post('/upload', upload.single('excelFile'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file uploaded' });
  }

  const filePath = req.file.path;

  try {
    const data = processExcelFile(filePath);
    if (data.length === 0) {
      return res.status(400).json({ error: 'No valid data found in the file' });
    }

    // Hapus file setelah diproses
    fs.unlink(filePath, err => {
      if (err) console.error(`Failed to delete file: ${err}`);
    });

    res.json(data);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Jalankan server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});