import React, { useEffect, useState } from 'react';
import ExcelJS from 'exceljs';
import FortuneSheet from '@fortune-sheet/react';
import '@fortune-sheet/react/dist/index.css';
import excelFile from './spreadsheet.xlsx'; // Import the Excel file

function ExcelRenderer() {
  const [data, setData] = useState([]);

  useEffect(() => {
    const fetchAndParseExcel = async () => {
      try {
        // Fetch the Excel file
        const response = await fetch(excelFile);
        const arrayBuffer = await response.arrayBuffer();
        
        // Read from the buffer
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(arrayBuffer);

        const worksheet = workbook.getWorksheet(1);
        
        // Process the worksheet data
        const sheetData = [];
        worksheet.eachRow((row) => {
          const rowData = row.values.map(cell => {
            if (cell) {
              return {
                v: cell.text || cell.toString(),
                s: cell.style,
              };
            }
            return { v: '', s: {} };
          });
          sheetData.push(rowData);
        });

        // Set the data for FortuneSheet
        setData([{ name: 'Sheet1', rows: sheetData }]);
      } catch (error) {
        console.error('Error fetching and parsing Excel file:', error);
      }
    };

    fetchAndParseExcel();
  }, []);

  const saveToFile = async () => {
    try {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Sheet1');
      
      // Fill workbook with existing data
      data[0].rows.forEach((row, rowIndex) => {
        const worksheetRow = worksheet.getRow(rowIndex + 1);
        row.forEach((cell, colIndex) => {
          const worksheetCell = worksheetRow.getCell(colIndex + 1);
          worksheetCell.value = cell.v;
          Object.assign(worksheetCell.style, cell.s);
        });
      });
      
      await workbook.xlsx.writeFile('output.xlsx');
      console.log('File saved successfully');
    } catch (error) {
      console.error('Error saving file:', error);
    }
  };

  return (
    <div className="container mx-auto p-4">
      <h1 className="text-2xl mb-4">Excel Data with Fortune Sheets</h1>
      {data.length > 0 ? (
        <div>
          <FortuneSheet data={data} />
          <button onClick={saveToFile} className="mt-4 p-2 bg-blue-500 text-white">Save to File</button>
        </div>
      ) : (
        <p>Loading...</p>
      )}
    </div>
  );
}

export default ExcelRenderer;
