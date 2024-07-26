import React, { useState } from 'react';
import * as XLSX from 'xlsx'; // Import all XLSX functions
import PptxGenJS from 'pptxgenjs';
import JSZip from 'jszip';
import { saveAs } from 'file-saver';

const FileUpload = () => {
  const [excelData, setExcelData] = useState(null);
  const [templateFile, setTemplateFile] = useState(null);

  const handleExcelUpload = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (event) => {
      const data = new Uint8Array(event.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet);
      setExcelData(jsonData);
    };
    reader.readAsArrayBuffer(file);
  };

  const handleTemplateUpload = (e) => {
    const file = e.target.files[0];
    setTemplateFile(file);
  };

  const generateCertificates = async () => {
    if (!excelData || !templateFile) {
      alert("Please upload both the Excel file and the PowerPoint template.");
      return;
    }

    const zip = new JSZip();
    for (let i = 0; i < excelData.length; i++) {
      const row = excelData[i];
      const presentation = new PptxGenJS();
      const slide = presentation.addSlide();
      slide.addText(`Student Name: ${row['Stud_name']}`, { x: 1, y: 1, fontSize: 24, bold: true });
      slide.addText(`Course Name: ${row['course_name']}`, { x: 1, y: 2, fontSize: 24, bold: true });
      // Add other text placeholders similarly

      const buffer = await presentation.write('arraybuffer');
      zip.file(`Certificate_${row['Stud_name']}.pptx`, buffer);
    }

    zip.generateAsync({ type: 'blob' }).then((content) => {
      saveAs(content, 'certificates.zip');
    });
  };

  return (
    <div>
      <input type="file" accept=".xlsx" onChange={handleExcelUpload} />
      <input type="file" accept=".pptx" onChange={handleTemplateUpload} />
      <button onClick={generateCertificates}>Generate Certificates</button>
    </div>
  );
};

export default FileUpload;
