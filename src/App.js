import React, { useState } from 'react';
import axios from 'axios';

const App = () => {
  const [pptFile, setPptFile] = useState(null);
  const [excelFile, setExcelFile] = useState(null);

  const handlePptChange = (e) => {
    setPptFile(e.target.files[0]);
  };

  const handleExcelChange = (e) => {
    setExcelFile(e.target.files[0]);
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    const formData = new FormData();
    formData.append('ppt', pptFile);
    formData.append('excel', excelFile);

    try {
      const response = await axios.post('http://localhost:5000/upload', formData, {
        responseType: 'blob', // Important for downloading files
      });
      const url = window.URL.createObjectURL(new Blob([response.data]));
      const link = document.createElement('a');
      link.href = url;
      link.setAttribute('download', 'certificates.zip');
      document.body.appendChild(link);
      link.click();
    } catch (error) {
      console.error('Error uploading files', error);
    }
  };

  return (
    <div>
      <h1>Upload PPT and Excel Files</h1>
      <form onSubmit={handleSubmit}>
        <div>
          <label>PPT File:</label>
          <input type="file" onChange={handlePptChange} accept=".pptx" />
        </div>
        <div>
          <label>Excel File:</label>
          <input type="file" onChange={handleExcelChange} accept=".xlsx" />
        </div>
        <button type="submit">Submit</button>
      </form>
    </div>
  );
};

export default App;
