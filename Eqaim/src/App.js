import React from 'react';
import ExcelRenderer from './ExcelRenderer';
import './App.css'; // Add any global styles here

function App() {
  return (
    <div className="App">
      <h1>Excel Data with Fortune Sheets</h1>
      <ExcelRenderer />
    </div>
  );
}

export default App;