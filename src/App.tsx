import reactLogo from "./assets/react.svg";
import viteLogo from "/vite.svg";
import "./App.css";
import { ExcelReader } from "./components";

function App() {
  return (
    <div className="app-container">
      <header className="app-header">
        <div className="logo-container">
          <a href="https://vite.dev" target="_blank" rel="noreferrer">
            <img src={viteLogo} className="logo" alt="Vite logo" />
          </a>
          <a href="https://react.dev" target="_blank" rel="noreferrer">
            <img src={reactLogo} className="logo react" alt="React logo" />
          </a>
        </div>
        <h1>Excel Style Parser</h1>
        <p>Upload an Excel file to extract and view cell styles</p>
      </header>

      <div className="excel-reader-table-container">
        <ExcelReader />
      </div>

      <p className="read-the-docs">
        Built with Vite, React, and ExcelJS library
      </p>
    </div>
  );
}

export default App;
