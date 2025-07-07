import { BrowserRouter as Router, Routes, Route } from "react-router-dom";
import "bootstrap/dist/css/bootstrap.min.css";
import MultiSheetBuilder from "./pages/SerialExtractor";
import SecondExtractor from "./pages/NewFormatEtractor";
import Navbar from "./pages/Navbar";
import ThirdExtractor from "./pages/ThirdExtractor";
export default function App() {
  return (
    <Router>
      <Navbar />
      <div className="container py-4">
        <Routes>
          <Route path="/first" element={<MultiSheetBuilder />} />
          <Route path="/second" element={<SecondExtractor />} />
          <Route path="/third" element={<ThirdExtractor />} />
          <Route
            path="/"
            element={
              <div className="text-center">
                <h2 className="mb-4">Welcome to Excel Formatter</h2>
                <p>Select a tool from the navigation bar above to begin.</p>
              </div>
            }
          />
        </Routes>
      </div>
    </Router>
  );
}
