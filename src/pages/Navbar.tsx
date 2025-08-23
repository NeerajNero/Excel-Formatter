import { Link } from "react-router-dom";

export default function Navbar() {
  return (
    <nav className="navbar navbar-expand-lg navbar-dark bg-dark">
      <div className="container-fluid">
        <Link to='/' className="navbar-brand">🧾 Excel Formatter</Link>
        <button
          className="navbar-toggler"
          type="button"
          data-bs-toggle="collapse"
          data-bs-target="#navbarNav"
        >
          <span className="navbar-toggler-icon"></span>
        </button>

        <div className="collapse navbar-collapse" id="navbarNav">
          <ul className="navbar-nav ms-auto">
            <li className="nav-item">
              <Link to="/first" className="nav-link">
                First Excel Extractor
              </Link>
            </li>
            <li className="nav-item">
              <Link to="/second" className="nav-link">
                Second Excel Extractor
              </Link>
            </li>
            <li className="nav-item">
              <Link to="/third" className="nav-link">
                Third Excel Extractor
              </Link>
            </li>
            <li className="nav-item">
              <Link to="/fourth" className="nav-link">
                Fourth Excel Extractor
              </Link>
            </li>
          </ul>
        </div>
      </div>
    </nav>
  );
}
