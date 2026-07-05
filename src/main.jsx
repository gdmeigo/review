import React from "react";
import ReactDOM from "react-dom/client";
import "./storagePolyfill.js"; // provides window.storage outside claude.ai — see comments in that file
import App from "./App.jsx";
import "./index.css";

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
