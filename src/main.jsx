import React from "react";
import ReactDOM from "react-dom/client";
import installLocalStorageShim from "./lib/storage.js";
import App from "./App.jsx";
import "./index.css";

installLocalStorageShim();

ReactDOM.createRoot(document.getElementById("root")).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
