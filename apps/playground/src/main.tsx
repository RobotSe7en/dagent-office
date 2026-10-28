import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import "@dagent-office/core";
import "@dagent-office/editor";
import "@dagent-office/office";
import { App } from "./App.js";
import "./styles.css";

const container = document.getElementById("root");
if (!container) throw new Error("Playground root element is missing.");

createRoot(container).render(
  <StrictMode>
    <App />
  </StrictMode>,
);
