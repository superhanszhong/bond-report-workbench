import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import Workbench from "../app/workbench";
import "../app/globals.css";

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <Workbench />
  </StrictMode>,
);
