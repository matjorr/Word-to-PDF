
const $ = (selector) => document.querySelector(selector);

const fileInput = $("#fileInput");
const dropZone = $("#dropZone");
const convertButton = $("#convertButton");
const selectedFile = $("#selectedFile");
const fileName = $("#fileName");
const errorMessage = $("#errorMessage");
const progressArea = $("#progressArea");
const progressBar = $("#progressBar");
const progressText = $("#progressText");
const renderStage = $("#renderStage");
const removeFileButton = $("#removeFile");

let currentFile = null;
let isConverting = false;

if ($("#year")) {
  $("#year").textContent = new Date().getFullYear();
}

function showError(message) {
  if (errorMessage) errorMessage.textContent = message;
}

function clearError() {
  if (errorMessage) errorMessage.textContent = "";
}

function updateFileUI() {
  const hasFile = Boolean(currentFile);

  if (selectedFile) selectedFile.hidden = !hasFile;
  if (fileName) fileName.textContent = hasFile ? currentFile.name : "";

  if (convertButton) {
    convertButton.disabled = !hasFile || isConverting;
  }
}

function setProgress(percent, message) {
  if (progressArea) progressArea.hidden = false;
  if (progressBar) progressBar.style.width = `${percent}%`;
  if (progressText) progressText.textContent = message;
}

function setFile(file) {
  clearError();

  if (!file) return;

  const isDocx = file.name.toLowerCase().endsWith(".docx");

  if (!isDocx) {
    currentFile = null;
    showError("Please select a Word .docx file. Older .doc files are not supported.");
    updateFileUI();
    return;
  }

  if (file.size > 15 * 1024 * 1024) {
    currentFile = null;
    showError("Your document exceeds the 15 MB limit.");
    updateFileUI();
    return;
  }

  currentFile = file;
  updateFileUI();
}

if (fileInput) {
  fileInput.addEventListener("change", () => {
    setFile(fileInput.files?.[0]);
  });
}

if (dropZone) {
  ["dragenter", "dragover"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropZone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropZone.classList.remove("dragover");
    });
  });

  dropZone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) setFile(file);
  });
}

if (removeFileButton) {
  removeFileButton.addEventListener("click", () => {
    if (isConverting) return;

    currentFile = null;
    if (fileInput) fileInput.value = "";

    clearError();
    updateFileUI();
  });
}

// Add print-specific styling.
// The website interface stays hidden when the print dialog opens.

function installPrintStyles() {
  let style = document.getElementById("document-print-styles");

  if (!style) {
    style = document.createElement("style");
    style.id = "document-print-styles";

    style.textContent = `
      @media print {
        @page {
          margin: 12mm;
        }

        html, body {
          background: #ffffff !important;
          margin: 0 !important;
          padding: 0 !important;
          height: auto !important;
          overflow: visible !important;
        }

        body * {
          visibility: hidden !important;
        }

        #renderStage,
        #renderStage * {
          visibility: visible !important;
        }

        #renderStage {
          display: block !important;
          position: static !important;
          left: auto !important;
          top: auto !important;
          width: 100% !important;
          height: auto !important;
          min-height: 0 !important;
          overflow: visible !important;
          background: #ffffff !important;
          color: #000000 !important;
          z-index: auto !important;
          margin: 0 !important;
          padding: 0 !important;
        }

        #renderStage .docx-wrapper {
          display: block !important;
          background: #ffffff !important;
          padding: 0 !important;
          margin: 0 !important;
        }

        #renderStage .docx {
          box-shadow: none !important;
          margin: 0 auto !important;
          break-after: page;
          page-break-after: always;
        }

        #renderStage .docx:last-child {
          break-after: auto;
          page-break-after: auto;
        }
      }
    `;

    document.head.appendChild(style);
  }

  return style;
}

function waitForLayout() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => {
      requestAnimationFrame(resolve);
    });
  });
}

async function waitForImages(container) {
  const images = Array.from(container.querySelectorAll("img"));

  await Promise.all(
    images.map((img) => {
      if (img.complete) return Promise.resolve();

      return new Promise((resolve) => {
        img.onload = resolve;
        img.onerror = resolve;
      });
    })
  );
}

async function convertDocument() {
  if (!currentFile || isConverting) return;

  clearError();
  isConverting = true;
  updateFileUI();

  const originalStyle = renderStage.getAttribute("style");

  try {
    if (
      !window.docx ||
      typeof window.docx.renderAsync !== "function"
    ) {
      throw new Error(
        "The Word rendering library could not load. Check your internet connection and refresh the page."
      );
    }

    renderStage.innerHTML = "";

    // Render the document outside the visible website.
    renderStage.style.position = "fixed";
    renderStage.style.left = "-100000px";
    renderStage.style.top = "0";
    renderStage.style.width = "794px";
    renderStage.style.height = "auto";
    renderStage.style.overflow = "visible";
    renderStage.style.background = "#ffffff";
    renderStage.style.color = "#000000";
    renderStage.style.zIndex = "-1";

    setProgress(10, "Reading your Word document…");

    const buffer = await currentFile.arrayBuffer();

    setProgress(30, "Rendering Word pages and images…");

    await window.docx.renderAsync(
      buffer,
      renderStage,
      null,
      {
        className: "docx",
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        ignoreFonts: false,
        breakPages: true,
        useBase64URL: true
      }
    );

    await waitForImages(renderStage);
    await waitForLayout();

    const pageElements = renderStage.querySelectorAll(".docx");
    const textContent = renderStage.innerText.trim();

    if (!pageElements.length || !textContent) {
      throw new Error(
        "The Word document could not be rendered. Please try another DOCX file."
      );
    }

    setProgress(75, "Preparing your document for printing…");

    installPrintStyles();

    // Move the rendered document into the printable page.
    renderStage.style.position = "static";
    renderStage.style.left = "auto";
    renderStage.style.top = "auto";
    renderStage.style.width = "100%";
    renderStage.style.height = "auto";
    renderStage.style.minHeight = "0";
    renderStage.style.overflow = "visible";
    renderStage.style.background = "#ffffff";
    renderStage.style.color = "#000000";
    renderStage.style.zIndex = "auto";
    renderStage.style.margin = "0";
    renderStage.style.padding = "0";

    await waitForLayout();

    setProgress(
      100,
      "Your document is ready. Choose Save as PDF in the print window."
    );

    // Open the browser's native print dialog.
    window.print();

  } catch (error) {
    console.error("Word to PDF error:", error);

    showError(
      error.message ||
      "Something went wrong while preparing your document."
    );

  } finally {
    // Restore the original hidden rendering area.
    if (originalStyle === null) {
      renderStage.removeAttribute("style");
    } else {
      renderStage.setAttribute("style", originalStyle);
    }

    renderStage.innerHTML = "";

    isConverting = false;
    updateFileUI();

    setTimeout(() => {
      if (progressArea) progressArea.hidden = true;
      if (progressBar) progressBar.style.width = "0%";
    }, 5000);
  }
}

if (convertButton) {
  convertButton.addEventListener("click", convertDocument);
}
