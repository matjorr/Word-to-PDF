
const $ = (selector) => document.querySelector(selector);

// Your live LibreOffice conversion service
const CONVERSION_API_URL =
  "https://word-to-pdf-api-xm4z.onrender.com/convert";

const fileInput = $("#fileInput");
const dropZone = $("#dropZone");
const convertButton = $("#convertButton");
const selectedFile = $("#selectedFile");
const fileName = $("#fileName");
const errorMessage = $("#errorMessage");
const progressArea = $("#progressArea");
const progressBar = $("#progressBar");
const progressText = $("#progressText");
const removeFileButton = $("#removeFile");

let currentFile = null;
let isConverting = false;

if ($("#year")) {
  $("#year").textContent = new Date().getFullYear();
}

function showError(message) {
  errorMessage.textContent = message;
}

function clearError() {
  errorMessage.textContent = "";
}

function updateProgress(percent, message) {
  progressArea.hidden = false;
  progressBar.style.width = `${percent}%`;
  progressText.textContent = message;
}

function updateFileUI() {
  const hasFile = Boolean(currentFile);

  selectedFile.hidden = !hasFile;
  fileName.textContent = hasFile ? currentFile.name : "";

  convertButton.disabled = !hasFile || isConverting;
}

function setFile(file) {
  clearError();

  if (!file) return;

  if (!file.name.toLowerCase().endsWith(".docx")) {
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

fileInput.addEventListener("change", () => {
  setFile(fileInput.files?.[0]);
});

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

removeFileButton.addEventListener("click", () => {
  if (isConverting) return;

  currentFile = null;
  fileInput.value = "";

  clearError();
  updateFileUI();
});

async function convertDocument() {
  if (!currentFile || isConverting) return;

  isConverting = true;
  clearError();
  updateFileUI();

  try {
    updateProgress(10, "Preparing your original Word document…");

    const formData = new FormData();
    formData.append("file", currentFile, currentFile.name);

    updateProgress(25, "Connecting to the conversion service…");

    const response = await fetch(CONVERSION_API_URL, {
      method: "POST",
      body: formData
    });

    if (!response.ok) {
      let message = "Conversion failed. Please try again.";

      try {
        const result = await response.json();
        if (result.detail) message = result.detail;
      } catch (_) {}

      throw new Error(message);
    }

    const contentType =
      response.headers.get("content-type") || "";

    if (!contentType.includes("application/pdf")) {
      throw new Error("The conversion service did not return a valid PDF.");
    }

    updateProgress(80, "Preparing your PDF download…");

    const pdfBlob = await response.blob();

    if (!pdfBlob.size) {
      throw new Error("The generated PDF is empty. Please try again.");
    }

    const originalName =
      currentFile.name.replace(/\.docx$/i, "");

    const downloadUrl = URL.createObjectURL(pdfBlob);
    const link = document.createElement("a");

    link.href = downloadUrl;
    link.download = `${originalName}.pdf`;

    document.body.appendChild(link);
    link.click();
    link.remove();

    setTimeout(() => URL.revokeObjectURL(downloadUrl), 10000);

    updateProgress(100, "Your PDF has been downloaded successfully!");

  } catch (error) {
    console.error("Conversion error:", error);

    showError(
      error.message ||
      "Something went wrong. Please try again."
    );

    progressText.textContent = "Conversion unsuccessful.";

  } finally {
    isConverting = false;
    updateFileUI();

    setTimeout(() => {
      progressArea.hidden = true;
      progressBar.style.width = "0%";
    }, 5000);
  }
}

convertButton.addEventListener("click", convertDocument);
