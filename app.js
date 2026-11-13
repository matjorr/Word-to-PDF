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

let currentFile = null;

$("#year").textContent = new Date().getFullYear();

function showError(message) {
  errorMessage.textContent = message;
}

function clearError() {
  errorMessage.textContent = "";
}

function setFile(file) {
  clearError();

  if (!file) return;

  const isDocx = file.name.toLowerCase().endsWith(".docx") ||
    file.type === "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  if (!isDocx) {
    currentFile = null;
    showError("Please choose a Microsoft Word .docx file. Older .doc files are not supported.");
    updateFileUI();
    return;
  }

  if (file.size > 15 * 1024 * 1024) {
    currentFile = null;
    showError("This file is larger than 15 MB. Please choose a smaller document.");
    updateFileUI();
    return;
  }

  currentFile = file;
  updateFileUI();
}

function updateFileUI() {
  const hasFile = Boolean(currentFile);
  selectedFile.hidden = !hasFile;
  fileName.textContent = hasFile ? currentFile.name : "";
  convertButton.disabled = !hasFile;
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

$("#removeFile").addEventListener("click", () => {
  currentFile = null;
  fileInput.value = "";
  clearError();
  updateFileUI();
});

function waitForImages(container) {
  const images = Array.from(container.querySelectorAll("img"));
  return Promise.all(images.map((img) => {
    if (img.complete) return Promise.resolve();
    return new Promise((resolve) => {
      img.onload = resolve;
      img.onerror = resolve;
    });
  }));
}

convertButton.addEventListener("click", async () => {
  if (!currentFile) return;

  if (typeof window.docx !== "object" || typeof window.html2pdf !== "function") {
    showError("The conversion tools could not load. Check your internet connection and refresh the page.");
    return;
  }

  clearError();
  convertButton.disabled = true;
  progressArea.hidden = false;
  progressBar.style.width = "15%";
  progressText.textContent = "Reading your Word document…";
  renderStage.innerHTML = "";

  try {
    const buffer = await currentFile.arrayBuffer();

    progressBar.style.width = "35%";
    progressText.textContent = "Preparing document preview…";

    await window.docx.renderAsync(buffer, renderStage, null, {
      className: "docx",
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      ignoreFonts: false,
      breakPages: true,
      useBase64URL: true
    });

    await waitForImages(renderStage);

    progressBar.style.width = "75%";
    progressText.textContent = "Creating your PDF…";

    const safeName = currentFile.name.replace(/\.docx$/i, "");
    const options = {
      margin: [12, 12, 12, 12],
      filename: `${safeName}.pdf`,
      image: { type: "jpeg", quality: 0.98 },
      html2canvas: {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
        logging: false
      },
      jsPDF: {
        unit: "mm",
        format: "a4",
        orientation: "portrait"
      },
      pagebreak: { mode: ["css", "legacy"] }
    };

    await window.html2pdf().set(options).from(renderStage).save();

    progressBar.style.width = "100%";
    progressText.textContent = "Your PDF download should begin shortly.";
  } catch (error) {
    console.error("DOCX to PDF conversion failed:", error);
    showError("Sorry, this document could not be converted. Try a simpler DOCX file or check whether it contains unsupported elements.");
    progressText.textContent = "Conversion unsuccessful.";
  } finally {
    convertButton.disabled = !currentFile;
    setTimeout(() => {
      progressArea.hidden = true;
      progressBar.style.width = "0%";
    }, 4500);
  }
});
