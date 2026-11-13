
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
let isConverting = false;

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

  const isDocx =
    file.name.toLowerCase().endsWith(".docx") ||
    file.type ===
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

  if (!isDocx) {
    currentFile = null;
    showError("Please select a Word .docx file. Older .doc files are not supported.");
    updateFileUI();
    return;
  }

  if (file.size > 15 * 1024 * 1024) {
    currentFile = null;
    showError("Your file is larger than 15 MB. Please choose a smaller document.");
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
  convertButton.disabled = !hasFile || isConverting;
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
  if (isConverting) return;

  currentFile = null;
  fileInput.value = "";
  clearError();
  updateFileUI();
});

function updateProgress(percent, message) {
  progressArea.hidden = false;
  progressBar.style.width = `${percent}%`;
  progressText.textContent = message;
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
    images.map((image) => {
      if (image.complete) return Promise.resolve();

      return new Promise((resolve) => {
        image.onload = resolve;
        image.onerror = resolve;
      });
    })
  );
}

function getLibraries() {
  const docxLibrary = window.docx;
  const pdfLibrary = window.html2pdf;

  if (
    !docxLibrary ||
    typeof docxLibrary.renderAsync !== "function"
  ) {
    throw new Error("The Word document rendering library did not load.");
  }

  if (typeof pdfLibrary !== "function") {
    throw new Error("The PDF conversion library did not load.");
  }

  return {
    docxLibrary,
    pdfLibrary
  };
}

async function convertDocument() {
  if (!currentFile || isConverting) return;

  clearError();
  isConverting = true;
  updateFileUI();

  const originalStageStyle = renderStage.getAttribute("style");

  try {
    const { docxLibrary, pdfLibrary } = getLibraries();

    renderStage.innerHTML = "";

    // Put the document inside the browser viewport.
    // This avoids the blank PDF issue caused by capturing
    // content positioned far outside the viewport.

    renderStage.style.position = "fixed";
    renderStage.style.left = "0";
    renderStage.style.top = "0";
    renderStage.style.width = "794px";
    renderStage.style.minHeight = "100px";
    renderStage.style.height = "auto";
    renderStage.style.overflow = "visible";
    renderStage.style.background = "#ffffff";
    renderStage.style.color = "#111111";
    renderStage.style.zIndex = "9999";
    renderStage.style.padding = "0";
    renderStage.style.margin = "0";

    updateProgress(10, "Reading your Word document…");

    const buffer = await currentFile.arrayBuffer();

    updateProgress(30, "Rendering document contents…");

    await docxLibrary.renderAsync(
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

    // Check that the document actually rendered.
    const renderedContent = renderStage.innerText.trim();
    const renderedHeight = renderStage.scrollHeight;

    if (!renderedContent || renderedHeight < 50) {
      throw new Error(
        "The Word document appears empty after rendering."
      );
    }

    updateProgress(60, "Preparing your PDF pages…");

    const safeName = currentFile.name.replace(/\.docx$/i, "");

    const options = {
      margin: [10, 10, 10, 10],
      filename: `${safeName}.pdf`,

      image: {
        type: "jpeg",
        quality: 0.98
      },

      html2canvas: {
        scale: 2,
        useCORS: true,
        backgroundColor: "#ffffff",
        logging: false,
        scrollX: 0,
        scrollY: 0,
        windowWidth: 794
      },

      jsPDF: {
        unit: "mm",
        format: "a4",
        orientation: "portrait",
        compress: true
      },

      pagebreak: {
        mode: ["css", "legacy"],
        avoid: ["tr", "img"]
      }
    };

    updateProgress(75, "Generating your PDF…");

    await waitForLayout();

    await pdfLibrary()
      .set(options)
      .from(renderStage)
      .save();

    updateProgress(100, "Your PDF has been downloaded!");

  } catch (error) {
    console.error("Conversion error:", error);

    showError(
      error.message ||
      "Conversion failed. Please try another DOCX document."
    );

    progressText.textContent = "Conversion unsuccessful.";

  } finally {
    // Restore the original hidden rendering stage.
    if (originalStageStyle === null) {
      renderStage.removeAttribute("style");
    } else {
      renderStage.setAttribute("style", originalStageStyle);
    }

    renderStage.innerHTML = "";

    isConverting = false;
    updateFileUI();

    setTimeout(() => {
      progressArea.hidden = true;
      progressBar.style.width = "0%";
    }, 5000);
  }
}

convertButton.addEventListener("click", convertDocument);
