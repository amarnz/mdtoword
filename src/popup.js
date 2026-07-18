import { markdownToDocx } from "./markdownToDocx.js";

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const fileInfo = document.getElementById("fileInfo");
const convertBtn = document.getElementById("convertBtn");
const status = document.getElementById("status");

let selectedFile = null;

function setStatus(message, kind) {
  status.textContent = message;
  status.className = "status" + (kind ? ` ${kind}` : "");
}

function selectFile(file) {
  if (!file) return;
  const isMarkdown = /\.(md|markdown)$/i.test(file.name);
  if (!isMarkdown) {
    setStatus("Please choose a .md or .markdown file.", "error");
    return;
  }
  selectedFile = file;
  fileInfo.textContent = file.name;
  fileInfo.classList.remove("hidden");
  convertBtn.disabled = false;
  setStatus("");
}

dropzone.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", (e) => {
  selectFile(e.target.files[0]);
});

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("dragover");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragover");
});

dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("dragover");
  selectFile(e.dataTransfer.files[0]);
});

convertBtn.addEventListener("click", async () => {
  if (!selectedFile) return;
  convertBtn.disabled = true;
  setStatus("Converting...");

  try {
    const text = await selectedFile.text();
    const blob = await markdownToDocx(text);
    const url = URL.createObjectURL(blob);
    const outputName = selectedFile.name.replace(/\.(md|markdown)$/i, "") + ".docx";

    const a = document.createElement("a");
    a.href = url;
    a.download = outputName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);

    setStatus(`Saved ${outputName}`, "success");
  } catch (err) {
    console.error(err);
    setStatus("Conversion failed: " + err.message, "error");
  } finally {
    convertBtn.disabled = false;
  }
});
