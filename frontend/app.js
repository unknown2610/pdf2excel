const uploadForm = document.getElementById("uploadForm");
const statusBlock = document.getElementById("status");
const uploadTrigger = document.getElementById("uploadTrigger");
const demoTrigger = document.getElementById("demoTrigger");
const fileInput = document.getElementById("statement");
const preview = document.getElementById("preview");
const apiBaseInput = document.getElementById("apiBase");

uploadTrigger?.addEventListener("click", () => fileInput?.click());
demoTrigger?.addEventListener("click", () => {
  preview?.scrollIntoView({ behavior: "smooth" });
});

uploadForm?.addEventListener("submit", async (event) => {
  event.preventDefault();
  statusBlock.textContent = "Uploading and converting...";
  const data = new FormData(uploadForm);
  const apiBase = apiBaseInput?.value.trim();
  const endpoint = apiBase ? `${apiBase.replace(/\/$/, "")}/api/parse` : "/api/parse";

  try {
    const response = await fetch(endpoint, {
      method: "POST",
      body: data,
    });

    if (!response.ok) {
      const payload = await response.json();
      throw new Error(payload.error || "Conversion failed");
    }

    const blob = await response.blob();
    const url = window.URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = "statement.xlsx";
    anchor.click();
    window.URL.revokeObjectURL(url);

    statusBlock.textContent = "Conversion complete. Your download should start now.";
  } catch (error) {
    statusBlock.textContent = error.message;
  }
});
