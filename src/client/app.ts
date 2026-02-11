interface ConversionResult {
  fileName: string;
  blob: Blob;
  success: true;
}

interface ConversionError {
  fileName: string;
  error: string;
  success: false;
}

type ConversionOutcome = ConversionResult | ConversionError;

class MsgToEmlConverter {
  private dropZone: HTMLElement;
  private fileInput: HTMLInputElement;
  private fileList: HTMLElement;
  private convertBtn: HTMLButtonElement;
  private downloadAllBtn: HTMLButtonElement;
  private status: HTMLElement;
  private selectedFiles: File[] = [];
  private results: ConversionOutcome[] = [];

  constructor() {
    this.dropZone = document.getElementById("drop-zone")!;
    this.fileInput = document.getElementById("file-input") as HTMLInputElement;
    this.fileList = document.getElementById("file-list")!;
    this.convertBtn = document.getElementById("convert-btn") as HTMLButtonElement;
    this.downloadAllBtn = document.getElementById("download-all-btn") as HTMLButtonElement;
    this.status = document.getElementById("status")!;

    this.setupEventListeners();
  }

  private setupEventListeners(): void {
    // Drag and drop events
    this.dropZone.addEventListener("dragover", (e) => {
      e.preventDefault();
      this.dropZone.classList.add("dragover");
    });

    this.dropZone.addEventListener("dragleave", () => {
      this.dropZone.classList.remove("dragover");
    });

    this.dropZone.addEventListener("drop", (e) => {
      e.preventDefault();
      this.dropZone.classList.remove("dragover");
      const files = Array.from(e.dataTransfer?.files || []);
      this.addFiles(files);
    });

    // Click to select files
    this.dropZone.addEventListener("click", () => {
      this.fileInput.click();
    });

    this.fileInput.addEventListener("change", () => {
      const files = Array.from(this.fileInput.files || []);
      this.addFiles(files);
      this.fileInput.value = "";
    });

    // Convert button
    this.convertBtn.addEventListener("click", () => {
      this.convertFiles();
    });

    // Download all button
    this.downloadAllBtn.addEventListener("click", () => {
      this.downloadAll();
    });
  }

  private addFiles(files: File[]): void {
    const msgFiles = files.filter((f) => f.name.toLowerCase().endsWith(".msg"));
    if (msgFiles.length === 0) {
      this.showStatus("Please select .msg files only", "error");
      return;
    }

    this.selectedFiles.push(...msgFiles);
    this.results = [];
    this.updateFileList();
    this.updateButtons();
  }

  private updateFileList(): void {
    if (this.selectedFiles.length === 0) {
      this.fileList.innerHTML = "<p>No files selected</p>";
      return;
    }

    const html = this.selectedFiles
      .map(
        (f, i) => `
      <div class="file-item">
        <span class="file-name">${this.escapeHtml(f.name)}</span>
        <span class="file-size">${this.formatSize(f.size)}</span>
        <button class="remove-btn" data-index="${i}">&times;</button>
      </div>
    `,
      )
      .join("");

    this.fileList.innerHTML = html;

    // Add remove button listeners
    this.fileList.querySelectorAll(".remove-btn").forEach((btn) => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        const index = parseInt((btn as HTMLElement).dataset.index!, 10);
        this.selectedFiles.splice(index, 1);
        this.updateFileList();
        this.updateButtons();
      });
    });
  }

  private updateButtons(): void {
    this.convertBtn.disabled = this.selectedFiles.length === 0;
    this.downloadAllBtn.disabled = this.results.filter((r) => r.success).length === 0;
  }

  private async convertFiles(): Promise<void> {
    this.results = [];
    this.convertBtn.disabled = true;
    this.showStatus(`Converting ${this.selectedFiles.length} file(s)...`, "info");

    for (const file of this.selectedFiles) {
      try {
        const result = await this.convertFile(file);
        this.results.push(result);
      } catch (error) {
        this.results.push({
          fileName: file.name,
          error: error instanceof Error ? error.message : String(error),
          success: false,
        });
      }
    }

    const successCount = this.results.filter((r) => r.success).length;
    const failCount = this.results.length - successCount;

    if (failCount === 0) {
      this.showStatus(`Successfully converted ${successCount} file(s)`, "success");
    } else {
      this.showStatus(
        `Converted ${successCount} file(s), ${failCount} failed`,
        failCount === this.results.length ? "error" : "warning",
      );
    }

    this.updateButtons();
    this.showResults();
  }

  private async convertFile(file: File): Promise<ConversionResult> {
    let response: Response;
    try {
      response = await fetch("/api/convert", {
        method: "POST",
        body: await file.arrayBuffer(),
        headers: {
          "Content-Type": "application/octet-stream",
        },
      });
    } catch {
      throw new Error("Network error: could not reach the server");
    }

    if (!response.ok) {
      let message = `Server error (${response.status})`;
      try {
        const error = await response.json();
        message = error.details || error.error || message;
      } catch {
        // Response body wasn't JSON — use the status text
        message = `Server error: ${response.status} ${response.statusText}`;
      }
      throw new Error(message);
    }

    const blob = await response.blob();
    const emlFileName = file.name.replace(/\.msg$/i, ".eml");

    return {
      fileName: emlFileName,
      blob,
      success: true,
    };
  }

  private showResults(): void {
    const resultsHtml = this.results
      .map((r) => {
        if (r.success) {
          return `
          <div class="result-item success">
            <span>${this.escapeHtml(r.fileName)}</span>
            <button class="download-btn" data-filename="${this.escapeHtml(r.fileName)}">Download</button>
          </div>
        `;
        } else {
          return `
          <div class="result-item error">
            <span>${this.escapeHtml(r.fileName)}</span>
            <span class="error-msg">${this.escapeHtml(r.error)}</span>
          </div>
        `;
        }
      })
      .join("");

    this.fileList.innerHTML = resultsHtml;

    // Add download button listeners
    this.fileList.querySelectorAll(".download-btn").forEach((btn) => {
      btn.addEventListener("click", () => {
        const fileName = (btn as HTMLElement).dataset.filename!;
        const result = this.results.find((r) => r.success && r.fileName === fileName) as ConversionResult;
        if (result) {
          this.downloadFile(result.blob, result.fileName);
        }
      });
    });
  }

  private downloadFile(blob: Blob, fileName: string): void {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  }

  private downloadAll(): void {
    const successful = this.results.filter((r): r is ConversionResult => r.success);
    successful.forEach((r) => {
      this.downloadFile(r.blob, r.fileName);
    });
  }

  private showStatus(message: string, type: "info" | "success" | "error" | "warning"): void {
    this.status.textContent = message;
    this.status.className = `status ${type}`;
  }

  private formatSize(bytes: number): string {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  private escapeHtml(str: string): string {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
  }
}

// Initialize app
document.addEventListener("DOMContentLoaded", () => {
  new MsgToEmlConverter();
});
