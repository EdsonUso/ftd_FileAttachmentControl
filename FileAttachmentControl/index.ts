import { IInputs, IOutputs } from "./generated/ManifestTypes";

// ─── Types ────────────────────────────────────────────────────────────────────

interface StagedFile {
  name: string;
  size: number;
  mimeType: string;
  base64: string;
  lastModified: number;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const DEFAULT_MAX_MB = 25;

/**Lista de arquivos permitidos */
const ALLOWED_MIME = new Set<string>([
  // PDF
  "application/pdf",
  // Imagens
  "image/png",
  "image/jpeg",
  "image/svg+xml",
  // Office
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document",   // .docx
  "application/vnd.openxmlformats-officedocument.presentationml.presentation", // .pptx
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",          // .xlsx
  // Office Antigo
  "application/msword",                                                         // .doc
  "application/vnd.ms-powerpoint",                                              // .ppt
  "application/vnd.ms-excel",                                                   // .xls
]);

const MIME_LABEL: Record<string, string> = {
  "application/pdf": "pdf",
  "image/png": "img",
  "image/jpeg": "img",
  "image/svg+xml": "img",
};
function mimeLabel(mime: string): string {
  if (MIME_LABEL[mime]) return MIME_LABEL[mime];
  if (mime.includes("word") || mime.includes("msword")) return "doc";
  if (mime.includes("presentation") || mime.includes("powerpoint")) return "ppt";
  if (mime.includes("sheet") || mime.includes("excel")) return "xls";
  return "file";
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}


export class FileAttachmentControl
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  // DOM
  private _container: HTMLDivElement;
  private _root: HTMLDivElement;
  private _dropZone: HTMLDivElement;
  private _fileInput: HTMLInputElement;
  private _chipList: HTMLDivElement;
  private _errorLabel: HTMLSpanElement;
  private _browseBtn: HTMLButtonElement;

  // State
  private _stagedFiles: StagedFile[] = [];
  private _validationError = "";
  private _notifyOutputChanged: () => void;
  private _maxMB = DEFAULT_MAX_MB;
  private _allowMultiple = true;

  // Drag counter (to avoid flickering on child elements)
  private _dragCounter = 0;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._container = container;
    this._notifyOutputChanged = notifyOutputChanged;

    this._readInputs(context);
    this._buildDOM();
    this._attachEvents();
  }

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._readInputs(context);

    this._fileInput.multiple = this._allowMultiple;

    const theme = (context.parameters.Theme?.raw ?? "light").toLowerCase();
    this._root.setAttribute("data-theme", theme === "dark" ? "dark" : "light");
  }

  public getOutputs(): IOutputs {
    return {
      FileCount: this._stagedFiles.length,
      FilesJson: JSON.stringify(this._stagedFiles),
      HasFiles: this._stagedFiles.length > 0,
      ValidationError: this._validationError,
    };
  }

  public destroy(): void {
  }

  // ── Private helpers ────────────────────────────────────────────────────────

  private _readInputs(context: ComponentFramework.Context<IInputs>): void {
    const rawMB = context.parameters.MaxFileSizeMB?.raw;
    this._maxMB = typeof rawMB === "number" && rawMB > 0 ? rawMB : DEFAULT_MAX_MB;

    const rawMultiple = context.parameters.AllowMultiple?.raw;
    this._allowMultiple =
      rawMultiple === undefined || rawMultiple === null ? true : Boolean(rawMultiple);
  }

  private _buildDOM(): void {
    // Root wrapper
    this._root = document.createElement("div");
    this._root.className = "fac-root";
    this._root.setAttribute("data-theme", "light");

    // Drop zone
    this._dropZone = document.createElement("div");
    this._dropZone.className = "fac-dropzone";
    this._dropZone.setAttribute("role", "button");
    this._dropZone.setAttribute("tabindex", "0");
    this._dropZone.setAttribute(
      "aria-label",
      "Arraste e solte arquivos aqui ou clique para selecionar"
    );

    const icon = document.createElement("div");
    icon.className = "fac-dz-icon";
    icon.innerHTML =
      '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2"/><polyline points="16 12 12 8 8 12"/><line x1="12" y1="8" x2="12" y2="20"/></svg>';

    const dzLabel = document.createElement("p");
    dzLabel.className = "fac-dz-label";
    dzLabel.innerHTML =
      'Arraste arquivos aqui ou <span class="fac-dz-link">navegue</span>';

    const dzSub = document.createElement("p");
    dzSub.className = "fac-dz-sub";
    dzSub.textContent = `PDF, Office, PNG, JPEG, SVG · Máx. ${this._maxMB} MB por arquivo`;

    this._dropZone.appendChild(icon);
    this._dropZone.appendChild(dzLabel);
    this._dropZone.appendChild(dzSub);

    // Hidden file input
    this._fileInput = document.createElement("input");
    this._fileInput.type = "file";
    this._fileInput.accept = [
      ".pdf",
      ".png",
      ".jpg",
      ".jpeg",
      ".svg",
      ".doc",
      ".docx",
      ".ppt",
      ".pptx",
      ".xls",
      ".xlsx",
    ].join(",");
    this._fileInput.multiple = this._allowMultiple;
    this._fileInput.style.display = "none";
    this._fileInput.setAttribute("aria-hidden", "true");

    this._errorLabel = document.createElement("span");
    this._errorLabel.className = "fac-error";
    this._errorLabel.setAttribute("role", "alert");
    this._errorLabel.style.display = "none";

    this._chipList = document.createElement("div");
    this._chipList.className = "fac-chip-list";

    this._browseBtn = document.createElement("button");
    this._browseBtn.className = "fac-browse-btn";
    this._browseBtn.type = "button";
    this._browseBtn.textContent = "Selecionar arquivos";

    this._root.appendChild(this._dropZone);
    this._root.appendChild(this._fileInput);
    this._root.appendChild(this._errorLabel);
    this._root.appendChild(this._chipList);
    this._root.appendChild(this._browseBtn);

    this._container.appendChild(this._root);
  }

  private _attachEvents(): void {

    this._dropZone.addEventListener("click", () => this._fileInput.click());
    this._dropZone.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key === "Enter" || e.key === " ") this._fileInput.click();
    });
    this._browseBtn.addEventListener("click", () => this._fileInput.click());

    this._fileInput.addEventListener("change", () => {
      if (this._fileInput.files) {
        this._processFiles(Array.from(this._fileInput.files));
        this._fileInput.value = "";
      }
    });

    this._dropZone.addEventListener("dragenter", (e: DragEvent) => {
      e.preventDefault();
      this._dragCounter++;
      this._dropZone.classList.add("fac-dropzone--active");
    });

    this._dropZone.addEventListener("dragleave", () => {
      this._dragCounter--;
      if (this._dragCounter === 0) {
        this._dropZone.classList.remove("fac-dropzone--active");
      }
    });

    this._dropZone.addEventListener("dragover", (e: DragEvent) => {
      e.preventDefault();
      e.stopPropagation();
    });

    this._dropZone.addEventListener("drop", (e: DragEvent) => {
      e.preventDefault();
      this._dragCounter = 0;
      this._dropZone.classList.remove("fac-dropzone--active");

      if (e.dataTransfer?.files) {
        this._processFiles(Array.from(e.dataTransfer.files));
      }
    });
  }

  private _processFiles(files: File[]): void {
    if (!this._allowMultiple) {
      this._stagedFiles = [];
    }

    const maxBytes = this._maxMB * 1024 * 1024;
    const errors: string[] = [];

    const pending = files.filter((f) => {
      if (!ALLOWED_MIME.has(f.type)) {
        errors.push(`"${f.name}": tipo não permitido (${f.type || "desconhecido"})`);
        return false;
      }
      if (f.size > maxBytes) {
        errors.push(`"${f.name}": excede ${this._maxMB} MB`);
        return false;
      }
      const duplicate = this._stagedFiles.some(
        (s) => s.name === f.name && s.size === f.size
      );
      if (duplicate) return false;
      return true;
    });

    if (errors.length > 0) {
      this._setError(errors.join(" | "));
    } else {
      this._setError("");
    }

    if (pending.length === 0) {
      if (errors.length > 0) {
        this._notifyOutputChanged();
      }
      return;
    }

    let readCount = 0;
    for (const file of pending) {
      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result as string;
        const base64 = dataUrl.split(",")[1] ?? "";

        this._stagedFiles.push({
          name: file.name,
          size: file.size,
          mimeType: file.type,
          base64,
          lastModified: file.lastModified,
        });

        readCount++;
        if (readCount === pending.length) {

          this._renderChips();
          this._notifyOutputChanged();
        }
      };
      reader.onerror = () => {
        this._setError(`Falha ao ler "${file.name}"`);
        readCount++;
        if (readCount === pending.length) {
          this._renderChips();
          this._notifyOutputChanged();
        }
      };
      reader.readAsDataURL(file);
    }
  }

  private _setError(msg: string): void {
    this._validationError = msg;
    if (msg) {
      this._errorLabel.textContent = msg;
      this._errorLabel.style.display = "block";
    } else {
      this._errorLabel.textContent = "";
      this._errorLabel.style.display = "none";
    }
  }

  private _renderChips(): void {

    while (this._chipList.firstChild) {
      this._chipList.removeChild(this._chipList.firstChild);
    }

    for (let i = 0; i < this._stagedFiles.length; i++) {
      const sf = this._stagedFiles[i];
      const chip = document.createElement("div");
      chip.className = `fac-chip fac-chip--${mimeLabel(sf.mimeType)}`;

      const label = document.createElement("span");
      label.className = "fac-chip-badge";
      label.textContent = mimeLabel(sf.mimeType).toUpperCase();

      const name = document.createElement("span");
      name.className = "fac-chip-name";
      name.title = sf.name;
      name.textContent = sf.name;

      const size = document.createElement("span");
      size.className = "fac-chip-size";
      size.textContent = formatBytes(sf.size);

      const removeBtn = document.createElement("button");
      removeBtn.className = "fac-chip-remove";
      removeBtn.type = "button";
      removeBtn.setAttribute("aria-label", `Remover ${sf.name}`);
      removeBtn.innerHTML =
        '<svg viewBox="0 0 16 16" fill="currentColor"><path d="M4.646 4.646a.5.5 0 01.708 0L8 7.293l2.646-2.647a.5.5 0 01.708.708L8.707 8l2.647 2.646a.5.5 0 01-.708.708L8 8.707l-2.646 2.647a.5.5 0 01-.708-.708L7.293 8 4.646 5.354a.5.5 0 010-.708z"/></svg>';

      const idx = i;
      removeBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        this._removeFile(idx);
      });

      chip.appendChild(label);
      chip.appendChild(name);
      chip.appendChild(size);
      chip.appendChild(removeBtn);
      this._chipList.appendChild(chip);
    }
  }

  private _removeFile(index: number): void {
    this._stagedFiles.splice(index, 1);
    this._renderChips();
    this._notifyOutputChanged();
  }
}
