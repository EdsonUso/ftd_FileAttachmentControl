import { IInputs, IOutputs } from "./generated/ManifestTypes";

// ─── Types ────────────────────────────────────────────────────────────────────

interface StagedFile {
  name: string;
  size: number;
  mimeType: string;
  base64: string;
  lastModified: number;
}

interface SavedFile {
  name: string;
  link: string;
  mimeType: string;
  id: string;
}

interface RemovedFile {
  id: string;
  link: string;
  name: string;
}

// ─── Constants ────────────────────────────────────────────────────────────────

const DEFAULT_MAX_MB = 50;
const DEFAULT_MAX_FILES = 3;

/**Lista de arquivos permitidos */
const ALLOWED_MIME = new Set<string>([
  // PDF
  "application/pdf",
  // Imagens
  "image/png",
  "image/jpeg",
  "image/svg+xml",
  // Office
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // .docx
  "application/vnd.openxmlformats-officedocument.presentationml.presentation", // .pptx
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
  // Office Antigo
  "application/msword", // .doc
  "application/vnd.ms-powerpoint", // .ppt
  "application/vnd.ms-excel", // .xls
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
  if (mime.includes("presentation") || mime.includes("powerpoint"))
    return "ppt";
  if (mime.includes("sheet") || mime.includes("excel")) return "xls";
  return "file";
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
}

export class FileAttachmentControl implements ComponentFramework.StandardControl<
  IInputs,
  IOutputs
> {
  // DOM
  private _container: HTMLDivElement;
  private _root: HTMLDivElement;
  private _dropZone: HTMLDivElement;
  private _fileInput: HTMLInputElement;
  private _chipList: HTMLDivElement;
  private _errorLabel: HTMLSpanElement;
  private _savedFileInput: HTMLDivElement;
  // private _browseBtn: HTMLButtonElement;

  // State
  private _savedFiles: SavedFile[] = [];
  private _stagedFiles: StagedFile[] = [];
  private _validationError = "";
  private _notifyOutputChanged: () => void;
  private _maxMB = DEFAULT_MAX_MB;
  private _allowMultiple = true;
  private _maxFile = DEFAULT_MAX_FILES;
  private _removedFile: RemovedFile[] = [];

  // Drag counter (to avoid flickering on child elements)
  private _dragCounter = 0;

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    _state: ComponentFramework.Dictionary,
    container: HTMLDivElement,
  ): void {
    this._container = container;
    this._notifyOutputChanged = notifyOutputChanged;
    this._buildSavedFiles(context);
    this._readInputs(context);
    this._buildDOM();
    this._attachEvents();
  }
  public _buildSavedFiles(context: ComponentFramework.Context<IInputs>): void {
    const rawValue = context.parameters.SavedFiles.raw;
    try {
      if (!rawValue || rawValue.trim() === "" || rawValue === "null") {
        this._savedFiles = [];
      } else {
        const parsed = JSON.parse(rawValue);
        this._savedFiles = Array.isArray(parsed) ? parsed : [parsed];
      }
    } catch {
      this._savedFiles = [];
    }
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
      FilesRemoved: JSON.stringify(this._removedFile),
    };
  }
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  public destroy(): void {}

  // ── Private helpers ────────────────────────────────────────────────────────

  private _readInputs(context: ComponentFramework.Context<IInputs>): void {
    const rawMB = context.parameters.MaxFileSizeMB?.raw;
    this._maxMB =
      typeof rawMB === "number" && rawMB > 0 ? rawMB : DEFAULT_MAX_MB;

    const rawMaxFile = context.parameters.MaxFile?.raw;
    this._maxFile =
      typeof rawMB === "number" && rawMaxFile != null && rawMaxFile > 0
        ? rawMaxFile
        : DEFAULT_MAX_FILES;

    const rawMultiple = context.parameters.AllowMultiple?.raw;
    this._allowMultiple =
      rawMultiple === undefined || rawMultiple === null
        ? true
        : Boolean(rawMultiple);
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
      "Arraste e solte arquivos aqui ou clique para selecionar",
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

    // this._browseBtn = document.createElement("button");
    // this._browseBtn.className = "fac-browse-btn";
    // this._browseBtn.type = "button";
    // this._browseBtn.textContent = "Selecionar arquivos";

    this._savedFileInput = document.createElement("div");
    this._savedFileInput.className = "fac-chip-list";
    this._renderSavedFiles();

    this._root.appendChild(this._dropZone);
    this._root.appendChild(this._fileInput);
    this._root.appendChild(this._errorLabel);
    this._root.appendChild(this._savedFileInput);
    this._root.appendChild(this._chipList);

    //this._root.appendChild(this._browseBtn);

    this._container.appendChild(this._root);
  }

  private _attachEvents(): void {
    this._dropZone.addEventListener("click", () => this._fileInput.click());
    this._dropZone.addEventListener("keydown", (e: KeyboardEvent) => {
      if (e.key === "Enter" || e.key === " ") this._fileInput.click();
    });
    // this._browseBtn.addEventListener("click", () => this._fileInput.click());

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
    console.log(files.length);
    console.log(this._stagedFiles.length);
    console.log(this._savedFiles.length);
    if (
      files.length + this._stagedFiles.length + this._savedFiles.length >
      this._maxFile
    ) {
      this._setError(`Só é permitido a inclusão de ${this._maxFile} arquivos.`);
      return;
    }
    const maxBytes = this._maxMB * 1024 * 1024;
    const errors: string[] = [];

    const pending = files.filter((f) => {
      if (!ALLOWED_MIME.has(f.type)) {
        errors.push(
          `"${f.name}": tipo não permitido (${f.type || "desconhecido"})`,
        );
        return false;
      }
      if (f.size > maxBytes) {
        errors.push(`"${f.name}": excede ${this._maxMB} MB`);
        return false;
      }

      const duplicate = this._stagedFiles.some(
        (s) => s.name === f.name && s.size === f.size,
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

  private _renderSavedFiles(): void {
    while (this._savedFileInput.firstChild) {
      this._savedFileInput.removeChild(this._savedFileInput.firstChild);
    }
    if (this._savedFiles?.length > 0) {
      const titleSavedFile = document.createElement("h3");
      titleSavedFile.textContent = "Arquivos Salvos";
      this._savedFileInput.appendChild(titleSavedFile);
    }

    for (let i = 0; i < this._savedFiles.length; i++) {
      const sf = this._savedFiles[i];
      const chip = document.createElement("button");

      chip.className = `fac-saved fac-chip fac-chip--${mimeLabel(sf.mimeType)}`;
      chip.type = "button";
      chip.addEventListener("click", (e) => {
        window.open(sf.link, "_blank");
      });

      const label = document.createElement("span");
      label.className = "fac-chip-badge";
      label.textContent = mimeLabel(sf.mimeType).toUpperCase();

      const name = document.createElement("span");
      name.className = "fac-chip-name";
      name.title = sf.name;
      name.textContent = sf.name;

      const removeBtn = document.createElement("button");
      removeBtn.className = "fac-chip-remove";
      removeBtn.type = "button";
      removeBtn.setAttribute("aria-label", `Remover ${sf.name}`);
      removeBtn.innerHTML =
        '<svg width="13" height="14" viewBox="0 0 13 14" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M4.37422 10.9355C4.37422 11.1761 4.17738 11.373 3.9368 11.373C3.69621 11.373 3.49937 11.1761 3.49937 10.9355V5.24906C3.49937 5.00848 3.69621 4.81164 3.9368 4.81164C4.17738 4.81164 4.37422 5.00848 4.37422 5.24906V10.9355ZM6.56133 10.9355C6.56133 11.1761 6.36449 11.373 6.12391 11.373C5.88332 11.373 5.68648 11.1761 5.68648 10.9355V5.24906C5.68648 5.00848 5.88332 4.81164 6.12391 4.81164C6.36449 4.81164 6.56133 5.00848 6.56133 5.24906V10.9355ZM8.74844 10.9355C8.74844 11.1761 8.5516 11.373 8.31102 11.373C8.07043 11.373 7.87359 11.1761 7.87359 10.9355V5.24906C7.87359 5.00848 8.07043 4.81164 8.31102 4.81164C8.5516 4.81164 8.74844 5.00848 8.74844 5.24906V10.9355ZM8.68009 0.681831L9.68343 2.18711H11.5917C11.9553 2.18711 12.2478 2.481 12.2478 2.84324C12.2478 3.20685 11.9553 3.49938 11.5917 3.49938H11.373V11.8104C11.373 13.0188 10.3942 13.9975 9.18586 13.9975H3.06195C1.85412 13.9975 0.874844 13.0188 0.874844 11.8104V3.49938H0.656133C0.293893 3.49938 0 3.20685 0 2.84324C0 2.481 0.293893 2.18711 0.656133 2.18711H2.56493L3.56772 0.681831C3.85205 0.25581 4.33048 0 4.84171 0H7.4061C7.91734 0 8.39577 0.255837 8.68009 0.681831ZM4.14184 2.18711H8.10597L7.58654 1.40959C7.54553 1.3489 7.47718 1.31227 7.4061 1.31227H4.84171C4.77063 1.31227 4.67768 1.3489 4.66128 1.40959L4.14184 2.18711ZM2.18711 11.8104C2.18711 12.2943 2.57888 12.6852 3.06195 12.6852H9.18586C9.66976 12.6852 10.0607 12.2943 10.0607 11.8104V3.49938H2.18711V11.8104Z" fill="#787878"/></svg>';

      const idx = i;
      removeBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        this._removeSavedFile(i, sf);
      });

      chip.appendChild(label);
      chip.appendChild(name);

      chip.appendChild(removeBtn);
      this._savedFileInput.appendChild(chip);
    }
  }

  private _renderChips(): void {
    while (this._chipList.firstChild) {
      this._chipList.removeChild(this._chipList.firstChild);
    }

    if (this._stagedFiles.length > 0) {
      const titlePendingFile = document.createElement("h3");
      titlePendingFile.textContent = "Arquivos Não Salvos";
      this._chipList.appendChild(titlePendingFile);
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
  private _removeSavedFile(index: number, removedFile: RemovedFile): void {
    this._removedFile.push(removedFile);
    this._savedFiles.splice(index, 1);
    this._renderSavedFiles();
    this._notifyOutputChanged();
  }
}
