import { IInputs, IOutputs } from "./generated/ManifestTypes";

/* ================================
   Firma Editor Type Definitions
================================ */

interface IFirmaTemplateEditorInstance {
  destroy(): void;
}

interface IFirmaTemplateEditorOptions {
  container: HTMLElement;
  jwt: string;
  templateId: string;
  theme?: "light" | "dark";
  readOnly?: boolean;
  width?: string;
  height?: string;
  onSave?: (data: unknown) => void;
  onLoad?: (template: unknown) => void;
  onError?: (error: unknown) => void;
}

type FirmaTemplateEditorConstructor = new (
  options: IFirmaTemplateEditorOptions
) => IFirmaTemplateEditorInstance;

/* ================================
   PCF Control
================================ */

export class IFrameControl
  implements ComponentFramework.StandardControl<IInputs, IOutputs> {

  private _context?: ComponentFramework.Context<IInputs>;
  private _container!: HTMLDivElement;
  private _editorContainer!: HTMLDivElement;
  private _editorInstance?: IFirmaTemplateEditorInstance;
  private _scriptLoaded = false;
  private _currentTemplateId?: string;
  private _currentJwt?: string;

  /* ================================
     INIT
  ================================ */

  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this._context = context;
    this._container = container;

    this._editorContainer = document.createElement("div");
    this._editorContainer.style.width = "100%";
    this._editorContainer.style.height = "100%";
    this._editorContainer.style.minHeight = "300px";

    this._container.appendChild(this._editorContainer);

    this.loadFirmaScript();
  }

  /* ================================
     UPDATE VIEW
  ================================ */

  public updateView(context: ComponentFramework.Context<IInputs>): void {
    this._context = context;

    const templateId = context.parameters.templateId?.raw ?? "";
    const jwt = context.parameters.jwt?.raw ?? "";

    if (!this._scriptLoaded || !templateId || !jwt) {
      return;
    }

    if (
      this._editorInstance &&
      templateId === this._currentTemplateId &&
      jwt === this._currentJwt
    ) {
      return;
    }

    this._currentTemplateId = templateId;
    this._currentJwt = jwt;

    this.initializeEditor(templateId, jwt);
  }

  /* ================================
     SCRIPT LOADER
  ================================ */

  private loadFirmaScript(): void {
    if (this.getFirmaConstructor()) {
      this._scriptLoaded = true;
      this.tryInitialize();
      return;
    }

    const script = document.createElement("script");
    script.src =
      "https://api.firma.dev/functions/v1/embed-proxy/template-editor.js";
    script.async = true;

    script.onload = () => {
      this._scriptLoaded = true;
      this.tryInitialize();
    };

    script.onerror = () => {
      console.error("Failed to load Firma Template Editor script");
    };

    document.body.appendChild(script);
  }

  /* ================================
     SAFE ACCESSOR
  ================================ */

  private getFirmaConstructor(): FirmaTemplateEditorConstructor | undefined {
    const win = window as unknown as {
      FirmaTemplateEditor?: FirmaTemplateEditorConstructor;
    };

    return win.FirmaTemplateEditor;
  }

  /* ================================
     INITIALIZE (SAFE)
  ================================ */

  private tryInitialize(): void {
    if (!this._context) return;

    const templateId = this._context.parameters.templateId?.raw ?? "";
    const jwt = this._context.parameters.jwt?.raw ?? "";

    if (!templateId || !jwt || !this._scriptLoaded) {
      return;
    }

    this.initializeEditor(templateId, jwt);
  }

  /* ================================
     CREATE EDITOR
  ================================ */

  private initializeEditor(templateId: string, jwt: string): void {
    if (this._editorInstance) {
      this._editorInstance.destroy();
    }

    const allocatedHeight =
      this._context?.mode.allocatedHeight && this._context.mode.allocatedHeight > 0
        ? `${this._context.mode.allocatedHeight}px`
        : "600px";

    this._editorContainer.innerHTML = "";
    this._editorContainer.style.height = allocatedHeight;

    const FirmaEditor = this.getFirmaConstructor();

    if (!FirmaEditor) {
      console.error("FirmaTemplateEditor not available");
      return;
    }

    this._editorInstance = new FirmaEditor({
      container: this._editorContainer,
      jwt,
      templateId,
      theme: "dark",
      readOnly: false,
      width: "100%",
      height: allocatedHeight,
      onSave: (data: unknown) => console.log("Saved:", data),
      onLoad: (template: unknown) => console.log("Loaded:", template),
      onError: (error: unknown) => console.error("Firma Error:", error),
    });
  }

  /* ================================
     OUTPUTS
  ================================ */

  public getOutputs(): IOutputs {
    return {};
  }

  /* ================================
     DESTROY
  ================================ */

  public destroy(): void {
    if (this._editorInstance) {
      this._editorInstance.destroy();
      this._editorInstance = undefined;
    }

    this._container.innerHTML = "";
  }
}
