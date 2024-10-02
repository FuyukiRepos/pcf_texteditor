/* eslint-disable @typescript-eslint/no-explicit-any */
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import tinymce from "tinymce/tinymce";
import "tinymce/themes/silver";
import "tinymce/icons/default";
import "tinymce/models/dom";
import "tinymce/plugins/link";
import "tinymce/plugins/autolink";
import "tinymce/plugins/autoresize";
import "tinymce/plugins/table";
import "tinymce/plugins/lists";
import "tinymce/plugins/advlist";
import "tinymce/plugins/image";
import "tinymce/plugins/preview";

export class pcftexteditor
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private _context: ComponentFramework.Context<IInputs>;
  private _inputvalue: string;
  private _notifyOutputChanged: () => void;
  private _container: HTMLDivElement;
  private _editor: any;
  private _domId: string;
  private _tinymce: any;

  /**
   * Empty constructor.
   */
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    // Add control initialization code
    this._tinymce = tinymce;
    this._context = context;
    this._inputvalue = context.parameters.pcftexteditorfield.raw || "";
    this._notifyOutputChanged = notifyOutputChanged;
    this._container = container;
    this._domId = this.ID();

    this._editor = document.createElement("textarea");
    this._editor.setAttribute("id", "text_editor" + this._domId);
    this._editor.innerHTML = this._inputvalue;
    container.appendChild(this._editor);

    this.initializeTinyMCE();
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: ComponentFramework.Context<IInputs>): void {
    // Add code to update control view
    this._context = context;
    const newValue = this._context.parameters.pcftexteditorfield.raw || "";

    if (newValue !== this._inputvalue) {
      console.log("update value");
      this._inputvalue = newValue;
      if (this._editor) {
        this._tinymce.get(this._editor.id).setContent(this._inputvalue);
      }
      setTimeout(() => {
        this._notifyOutputChanged();
      }, 1000);
    }
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
   */
  public getOutputs(): IOutputs {
    return {
      pcftexteditorfield: this._inputvalue,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }

  private initializeTinyMCE() {
    // Load TinyMCE editor
    this._tinymce.init({
      selector: "#text_editor" + this._domId,
      skin: "oxide",
      icons: "default",
      theme: "silver",
      license_key: "gpl",
      width: "100%",
      height: "100%",
      promotion: false,
      branding: false,
      elementpath: false,
      plugins: [
        "table",
        "link",
        "autoresize",
        "autolink",
        "advlist",
        "preview",
        "lists",
        "image",
      ],
      menubar: "edit insert format table",
      toolbar:
        "undo redo | bold italic | fontsizeinput | alignleft aligncenter alignright alignjustify | outdent indent | bullist numlist",

      images_upload_handler: (
        blobInfo: { blob: () => Blob; filename: () => string | undefined },
        success: (arg0: any) => void,
        failure: (arg0: string) => void
      ) => {
        const xhr = new XMLHttpRequest();
        xhr.withCredentials = false;

        xhr.open("POST", "http://localhost");

        xhr.onload = () => {
          if (xhr.status < 200 || xhr.status >= 300) {
            failure("HTTP Error: " + xhr.status);
            return;
          }
          const json = JSON.parse(xhr.responseText);
          success(json.location); // URL of the uploaded image
        };

        const formData = new FormData();
        formData.append("file", blobInfo.blob(), blobInfo.filename());
        xhr.send(formData);
      },

      file_picker_callback: (
        callback: (arg0: string) => void,
        value: any,
        meta: { filetype: string }
      ) => {
        if (meta.filetype === "image") {
          // Open a file picker dialog
          const input = document.createElement("input");
          input.setAttribute("type", "file");
          input.setAttribute("accept", "image/*");

          input.onchange = () => {
            // Check if input.files is not null and has at least one file
            if (input.files && input.files.length > 0) {
              const file = input.files[0];
              const reader = new FileReader();

              reader.onload = (e) => {
                // Use type assertion to ensure e.target is defined
                const target = e.target as FileReader;
                if (target.result) {
                  // Call the callback with the image URL
                  callback(target.result as string);
                } else {
                  console.error("Failed to read file.");
                }
              };

              reader.readAsDataURL(file);
            } else {
              // Handle the case where no file was selected
              console.warn("No file selected or file input is null.");
            }
          };

          input.click();
        }
      },

      setup: (ed: any) => {
        ed.on("change", (e: any) => {
          this._inputvalue = ed.getContent();
          this._editor.innerHTML = this._inputvalue;
          this._notifyOutputChanged();
        });
      },
    });
  }

  ID = function () {
    // Math.random should be unique because of its seeding algorithm.
    // Convert it to base 36 (numbers + letters), and grab the first 9 characters
    // after the decimal.
    return "_" + Math.random().toString(36).substr(2, 9);
  };
}
