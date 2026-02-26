import { App, Plugin, PluginSettingTab, Setting } from "obsidian";

export interface XlsxCellsSettings {
  defaultFile: string;   // vault-relative path, e.g. "_external/Budget.xlsx"
  defaultSheet: string;  // sheet tab name, e.g. "Sheet1"
}

export const DEFAULT_SETTINGS: XlsxCellsSettings = {
  defaultFile: "",
  defaultSheet: "",
};

type SettingsPlugin = Plugin & {
  settings: XlsxCellsSettings;
  saveSettings(): Promise<void>;
};

export class XlsxCellsSettingTab extends PluginSettingTab {
  private plugin: SettingsPlugin;

  constructor(app: App, plugin: SettingsPlugin) {
    super(app, plugin);
    this.plugin = plugin;
  }

  display(): void {
    const { containerEl } = this;
    containerEl.empty();

    new Setting(containerEl).setName("XLSX Cells v2").setHeading();

    new Setting(containerEl)
      .setName("Default file")
      .setDesc(
        "Vault-relative path to the default .xlsx file " +
        "(e.g. _external/Budget.xlsx). " +
        "Used when a code block or inline reference omits the 'file' key."
      )
      .addText((text) =>
        text
          .setPlaceholder("_external/Budget.xlsx")
          .setValue(this.plugin.settings.defaultFile)
          .onChange(async (value) => {
            this.plugin.settings.defaultFile = value.trim();
            await this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName("Default sheet")
      .setDesc(
        "Default sheet tab name (e.g. Sheet1). " +
        "Used when a code block or inline reference omits the 'sheet' key."
      )
      .addText((text) =>
        text
          .setPlaceholder("Sheet1")
          .setValue(this.plugin.settings.defaultSheet)
          .onChange(async (value) => {
            this.plugin.settings.defaultSheet = value.trim();
            await this.plugin.saveSettings();
          })
      );
  }
}
