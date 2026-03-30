import { Plugin, Setting, PluginSettingTab, App, Notice, TFile, TFolder, MarkdownView, Menu, MenuItem } from 'obsidian';
import * as fs from 'fs/promises';
import * as path from 'path';
import i18n from './i18n';
import { ExportConfigModal } from './modal';
import type { BetterExportDocxSettings } from './type';

declare const electron: any;

const DEFAULT_SETTINGS: BetterExportDocxSettings = {
  showTitle: true,
  maxLevel: '6',
  displayHeader: true,
  displayFooter: true,
  headerTemplate: `<div style="width: 100vw;font-size:10px;text-align:center;"><span class="title"></span></div>`,
  footerTemplate: `<div style="width: 100vw;font-size:10px;text-align:center;"><span class="pageNumber"></span> / <span class="totalPages"></span></div>`,
  displayMetadata: false,
  debug: false,
  isTimestamp: false,
  enabledCss: false,
  concurrency: '5'
};

export default class ExportDocxPlugin extends Plugin {
  settings: BetterExportDocxSettings;
  i18n: any;

  async onload() {
    await this.loadSettings();
    this.i18n = i18n.current;
    this.registerCommand();
    this.registerSetting();
    this.registerEvents();
  }

  registerCommand() {
    this.addCommand({
      id: 'export-current-file-to-docx',
      name: this.i18n.exportCurrentFile,
      checkCallback: (checking: boolean) => {
        const view = this.app.workspace.getActiveViewOfType(MarkdownView);
        const file = view?.file;
        if (!file) {
          return false;
        }
        if (checking) {
          return true;
        }
        new ExportConfigModal(this, file).open();
        return true;
      }
    });
  }

  registerSetting() {
    this.addSettingTab(new ConfigSettingTab(this.app, this));
  }

  registerEvents() {
    this.registerEvent(
      this.app.workspace.on('file-menu', (menu: Menu, file) => {
        let title = file instanceof TFolder ? 'Export folder to DOCX' : 'Export to DOCX';
        menu.addItem((item) => {
          item
            .setTitle(title)
            .setIcon('download')
            .setSection('action')
            .onClick(async () => {
              new ExportConfigModal(this, file as TFile | TFolder).open();
            });
        });
      })
    );

    this.registerEvent(
      this.app.workspace.on('file-menu', (menu: Menu, file) => {
        if (file instanceof TFolder) {
          let title = 'Export to DOCX...';
          menu.addItem((item: MenuItem) => {
            item.setTitle(title).setIcon('folder-down').setSection('action');
            const subMenu = (item as any).setSubmenu();
            subMenu.addItem((item2: MenuItem) =>
              item2
                .setTitle('Export each file to DOCX')
                .setIcon('file-stack')
                .onClick(async () => {
                  new ExportConfigModal(this, file, true).open();
                })
            );
            subMenu.addItem((item2: MenuItem) =>
              item2
                .setTitle('Generate TOC.md file')
                .setIcon('file-text')
                .onClick(async () => {
                  await this.generateToc(file);
                })
            );
          });
        }
      })
    );
  }

  async generateToc(root: TFolder) {
    const basePath = (this.app.vault.adapter as any).basePath ?? '';
    const toc = path.join(basePath, root.path, '_TOC_.md');
    const content = `---
toc: true
title: ${root.name}
---
`;
    await fs.writeFile(toc, content);

    if (root instanceof TFolder) {
      const { traverseFolder } = await import('./utils');
      const files = traverseFolder(root);
      for (const file of files) {
        if (file.name === '_TOC_.md') {
          continue;
        }
        await fs.appendFile(toc, `[[${file.path}]]\n`);
      }
    }
  }

  onunload() {}

  async loadSettings() {
    this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.loadData());
  }

  async saveSettings() {
    await this.saveData(this.settings);
  }
}

function setAttributes(element: HTMLElement, attributes: Record<string, string>) {
  for (const key in attributes) {
    element.setAttribute(key, attributes[key]);
  }
}

function renderBuyMeACoffeeBadge(contentEl: HTMLElement, width: number = 175) {
  const linkEl = contentEl.createEl('a', {
    href: 'https://www.buymeacoffee.com/l1xnan'
  });
  const imgEl = linkEl.createEl('img');
  imgEl.src =
    'https://img.buymeacoffee.com/button-api/?text=Buy me a coffee&emoji=&slug=nathangeorge&button_colour=6a8696&font_colour=ffffff&font_family=Poppins&outline_colour=000000&coffee_colour=FFDD00';
  imgEl.alt = 'Buy me a coffee';
  imgEl.width = width;
}

class ConfigSettingTab extends PluginSettingTab {
  plugin: ExportDocxPlugin;
  i18n: any;

  constructor(app: App, plugin: ExportDocxPlugin) {
    super(app, plugin);
    this.plugin = plugin;
    this.i18n = i18n.current;
  }

  display() {
    const { containerEl } = this;
    containerEl.empty();

    const supportDesc = new DocumentFragment();
    supportDesc.createDiv({
      text: 'Support the continued development of this plugin.'
    });
    new Setting(containerEl).setDesc(supportDesc);
    renderBuyMeACoffeeBadge(containerEl);

    new Setting(containerEl)
      .setName(this.i18n.settings.showTitle)
      .addToggle((toggle) =>
        toggle
          .setTooltip(this.i18n.settings.showTitle)
          .setValue(this.plugin.settings.showTitle)
          .onChange(async (value) => {
            this.plugin.settings.showTitle = value;
            this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(this.i18n.settings.displayHeader)
      .addToggle((toggle) =>
        toggle
          .setTooltip('Display header')
          .setValue(this.plugin.settings.displayHeader)
          .onChange(async (value) => {
            this.plugin.settings.displayHeader = value;
            this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(this.i18n.settings.displayFooter)
      .addToggle((toggle) =>
        toggle
          .setTooltip('Display footer')
          .setValue(this.plugin.settings.displayFooter)
          .onChange(async (value) => {
            this.plugin.settings.displayFooter = value;
            this.plugin.saveSettings();
          })
      );

    new Setting(containerEl)
      .setName(this.i18n.settings.maxLevel)
      .addDropdown((dropdown) => {
        dropdown
          .addOptions(Object.fromEntries(['1', '2', '3', '4', '5', '6'].map((level) => [level, `h${level}`])))
          .setValue(this.plugin.settings.maxLevel)
          .onChange(async (value) => {
            this.plugin.settings.maxLevel = value;
            this.plugin.saveSettings();
          });
      });

    new Setting(containerEl)
      .setName(this.i18n.settings.displayMetadata)
      .setDesc('Add frontMatter(title, author, keywords, subject creator, etc) to docx metadata')
      .addToggle((toggle) =>
        toggle.setValue(this.plugin.settings.displayMetadata).onChange(async (value) => {
          this.plugin.settings.displayMetadata = value;
          this.plugin.saveSettings();
        })
      );

    new Setting(containerEl).setName('Advanced').setHeading();

    const headerContentAreaSetting = new Setting(containerEl);
    headerContentAreaSetting.settingEl.setAttribute('style', 'display: grid; grid-template-columns: 1fr;');
    headerContentAreaSetting
      .setName(this.i18n.settings.headerTemplate)
      .setDesc(
        'HTML template for the print header. Should be valid HTML markup with following classes used to inject printing values into them: date, title, author, pageNumber, totalPages.'
      );
    const headerContentArea = headerContentAreaSetting.addTextArea((textComponent) => {
      setAttributes(textComponent.inputEl, {
        style: 'margin-top: 12px; width: 100%; height: 6vh;'
      });
      textComponent.setValue(this.plugin.settings.headerTemplate).onChange(async (value: string) => {
        this.plugin.settings.headerTemplate = value;
        this.plugin.saveSettings();
      });
    });

    const footerContentAreaSetting = new Setting(containerEl);
    footerContentAreaSetting.settingEl.setAttribute('style', 'display: grid; grid-template-columns: 1fr;');
    footerContentAreaSetting
      .setName(this.i18n.settings.footerTemplate)
      .setDesc('HTML template for the print footer. Should use the same format as the headerTemplate.');
    const footerContentArea = footerContentAreaSetting.addTextArea((textComponent) => {
      setAttributes(textComponent.inputEl, {
        style: 'margin-top: 12px; width: 100%; height: 6vh;'
      });
      textComponent.setValue(this.plugin.settings.footerTemplate).onChange(async (value: string) => {
        this.plugin.settings.footerTemplate = value;
        this.plugin.saveSettings();
      });
    });

    new Setting(containerEl)
      .setName(this.i18n.settings.isTimestamp)
      .setDesc('Add timestamp to output file name')
      .addToggle((cb) => {
        cb.setValue(this.plugin.settings.isTimestamp).onChange(async (value) => {
          this.plugin.settings.isTimestamp = value;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName(this.i18n.settings.enabledCss)
      .setDesc('Select the css snippets that are not enabled')
      .addToggle((cb) => {
        cb.setValue(this.plugin.settings.enabledCss).onChange(async (value) => {
          this.plugin.settings.enabledCss = value;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl)
      .setName(this.i18n.settings.concurrency)
      .setDesc('Limit the number of concurrent renders')
      .addText((cb) => {
        const concurrency = this.plugin.settings?.concurrency;
        cb.setValue(concurrency?.length > 0 ? concurrency : '5').onChange(async (value) => {
          this.plugin.settings.concurrency = value;
          await this.plugin.saveSettings();
        });
      });

    new Setting(containerEl).setName('Debug').setHeading();

    new Setting(containerEl)
      .setName(this.i18n.settings.debugMode)
      .setDesc('This is useful for troubleshooting.')
      .addToggle((cb) => {
        cb.setValue(this.plugin.settings.debug).onChange(async (value) => {
          this.plugin.settings.debug = value;
          await this.plugin.saveSettings();
        });
      });
  }
}
