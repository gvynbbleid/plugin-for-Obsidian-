import { Modal, App, Notice, Setting, TFile, TFolder, debounce } from 'obsidian';
import * as fs from 'fs/promises';
import * as path from 'path';
import ExportDocxPlugin from './main';
import { TConfig, RenderResult } from './type';
import { renderMarkdown, fixDoc, getAllStyles, createWebview } from './render';
import { generateDocx } from './docx';
import { getHeadingTree, traverseFolder, safeParseInt, safeParseFloat, isNumber } from './utils';
import { PageSize } from './constant';
import i18n from './i18n';
import pLimit from 'p-limit';
import { mount, unmount } from 'svelte';
import Progress from './Progress.svelte';

declare const electron: any;

function fullWidthButton(button: any) {
  button.buttonEl.setAttribute('style', `margin: 0 auto; width: -webkit-fill-available`);
}

function setInputWidth(inputEl: HTMLElement) {
  inputEl.setAttribute('style', `width: 100px;`);
}

export class ExportConfigModal extends Modal {
  plugin: ExportDocxPlugin;
  file: TFile | TFolder;
  completed: boolean = false;
  i18n: any;
  docs: RenderResult[] = [];
  scale: number = 0.75;
  webviews: any[] = [];
  multipleDocx: boolean;
  config: TConfig;
  svelte: any;
  previewDiv!: HTMLElement;
  preview: any;

  constructor(plugin: ExportDocxPlugin, file: TFile | TFolder, multipleDocx: boolean = false) {
    super(plugin.app);
    this.plugin = plugin;
    this.file = file;
    this.multipleDocx = multipleDocx;
    this.i18n = i18n.current;
    this.config = {
      pageSize: 'A4',
      marginType: '1',
      showTitle: plugin.settings.showTitle ?? true,
      open: true,
      landscape: false,
      marginTop: '10',
      marginBottom: '10',
      marginLeft: '10',
      marginRight: '10',
      displayHeader: plugin.settings.displayHeader ?? true,
      displayFooter: plugin.settings.displayFooter ?? true,
      cssSnippet: '0',
      ...(plugin.settings.prevConfig ?? {})
    };
  }

  async getAllFiles() {
    const app = this.plugin.app;
    const data: any[] = [];
    const docs: RenderResult[] = [];

    if (this.file instanceof TFolder) {
      const files = traverseFolder(this.file);
      for (const file of files) {
        data.push({
          app,
          file,
          config: this.config
        });
      }
    } else {
      const { doc, frontMatter, file } = await renderMarkdown({ app, file: this.file, config: this.config });
      docs.push({ doc, frontMatter, file });

      if (frontMatter.toc) {
        const files = this.parseToc(doc);
        for (const item of files) {
          data.push({
            app,
            file: item.file,
            config: this.config,
            extra: item
          });
        }
      }
    }

    return { data, docs };
  }

  async renderFiles(data: any[], docs: RenderResult[], cb?: (i: number) => void) {
    const concurrency = safeParseInt(this.plugin.settings.concurrency) || 5;
    const limit = pLimit(concurrency);

    const inputs = data.map(
      (param, i) => limit(async () => {
        const res = await renderMarkdown(param);
        cb?.(i);
        return res;
      })
    );

    let renderedDocs = [...(docs ?? []), ...await Promise.all(inputs)];

    if (this.file instanceof TFile) {
      const leaf = this.app.workspace.getLeaf();
      await leaf.openFile(this.file);
    }

    if (!this.multipleDocx) {
      renderedDocs = this.mergeDoc(renderedDocs);
    }

    this.docs = renderedDocs.map(({ doc, ...rest }) => {
      return { ...rest, doc: fixDoc(doc, doc.title) };
    });
  }

  parseToc(doc: Document) {
    const cache = this.app.metadataCache.getFileCache(this.file as TFile);
    const files = cache?.links?.map(({ link, displayText }: any) => {
      const id = crypto.randomUUID();
      const elem = doc.querySelector(`a[data-href="${link}"]`);
      if (elem) {
        (elem as HTMLAnchorElement).href = `#${id}`;
      }
      return {
        title: displayText,
        file: this.app.metadataCache.getFirstLinkpathDest(link, this.file.path),
        id
      };
    }).filter((item) => item.file instanceof TFile) ?? [];

    return files;
  }

  mergeDoc(docs: RenderResult[]): RenderResult[] {
    const { doc: doc0, frontMatter, file } = docs[0];
    const sections: HTMLElement[] = [];

    for (const { doc } of docs) {
      const element = doc.querySelector('.markdown-preview-view');
      if (element) {
        const section = doc0.createElement('section');
        Array.from(element.children).forEach((child) => {
          section.appendChild(doc0.importNode(child, true));
        });
        sections.push(section);
      }
    }

    const root = doc0.querySelector('.markdown-preview-view');
    if (root) {
      root.innerHTML = '';
    }

    sections.forEach((section) => {
      root?.appendChild(section);
    });

    return [{ doc: doc0, frontMatter, file }];
  }

  calcPageSize(element?: HTMLElement, config?: TConfig) {
    const { pageSize, pageWidth } = config ?? this.config;
    const el = element ?? this.previewDiv;
    const width = PageSize[pageSize]?.[0] ?? safeParseFloat(pageWidth ?? '210', 210);
    const scale = Math.floor(this.mm2px(width) / el.offsetWidth * 100) / 100;

    this.webviews.forEach((wb) => {
      wb.style.transform = `scale(${1 / scale},${1 / scale})`;
      wb.style.width = `calc(${scale} * 100%)`;
      wb.style.height = `calc(${scale} * 100%)`;
    });

    this.scale = scale;
    return scale;
  }

  mm2px(mm: number): number {
    return Math.round(mm * 3.779527559);
  }

  px2mm(px: number): number {
    return Math.round(px * 0.26458333333719);
  }

  async calcWebviewSize() {
    await new Promise(resolve => setTimeout(resolve, 500));

    this.webviews.forEach(async (e, i) => {
      const [width, height] = await e.executeJavaScript('[document.body.offsetWidth, document.body.offsetHeight]');
      const sizeEl = e.parentNode?.querySelector('.print-size');
      if (sizeEl) {
        sizeEl.innerHTML = `${width}x${height}px\n${this.px2mm(width)}x${this.px2mm(height)}mm`;
      }
    });
  }

  async togglePrintSize() {
    document.querySelectorAll('.print-size')?.forEach((sizeEl: Element) => {
      if (this.config['pageSize'] === 'Custom') {
        (sizeEl as HTMLElement).style.visibility = 'visible';
      } else {
        (sizeEl as HTMLElement).style.visibility = 'hidden';
      }
    });
  }

  makeWebviewJs(doc: Document) {
    return `
      document.body.innerHTML = decodeURIComponent(\`${encodeURIComponent(doc.body.innerHTML)}\`);
      document.head.innerHTML = decodeURIComponent(\`${encodeURIComponent(document.head.innerHTML)}\`);

      function decodeAndReplaceEmbed(element) {
        element.innerHTML = decodeURIComponent(element.innerHTML);
        const newEmbeds = element.querySelectorAll("span.markdown-embed");
        newEmbeds.forEach(decodeAndReplaceEmbed);
      }

      document.querySelectorAll("span.markdown-embed").forEach(decodeAndReplaceEmbed);

      document.body.setAttribute("class", \`${document.body.getAttribute('class')}\`);
      document.body.setAttribute("style", \`${document.body.getAttribute('style')}\`);
      document.body.classList.add("theme-light");
      document.body.classList.remove("theme-dark");
      document.title = \`${doc.title}\`;
    `;
  }

  async appendWebview(e: HTMLElement, doc: Document) {
    const webview = createWebview(this.scale);
    const preview = e.appendChild(webview);
    this.webviews.push(preview);
    this.preview = preview;

    preview.addEventListener('dom-ready', async () => {
      this.completed = true;
      getAllStyles().forEach(async (css) => {
        await preview.insertCSS(css);
      });

      if (this.config.cssSnippet && this.config.cssSnippet !== '0') {
        try {
          const cssSnippet = await fs.readFile(this.config.cssSnippet, { encoding: 'utf8' });
          const printCss = cssSnippet.replace(/@media print\s*{([^}]+)}/g, '$1');
          await preview.insertCSS(printCss);
          await preview.insertCSS(cssSnippet);
        } catch (error) {
          console.warn(error);
        }
      }

      await preview.executeJavaScript(this.makeWebviewJs(doc));
    });
  }

  async appendWebviews(el: HTMLElement, render: boolean = true) {
    el.empty();

    if (render) {
      this.svelte = mount(Progress, {
        target: el,
        props: {
          startCount: 5
        }
      });

      const { data, docs } = await this.getAllFiles();
      this.svelte.initRenderStates(data);
      await this.renderFiles(data, docs, this.svelte.updateRenderStates);
    }

    el.empty();

    await Promise.all(
      this.docs?.map(async ({ doc }, i) => {
        if (this.multipleDocx) {
          el.createDiv({
            text: `${i + 1}-${doc.title}`,
            attr: { class: 'filename' }
          });
        }

        const div = el.createDiv({ attr: { class: 'webview-wrapper' } });
        div.createDiv({ attr: { class: 'print-size' } });
        await this.appendWebview(div, doc);
      })
    );

    await this.calcWebviewSize();
  }

  async onOpen() {
    this.contentEl.empty();
    this.containerEl.style.setProperty('--dialog-width', '60vw');
    this.titleEl.setText('Export to DOCX');

    const wrapper = this.contentEl.createDiv({ attr: { id: 'better-export-docx' } });

    const title = this.file instanceof TFile ? this.file.basename : this.file.name;

    this.previewDiv = wrapper.createDiv({ attr: { class: 'pdf-preview' } }, async (el) => {
      el.empty();
      const resizeObserver = new ResizeObserver(() => {
        this.calcPageSize(el);
      });
      resizeObserver.observe(el);
      await this.appendWebviews(el);
      this.togglePrintSize();
    });

    const contentEl = wrapper.createDiv({ attr: { class: 'setting-wrapper' } });

    const handleExport = async () => {
      this.plugin.settings.prevConfig = this.config;
      await this.plugin.saveSettings();

      if (this.config['pageSize'] === 'Custom') {
        if (!isNumber(this.config['pageWidth'] ?? '') || !isNumber(this.config['pageHeight'] ?? '')) {
          alert('When the page size is Custom, the Width/Height cannot be empty.');
          return;
        }
      }

      if (this.multipleDocx) {
        const outputPath = await this.getOutputPath(title);
        if (outputPath) {
          await Promise.all(
            this.docs.map(async ({ doc, frontMatter, file }, i) => {
              const data = await generateDocx(
                { doc, frontMatter, file },
                {
                  config: { ...this.plugin.settings, ...this.config } as any,
                  frontMatter,
                  displayMetadata: this.plugin.settings.displayMetadata,
                  maxLevel: safeParseInt(this.plugin.settings.maxLevel, 6),
                  headings: getHeadingTree(doc)
                }
              );

              const outputFile = path.join(outputPath, `${file.basename}.docx`);
              await fs.writeFile(outputFile, data);

              if (this.config.open) {
                electron.remote.shell.openPath(outputFile);
              }
            })
          );
          this.close();
        }
      } else {
        const outputFile = await this.getOutputFile(title, this.plugin.settings.isTimestamp);
        if (outputFile) {
          const { doc, frontMatter, file } = this.docs[0];
          const data = await generateDocx(
            { doc, frontMatter, file },
            {
              config: { ...this.plugin.settings, ...this.config } as any,
              frontMatter,
              displayMetadata: this.plugin.settings.displayMetadata,
              maxLevel: safeParseInt(this.plugin.settings.maxLevel, 6),
              headings: getHeadingTree(doc)
            }
          );

          await fs.writeFile(outputFile, data);

          if (this.config.open) {
            electron.remote.shell.openPath(outputFile);
          }

          this.close();
        }
      }
    };

    contentEl.addEventListener('keyup', (event) => {
      if ((event as KeyboardEvent).key === 'Enter') {
        handleExport();
      }
    });

    this.generateForm(contentEl);

    new Setting(contentEl).setHeading().addButton((button) => {
      button.setButtonText('Export').onClick(handleExport);
      button.setCta();
      fullWidthButton(button);
    });

    new Setting(contentEl).setHeading().addButton((button) => {
      button.setButtonText('Refresh').onClick(async () => {
        await this.appendWebviews(this.previewDiv);
      });
      fullWidthButton(button);
    });

    const debugEl = new Setting(contentEl).setHeading().addButton((button) => {
      button.setButtonText('Debug').onClick(async () => {
        this.preview?.openDevTools();
      });
      fullWidthButton(button);
    });
    (debugEl.settingEl as HTMLElement).hidden = !this.plugin.settings.debug;
  }

  generateForm(contentEl: HTMLElement) {
    new Setting(contentEl)
      .setName(this.i18n.exportDialog.filenameAsTitle)
      .addToggle((toggle) =>
        toggle
          .setTooltip('Include file name as title')
          .setValue(this.config['showTitle'])
          .onChange(async (value) => {
            this.config['showTitle'] = value;
            this.webviews.forEach((wv, i) => {
              wv.executeJavaScript(`
                var _title = document.querySelector("h1.__title__");
                if (_title) {
                  _title.style.display = "${value ? 'block' : 'none'}";
                }
              `);
              const _title = this.docs[i]?.doc?.querySelector('h1.__title__');
              if (_title) {
                (_title as HTMLElement).style.display = value ? 'block' : 'none';
              }
            });
          })
      );

    const pageSizes = ['A0', 'A1', 'A2', 'A3', 'A4', 'A5', 'A6', 'Legal', 'Letter', 'Tabloid', 'Ledger', 'Custom'];

    new Setting(contentEl)
      .setName(this.i18n.exportDialog.pageSize)
      .addDropdown((dropdown) => {
        dropdown
          .addOptions(Object.fromEntries(pageSizes.map((size) => [size, size])))
          .setValue(this.config.pageSize)
          .onChange(async (value) => {
            this.config['pageSize'] = value;
            if (value === 'Custom') {
              sizeEl.settingEl.hidden = false;
            } else {
              sizeEl.settingEl.hidden = true;
            }
            this.togglePrintSize();
            this.calcPageSize();
            await this.calcWebviewSize();
          });
      });

    const sizeEl = new Setting(contentEl)
      .setName('Width/Height')
      .addText((text) => {
        setInputWidth(text.inputEl);
        text
          .setPlaceholder('width')
          .setValue(this.config['pageWidth'] ?? '')
          .onChange(
            debounce(async (value) => {
              this.config['pageWidth'] = value;
              this.calcPageSize();
              await this.calcWebviewSize();
            }, 500, true)
          );
      })
      .addText((text) => {
        setInputWidth(text.inputEl);
        text
          .setPlaceholder('height')
          .setValue(this.config['pageHeight'] ?? '')
          .onChange((value) => {
            this.config['pageHeight'] = value;
          });
      });
    sizeEl.settingEl.hidden = this.config['pageSize'] !== 'Custom';

    new Setting(contentEl)
      .setName(this.i18n.exportDialog.margin)
      .setDesc('The unit is millimeters.')
      .addDropdown((dropdown) => {
        dropdown
          .addOption('0', 'None')
          .addOption('1', 'Default')
          .addOption('2', 'Small')
          .addOption('3', 'Custom')
          .setValue(this.config['marginType'])
          .onChange(async (value) => {
            this.config['marginType'] = value;
            if (value === '3') {
              topEl.settingEl.hidden = false;
              btmEl.settingEl.hidden = false;
            } else {
              topEl.settingEl.hidden = true;
              btmEl.settingEl.hidden = true;
            }
          });
      });

    const topEl = new Setting(contentEl)
      .setName('Top/Bottom')
      .addText((text) => {
        setInputWidth(text.inputEl);
        text
          .setPlaceholder('margin top')
          .setValue(this.config['marginTop'])
          .onChange((value) => {
            this.config['marginTop'] = value;
          });
      })
      .addText((text) => {
        setInputWidth(text.inputEl);
        text
          .setPlaceholder('margin bottom')
          .setValue(this.config['marginBottom'])
          .onChange((value) => {
            this.config['marginBottom'] = value;
          });
      });
    topEl.settingEl.hidden = this.config['marginType'] !== '3';

    const btmEl = new Setting(contentEl)
      .setName('Left/Right')
      .addText((text) => {
        setInputWidth(text.inputEl);
        text
          .setPlaceholder('margin left')
          .setValue(this.config['marginLeft'])
          .onChange((value) => {
            this.config['marginLeft'] = value;
          });
      })
      .addText((text) => {
        setInputWidth(text.inputEl);
        text
          .setPlaceholder('margin right')
          .setValue(this.config['marginRight'])
          .onChange((value) => {
            this.config['marginRight'] = value;
          });
      });
    btmEl.settingEl.hidden = this.config['marginType'] !== '3';

    new Setting(contentEl)
      .setName(this.i18n.exportDialog.displayHeader)
      .addToggle((toggle) =>
        toggle
          .setTooltip('Display header')
          .setValue(this.config['displayHeader'])
          .onChange(async (value) => {
            this.config['displayHeader'] = value;
          })
      );

    new Setting(contentEl)
      .setName(this.i18n.exportDialog.displayFooter)
      .addToggle((toggle) =>
        toggle
          .setTooltip('Display footer')
          .setValue(this.config['displayFooter'])
          .onChange(async (value) => {
            this.config['displayFooter'] = value;
          })
      );

    new Setting(contentEl)
      .setName(this.i18n.exportDialog.openAfterExport)
      .addToggle((toggle) =>
        toggle
          .setTooltip('Open the exported file after exporting.')
          .setValue(this.config['open'])
          .onChange(async (value) => {
            this.config['open'] = value;
          })
      );

    const snippets = this.cssSnippets();
    if (Object.keys(snippets).length > 0 && this.plugin.settings.enabledCss) {
      new Setting(contentEl)
        .setName(this.i18n.exportDialog.cssSnippets)
        .addDropdown((dropdown) => {
          dropdown
            .addOption('0', 'Not select')
            .addOptions(snippets)
            .setValue(this.config['cssSnippet'] ?? '0')
            .onChange(async (value) => {
              this.config['cssSnippet'] = value;
              await this.appendWebviews(this.previewDiv, false);
            });
        });
    }
  }

  onClose() {
    const { contentEl } = this;
    contentEl.empty();
    if (this.svelte) {
      unmount(this.svelte);
    }
  }

  cssSnippets(): Record<string, string> {
    const customCss = (this.app as any).customCss;
    const { snippets, enabledSnippets } = customCss ?? {};
    const basePath = (this.app.vault.adapter as any).basePath ?? '';

    return Object.fromEntries(
      (snippets ?? [])
        .filter((item: string) => !enabledSnippets?.has(item))
        .map((name: string) => {
          const file = path.join(basePath, '.obsidian/snippets', name + '.css');
          return [file, name];
        })
    );
  }

  async getOutputFile(filename: string, isTimestamp: boolean): Promise<string | undefined> {
    const result = await electron.remote.dialog.showSaveDialog({
      title: 'Export to DOCX',
      defaultPath: filename + (isTimestamp ? '-' + Date.now() : '') + '.docx',
      filters: [
        { name: 'All Files', extensions: ['*'] },
        { name: 'DOCX', extensions: ['docx'] }
      ],
      properties: ['showOverwriteConfirmation', 'createDirectory']
    });

    if (result.canceled) {
      return undefined;
    }

    return result.filePath;
  }

  async getOutputPath(filename: string): Promise<string | undefined> {
    const result = await electron.remote.dialog.showOpenDialog({
      title: 'Export to DOCX',
      defaultPath: filename,
      properties: ['openDirectory']
    });

    if (result.canceled) {
      return undefined;
    }

    return result.filePaths[0];
  }
}
