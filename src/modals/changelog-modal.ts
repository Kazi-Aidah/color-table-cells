import { App, Modal, MarkdownRenderer, Component, requestUrl } from "obsidian";
import type TableColorPlugin from "../plugin";

export class ChangelogModal extends Modal {
  private plugin: TableColorPlugin;

  constructor(app: App, plugin: TableColorPlugin) {
    super(app);
    this.plugin = plugin;
  }

  async onOpen(): Promise<void> {
    const { contentEl, modalEl } = this;
    modalEl.addClass("ctc-release-modal");

    const header = contentEl.createDiv({ cls: "ctc-release-header" });
    header.createEl("h2", { text: "Release Notes", cls: "ctc-release-title" });
    const ghLink = header.createEl("a", {
      text: "View on GitHub",
      cls: "ctc-release-link",
    });
    ghLink.href = "https://github.com/Kazi-Aidah/color-table-cells/releases";
    ghLink.target = "_blank";

    const body = contentEl.createDiv({ cls: "ctc-release-body" });
    const loading = body.createEl("p", {
      text: "Loading release notes...",
      cls: "ctc-release-loading",
    });

    try {
      const releases = await this.plugin.fetchAllReleases();
      loading.remove();

      if (!releases || releases.length === 0) {
        body.createEl("p", {
          text: "No release notes available.",
          cls: "ctc-release-empty",
        });
        return;
      }

      for (const release of releases) {
        const meta = body.createDiv({ cls: "ctc-release-meta" });
        meta.createEl("div", {
          text: release.name || release.tag_name,
          cls: "ctc-release-name",
        });
        if (release.published_at) {
          meta.createEl("span", {
            text: new Date(release.published_at).toLocaleDateString(),
            cls: "ctc-release-date",
          });
        }

        const notesEl = body.createDiv({ cls: "ctc-release-notes" });
        if (release.body) {
          try {
            await MarkdownRenderer.render(
              this.app,
              release.body,
              notesEl,
              "",
              new Component(),
            );
          } catch {
            notesEl.createEl("pre", {
              text: release.body,
              cls: "ctc-release-notes-fallback",
            });
          }
        } else {
          notesEl.createEl("p", { text: "No notes for this release." });
        }
      }
    } catch (e) {
      loading.remove();
      body.createEl("p", {
        text: "Failed to load release notes. Check your internet connection.",
        cls: "ctc-release-error",
      });
    }
  }

  onClose(): void {
    this.contentEl.empty();
  }
}
