/* @vitest-environment jsdom */
import { describe, expect, it, vi } from "vitest";
import { renderSettingsView, wireSettingsView } from "../../src/ui/settings-view.js";
import { createDefaultOllamaSettings } from "../../src/preferences/ollama-profile.js";

describe("asynchronous credential saves", () => {
  it("waits for secure storage before reporting success or closing", async () => {
    const view = renderSettingsView({
      settings: createDefaultOllamaSettings(),
      indexStatus: { state: "idle", totalItems: 0, indexedItems: 0, failedItems: 0 }
    });
    let finish: (() => void) | undefined;
    const saved = new Promise<void>((resolve) => {
      finish = resolve;
    });
    const close = vi.fn();
    const onSave = vi.fn(() => saved);
    const handle = wireSettingsView({
      view,
      validate: () => Promise.resolve({ ok: true }),
      onSave,
      close,
      setTimeout: (callback) => {
        callback();
        return 0;
      }
    });
    view.querySelector<HTMLButtonElement>('[data-action="save-settings"]')?.click();
    await vi.waitFor(() => {
      expect(onSave).toHaveBeenCalled();
    });
    expect(close).not.toHaveBeenCalled();
    expect(view.querySelector('[data-role="status"]')?.textContent).not.toBe("Saved");
    finish?.();
    await vi.waitFor(() => {
      expect(close).toHaveBeenCalled();
    });
    expect(view.querySelector('[data-role="status"]')?.textContent).toBe("Saved");
    handle.detach();
  });

  it("keeps the form open and enables retry when a keyring save fails", async () => {
    const view = renderSettingsView({
      settings: createDefaultOllamaSettings(),
      indexStatus: { state: "idle", totalItems: 0, indexedItems: 0, failedItems: 0 }
    });
    const close = vi.fn();
    const handle = wireSettingsView({
      view,
      validate: () => Promise.resolve({ ok: true }),
      onSave: () => Promise.reject(new Error("Unlock your keyring.")),
      close
    });
    const button = view.querySelector<HTMLButtonElement>('[data-action="save-settings"]');
    button?.click();
    await vi.waitFor(() => {
      expect(view.textContent).toContain("Unlock your keyring.");
    });
    expect(close).not.toHaveBeenCalled();
    expect(button?.disabled).toBe(false);
    expect(view.querySelector('[data-role="status"]')?.textContent).not.toBe("Saved");
    handle.detach();
  });
});
