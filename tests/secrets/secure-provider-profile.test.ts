import { describe, expect, it, vi } from "vitest";
import { createSecureProviderProfileStore } from "../../src/secrets/secure-provider-profile.js";
import type { ApiKeyProvider, CredentialStore } from "../../src/secrets/credential-store.js";
import {
  OPENAI_API_KEY_PREF,
  ANTHROPIC_API_KEY_PREF,
  CHAT_PROVIDER_PREF,
  createDefaultProviderProfileSettings
} from "../../src/preferences/provider-profile.js";

function fixture() {
  const prefs = new Map<string, string>();
  const vault = new Map<ApiKeyProvider, string>();
  const writes: string[] = [];
  const credentials: CredentialStore = {
    get: vi.fn((provider: ApiKeyProvider) => Promise.resolve(vault.get(provider) ?? null)),
    set: vi.fn((provider: ApiKeyProvider, secret: string) => {
      vault.set(provider, secret);
      return Promise.resolve();
    }),
    delete: vi.fn((provider: ApiKeyProvider) => {
      vault.delete(provider);
      return Promise.resolve();
    })
  };
  const writer = {
    set: (key: string, value: string) => {
      prefs.set(key, value);
      writes.push(value);
    },
    clear: (key: string) => {
      prefs.delete(key);
    }
  };
  const store = createSecureProviderProfileStore({
    prefs: { get: (key) => prefs.get(key) },
    writer,
    credentials,
    storageName: "test keyring"
  });
  return { store, prefs, vault, credentials, writes, writer };
}

describe("secure provider profiles", () => {
  it("warns accurately if preferences fail after a verified secure save", async () => {
    const f = fixture();
    vi.spyOn(f.writer, "set").mockImplementation(() => {
      throw new Error("sensitive backend error");
    });
    await expect(f.store.save({ ...f.store.read(), openaiApiKey: "new-key" })).rejects.toThrow(
      "API keys were saved securely"
    );
    expect(f.vault.get("openai")).toBe("new-key");
    expect(f.store.read().openaiApiKey).toBe("new-key");
    expect(f.prefs.has(OPENAI_API_KEY_PREF)).toBe(false);
  });

  it("does not repeat an unavailable-store attempt for every provider", async () => {
    const f = fixture();
    vi.mocked(f.credentials.get).mockRejectedValue(new Error("unavailable"));
    await f.store.initialize();
    expect(f.credentials.get).toHaveBeenCalledTimes(1);
  });
  it("migrates legacy keys only after verified storage, and loads them on restart", async () => {
    const f = fixture();
    f.prefs.set(OPENAI_API_KEY_PREF, "legacy-test-key");
    await f.store.initialize();
    expect(f.vault.get("openai")).toBe("legacy-test-key");
    expect(f.prefs.has(OPENAI_API_KEY_PREF)).toBe(false);
    expect(f.store.read().openaiApiKey).toBe("legacy-test-key");
    await f.store.initialize();
    expect(f.store.read().openaiApiKey).toBe("legacy-test-key");
    expect(f.writes).not.toContain("legacy-test-key");
  });

  it("preserves legacy keys on locked-store failure without using or exposing them", async () => {
    const f = fixture();
    f.prefs.set(OPENAI_API_KEY_PREF, "legacy-test-key");
    vi.mocked(f.credentials.get).mockRejectedValue(new Error("sensitive backend text"));
    await f.store.initialize();
    expect(f.prefs.get(OPENAI_API_KEY_PREF)).toBe("legacy-test-key");
    expect(f.store.read().openaiApiKey).toBe("");
    expect(f.store.message()).toContain("legacy");
    expect(f.store.message()).not.toContain("sensitive backend text");
    await expect(f.store.save(f.store.read())).rejects.toThrow("not migrated");
    expect(f.prefs.has(OPENAI_API_KEY_PREF)).toBe(true);
  });

  it("keeps legacy prefs if read-back does not confirm the migrated key", async () => {
    const f = fixture();
    f.prefs.set(OPENAI_API_KEY_PREF, "legacy-test-key");
    vi.mocked(f.credentials.get).mockResolvedValue(null);
    await f.store.initialize();
    expect(f.prefs.has(OPENAI_API_KEY_PREF)).toBe(true);
    expect(f.store.read().openaiApiKey).toBe("");
  });

  it("prefers an existing secure key over a stale legacy key", async () => {
    const f = fixture();
    f.prefs.set(OPENAI_API_KEY_PREF, "old-key");
    f.vault.set("openai", "current-key");
    await f.store.initialize();
    expect(f.store.read().openaiApiKey).toBe("current-key");
    expect(f.credentials.set).not.toHaveBeenCalled();
    expect(f.prefs.has(OPENAI_API_KEY_PREF)).toBe(false);
  });

  it("saves and deletes keys without putting secrets in preferences", async () => {
    const f = fixture();
    await f.store.initialize();
    await f.store.save({ ...f.store.read(), chatProvider: "codex-api", openaiApiKey: "new-key" });
    expect(f.prefs.get(CHAT_PROVIDER_PREF)).toBe("codex-api");
    expect(f.store.read().openaiApiKey).toBe("new-key");
    expect(f.prefs.has(OPENAI_API_KEY_PREF)).toBe(false);
    expect(f.writes).not.toContain("new-key");
    await f.store.save({ ...f.store.read(), openaiApiKey: "" });
    expect(f.vault.has("openai")).toBe(false);
    expect(f.store.read().openaiApiKey).toBe("");
  });

  it("rolls back earlier credential updates if a later provider fails", async () => {
    const f = fixture();
    f.vault.set("openai", "old-key");
    await f.store.initialize();
    vi.mocked(f.credentials.set).mockImplementation((provider, secret) => {
      if (provider === "anthropic") return Promise.reject(new Error("do not leak new-key"));
      f.vault.set(provider, secret);
      return Promise.resolve();
    });
    await expect(
      f.store.save({ ...f.store.read(), openaiApiKey: "new-key", anthropicApiKey: "other-key" })
    ).rejects.toThrow("Secure key save failed");
    expect(f.vault.get("openai")).toBe("old-key");
    expect(f.store.read().openaiApiKey).toBe("old-key");
    expect(f.prefs.has(CHAT_PROVIDER_PREF)).toBe(false);
    expect(f.prefs.has(ANTHROPIC_API_KEY_PREF)).toBe(false);
  });

  it("validates every key before making any credential changes", async () => {
    const f = fixture();
    await expect(
      f.store.save({
        ...createDefaultProviderProfileSettings(),
        openaiApiKey: "good",
        geminiApiKey: "bad\nkey"
      })
    ).rejects.toThrow("control characters");
    expect(f.credentials.set).not.toHaveBeenCalled();
    expect(f.writes).toEqual([]);
  });

  it("does not claim success or cache a key when secure storage is unavailable", async () => {
    const f = fixture();
    vi.mocked(f.credentials.get).mockRejectedValue(new Error("secret in backend error"));
    await expect(f.store.save({ ...f.store.read(), openaiApiKey: "new-key" })).rejects.toThrow(
      "Secure key save failed"
    );
    expect(f.store.read().openaiApiKey).toBe("");
    expect(f.writes).toEqual([]);
  });
});
