import {
  loadProviderProfileSettingsFromPrefs,
  saveProviderProfileSettingsToPrefs,
  OPENAI_API_KEY_PREF,
  ANTHROPIC_API_KEY_PREF,
  GEMINI_API_KEY_PREF,
  type ProviderProfileSettings
} from "../preferences/provider-profile.js";
import type { StringPrefReader, StringPrefWriter } from "../preferences/ollama-profile.js";
import type { ApiKeyProvider, CredentialStore } from "./credential-store.js";

const keys = [
  ["openai", "openaiApiKey", OPENAI_API_KEY_PREF],
  ["anthropic", "anthropicApiKey", ANTHROPIC_API_KEY_PREF],
  ["gemini", "geminiApiKey", GEMINI_API_KEY_PREF]
] as const satisfies readonly (readonly [ApiKeyProvider, keyof ProviderProfileSettings, string])[];

export function createSecureProviderProfileStore(deps: {
  readonly prefs: StringPrefReader;
  readonly writer: StringPrefWriter;
  readonly credentials: CredentialStore;
  readonly storageName: string;
}) {
  const cache = { openaiApiKey: "", anthropicApiKey: "", geminiApiKey: "" };
  let problem = "";
  let pending: Promise<unknown> = Promise.resolve();

  function clearLegacy(pref: string): void {
    if (deps.writer.clear !== undefined) deps.writer.clear(pref);
    else deps.writer.set(pref, "");
  }

  function validate(secret: string): string {
    const value = secret.trim();
    if (
      value.length > 1280 ||
      Array.from(value).some(
        (character) => character.charCodeAt(0) <= 32 || character.charCodeAt(0) === 127
      )
    ) {
      throw new Error(
        "API keys must be at most 1280 characters and contain no whitespace or control characters."
      );
    }
    return value;
  }

  return {
    async initialize(): Promise<void> {
      for (const [provider, field, pref] of keys) {
        try {
          const stored = await deps.credentials.get(provider);
          const legacy = deps.prefs.get(pref)?.trim() ?? "";
          if (stored !== null) {
            cache[field] = stored;
            if (legacy !== "") clearLegacy(pref);
          } else if (legacy !== "") {
            const value = validate(legacy);
            await deps.credentials.set(provider, value);
            if ((await deps.credentials.get(provider)) !== value) {
              throw new Error("Credential verification failed.");
            }
            cache[field] = value;
            // Retain the old pref on any failure before verified durable storage.
            clearLegacy(pref);
          }
        } catch {
          problem =
            " Secure storage could not be loaded or legacy migration could not finish. " +
            "Unlock your credential store and restart Zotero. Unmigrated legacy keys remain in preferences.";
          // Do not fall back to using a plaintext legacy credential.
          // Avoid repeating a slow/cancelled unlock prompt for every provider.
          break;
        }
      }
    },
    read(): ProviderProfileSettings {
      return { ...loadProviderProfileSettingsFromPrefs(deps.prefs), ...cache };
    },
    message(): string {
      return (
        `API keys are saved in ${deps.storageName}, never in Zotero preferences. ` +
        "Clear a key field and Save to remove its stored credential. " +
        (deps.storageName.includes("Linux")
          ? "Requires libsecret and a running, unlocked Secret Service keyring. "
          : "") +
        problem
      );
    },
    save(settings: ProviderProfileSettings): Promise<void> {
      // Serialize saves so two settings windows cannot interleave credential updates.
      const task = pending.then(async () => {
        const values = keys.map(([, field]) => validate(settings[field]));
        for (const [index, [, , pref]] of keys.entries()) {
          if (values[index] === "" && (deps.prefs.get(pref)?.trim() ?? "") !== "") {
            throw new Error(
              "A legacy API key has not migrated. Unlock secure storage and restart Zotero."
            );
          }
        }
        const previous = new Map<ApiKeyProvider, string | null>();
        const attempted: ApiKeyProvider[] = [];
        try {
          for (const [index, [provider, field]] of keys.entries()) {
            const value = values[index] ?? "";
            if (value !== cache[field]) {
              previous.set(provider, await deps.credentials.get(provider));
              attempted.push(provider);
              if (value === "") await deps.credentials.delete(provider);
              else await deps.credentials.set(provider, value);
              if ((await deps.credentials.get(provider)) !== (value === "" ? null : value)) {
                throw new Error("Credential verification failed; settings were not saved.");
              }
            }
          }
        } catch {
          let rollbackFailed = false;
          for (const provider of attempted.reverse()) {
            try {
              const old = previous.get(provider) ?? null;
              if (old === null) await deps.credentials.delete(provider);
              else await deps.credentials.set(provider, old);
              if ((await deps.credentials.get(provider)) !== old) rollbackFailed = true;
            } catch {
              rollbackFailed = true;
            }
          }
          throw new Error(
            rollbackFailed
              ? "Secure key save failed and previous keys could not be fully restored. Unlock your credential store and restart Zotero before retrying."
              : "Secure key save failed. Check that your credential store is installed, running, and unlocked. No keys were written to preferences."
          );
        }
        for (const [index, [, field]] of keys.entries()) cache[field] = values[index] ?? "";
        try {
          saveProviderProfileSettingsToPrefs(deps.writer, settings);
          for (const [, , pref] of keys) clearLegacy(pref);
        } catch {
          throw new Error(
            "API keys were saved securely, but Zotero preferences could not be updated or legacy keys could not be cleared. Restart Zotero and check settings before retrying."
          );
        }
        problem = "";
      });
      pending = task.catch(() => undefined);
      return task;
    }
  };
}
