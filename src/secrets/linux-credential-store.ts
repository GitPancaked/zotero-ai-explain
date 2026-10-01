import type {
  ApiKeyProvider,
  CredentialCommandRunner,
  CredentialStore
} from "./credential-store.js";

const COMMAND = "/usr/bin/secret-tool";
const HELP =
  "Linux credential storage is unavailable or locked. On Arch, install libsecret and a Secret Service backend (for example: sudo pacman -S libsecret gnome-keyring). Start your desktop's Secret Service and unlock its password-protected default keyring, then retry.";

function attributes(provider: ApiKeyProvider): string[] {
  if (!["openai", "anthropic", "gemini"].includes(provider)) {
    throw new Error("Unsupported API key provider.");
  }
  return ["service", "zotero-ai-explain", "provider", provider];
}

/** Uses the user's Secret Service; never falls back to preferences or files. */
export function createLinuxCredentialStore(run: CredentialCommandRunner): CredentialStore {
  const invoke = async (args: string[], stdin?: string) => {
    try {
      return await run({ command: COMMAND, args, ...(stdin === undefined ? {} : { stdin }) });
    } catch {
      // Subprocess exceptions may contain stdin or captured output.
      throw new Error(HELP);
    }
  };

  const get = async (provider: ApiKeyProvider): Promise<string | null> => {
    const attrs = attributes(provider);
    const result = await invoke(["lookup", ...attrs]);
    if (result.exitCode === 0 && result.stdout.length > 0) {
      // Non-TTY secret-tool writes the exact secret, without a trailing newline.
      return result.stdout;
    }
    if (result.exitCode !== 1 || result.stdout.length > 0) {
      throw new Error(HELP);
    }
    // lookup returns 1 for both absence and errors. Search returns 0 for an
    // empty result and 1 for service errors; locked items are explicitly unlocked.
    // Any matched item means retrieval failed, rather than a missing credential.
    const search = await invoke(["search", "--all", "--unlock", ...attrs]);
    if (search.exitCode === 0 && search.stdout.length === 0) return null;
    throw new Error(HELP);
  };

  return {
    get,
    async set(provider, secret) {
      const attrs = attributes(provider);
      // libsecret's CLI reads at most 8192 bytes. API keys are single-line text.
      if (
        !secret ||
        /\s/.test(secret) ||
        secret.includes(String.fromCharCode(0)) ||
        new TextEncoder().encode(secret).length >= 8192
      ) {
        throw new Error("API key must be nonempty text without whitespace.");
      }
      const result = await invoke(["store", "--label=Zotero AI Explain API key", ...attrs], secret);
      if (result.exitCode !== 0) throw new Error(HELP);
      // A successful exit alone is insufficient to confirm durable storage.
      if ((await get(provider)) !== secret) throw new Error(HELP);
    },
    async delete(provider) {
      const attrs = attributes(provider);
      if ((await get(provider)) === null) return;
      const result = await invoke(["clear", ...attrs]);
      if (result.exitCode !== 0 || (await get(provider)) !== null) {
        throw new Error(HELP);
      }
    }
  };
}
