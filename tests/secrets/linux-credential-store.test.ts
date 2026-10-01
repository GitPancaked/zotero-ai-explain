import { describe, expect, it, vi } from "vitest";
import { createLinuxCredentialStore } from "../../src/secrets/linux-credential-store.js";
import type {
  ApiKeyProvider,
  CredentialCommandRunner
} from "../../src/secrets/credential-store.js";

const ok = (stdout = "") => ({ exitCode: 0, stdout });
const missing = { exitCode: 1, stdout: "" };

describe("Linux credential storage", () => {
  it("transfers secrets only through stdin and preserves exact lookup output", async () => {
    const secret = "sk-'$()`\\secret";
    const run = vi
      .fn<CredentialCommandRunner>()
      .mockResolvedValueOnce(ok())
      .mockResolvedValueOnce(ok(secret));
    const store = createLinuxCredentialStore(run);
    await store.set("openai", secret);
    expect(run.mock.calls[0]?.[0]).toEqual({
      command: "/usr/bin/secret-tool",
      args: [
        "store",
        "--label=Zotero AI Explain API key",
        "service",
        "zotero-ai-explain",
        "provider",
        "openai"
      ],
      stdin: secret
    });
    expect(JSON.stringify(run.mock.calls.map(([call]) => call.args))).not.toContain(secret);
  });

  it("distinguishes a missing item from an inaccessible service", async () => {
    const run = vi
      .fn<CredentialCommandRunner>()
      .mockResolvedValueOnce(missing)
      .mockResolvedValueOnce(ok());
    expect(await createLinuxCredentialStore(run).get("gemini")).toBeNull();
    expect(run.mock.calls[1]?.[0].args).toEqual([
      "search",
      "--all",
      "--unlock",
      "service",
      "zotero-ai-explain",
      "provider",
      "gemini"
    ]);
    const unavailable = vi.fn().mockResolvedValue(missing);
    await expect(createLinuxCredentialStore(unavailable).get("gemini")).rejects.toThrow(
      "libsecret"
    );
  });

  it("does not treat locked items or canceled unlocks as missing", async () => {
    const run = vi
      .fn()
      .mockResolvedValueOnce(missing)
      .mockResolvedValueOnce(ok("[item]\nlabel = API key\n"));
    await expect(createLinuxCredentialStore(run).get("anthropic")).rejects.toThrow("unlock");
  });

  it("redacts subprocess exceptions and unsuccessful stdout", async () => {
    const secret = "sk-private";
    for (const run of [
      vi.fn().mockRejectedValue(new Error(secret)),
      vi.fn().mockResolvedValue({ exitCode: 2, stdout: secret })
    ]) {
      await expect(createLinuxCredentialStore(run).get("openai")).rejects.not.toThrow(secret);
    }
  });

  it("rejects forged providers and multiline or oversized credentials before launching", async () => {
    const run = vi.fn();
    const store = createLinuxCredentialStore(run);
    await expect(store.get("--help" as ApiKeyProvider)).rejects.toThrow("Unsupported");
    for (const secret of ["", "key\nsecond", "key\0tail", "x".repeat(8192)]) {
      await expect(store.set("openai", secret)).rejects.toThrow("API key");
    }
    expect(run).not.toHaveBeenCalled();
  });

  it("requires read-back confirmation before claiming storage succeeded", async () => {
    const run = vi.fn().mockResolvedValueOnce(ok()).mockResolvedValueOnce(ok("different"));
    await expect(createLinuxCredentialStore(run).set("openai", "secret")).rejects.toThrow(
      "credential storage"
    );
  });

  it("unlocks, clears only the selected namespace, and verifies deletion", async () => {
    const run = vi
      .fn<CredentialCommandRunner>()
      .mockResolvedValueOnce(ok("secret"))
      .mockResolvedValueOnce(ok())
      .mockResolvedValueOnce(missing)
      .mockResolvedValueOnce(ok());
    await createLinuxCredentialStore(run).delete("anthropic");
    expect(run.mock.calls[1]?.[0].args).toEqual([
      "clear",
      "service",
      "zotero-ai-explain",
      "provider",
      "anthropic"
    ]);
  });

  it("fails deletion if the secret is still present", async () => {
    const run = vi
      .fn()
      .mockResolvedValueOnce(ok("secret"))
      .mockResolvedValueOnce(ok())
      .mockResolvedValueOnce(ok("secret"));
    await expect(createLinuxCredentialStore(run).delete("openai")).rejects.toThrow(
      "credential storage"
    );
  });
});
