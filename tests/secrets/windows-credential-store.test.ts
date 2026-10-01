import { describe, expect, it, vi } from "vitest";
import { spawn } from "node:child_process";
import { randomUUID } from "node:crypto";
import type {
  ApiKeyProvider,
  CredentialCommandRunner
} from "../../src/secrets/credential-store.js";
import { createWindowsCredentialStore } from "../../src/secrets/windows-credential-store.js";

describe("Windows Credential Manager", () => {
  it("sends Unicode secrets only through stdin and uses fixed command arguments", async () => {
    const run = vi.fn<CredentialCommandRunner>().mockResolvedValue({ exitCode: 0, stdout: "{}" });
    const store = createWindowsCredentialStore(run);
    const secret = 'sensitive-雪-🔑-"-$(Get-Process)';
    await store.set("openai", secret);
    await store.delete("gemini");
    const input = run.mock.calls[0]?.[0];
    if (!input) throw new Error("Missing credential command");
    expect(input.command).toBe("powershell.exe");
    expect(input.args).toEqual(run.mock.calls[1]?.[0].args);
    expect(input.args.join(" ")).not.toContain(secret);
    expect(JSON.parse(input.stdin ?? "")).toEqual({ operation: "set", provider: "openai", secret });
  });

  it("returns a Unicode secret and distinguishes missing credentials", async () => {
    const run = vi
      .fn()
      .mockResolvedValueOnce({ exitCode: 0, stdout: JSON.stringify({ secret: "雪🔑" }) })
      .mockResolvedValueOnce({ exitCode: 0, stdout: '{"secret":null}' });
    const store = createWindowsCredentialStore(run);
    expect(await store.get("anthropic")).toBe("雪🔑");
    expect(await store.get("gemini")).toBeNull();
  });

  it.each([
    { exitCode: 1, stdout: "secret=do-not-disclose" },
    { exitCode: 0, stdout: "do-not-disclose" },
    { exitCode: 0, stdout: "{}" },
    { exitCode: 0, stdout: '{"secret":123}' }
  ])("fails closed for failed or malformed reads", async (result) => {
    const store = createWindowsCredentialStore(vi.fn().mockResolvedValue(result));
    await expect(store.get("openai")).rejects.toThrow("Windows Credential Manager is unavailable");
  });

  it("redacts runner errors for all operations", async () => {
    const run = vi.fn().mockRejectedValue(new Error("my-secret"));
    const store = createWindowsCredentialStore(run);
    for (const operation of [
      store.get("openai"),
      store.set("openai", "my-secret"),
      store.delete("openai")
    ]) {
      await expect(operation).rejects.toThrow(
        /^Windows Credential Manager is unavailable or the credential operation failed\.$/
      );
    }
  });

  it("rejects unknown providers and excessive secrets before launching a process", async () => {
    const run = vi.fn();
    const store = createWindowsCredentialStore(run);
    await expect(store.get("../other-app" as ApiKeyProvider)).rejects.toThrow();
    await expect(store.set("openai", "a".repeat(1281))).rejects.toThrow();
    await expect(store.set("openai", "")).rejects.toThrow();
    expect(run).not.toHaveBeenCalled();
  });

  it.skipIf(process.platform !== "win32")(
    "round-trips a synthetic credential through native Windows APIs",
    async () => {
      const namespace = `zotero-ai-explain-test-${randomUUID()}/`;
      const store = createWindowsCredentialStore(async (input) => {
        const args = [...input.args];
        const index = args.indexOf("-EncodedCommand") + 1;
        const encoded = args[index];
        if (!encoded) throw new Error("Missing encoded credential script");
        const script = Buffer.from(encoded, "base64")
          .toString("utf16le")
          .replace("'zotero-ai-explain/'", `'${namespace}'`);
        args[index] = Buffer.from(script, "utf16le").toString("base64");
        return await new Promise((resolve, reject) => {
          const child = spawn(input.command, args, {
            windowsHide: true,
            stdio: ["pipe", "pipe", "pipe"]
          });
          let stdout = "";
          child.stdout.setEncoding("utf8");
          child.stdout.on("data", (data: string) => {
            stdout += data;
          });
          child.stderr.resume();
          child.on("error", () => {
            reject(new Error("Native credential test process failed"));
          });
          child.on("close", (exitCode) => {
            resolve({ exitCode: exitCode ?? 1, stdout });
          });
          child.stdin.end(input.stdin);
        });
      });
      try {
        expect(await store.get("openai")).toBeNull();
        await store.set("openai", "synthetic-雪-🔑");
        expect(await store.get("openai")).toBe("synthetic-雪-🔑");
        await store.set("openai", "synthetic-replacement");
        expect(await store.get("openai")).toBe("synthetic-replacement");
        await store.delete("openai");
        expect(await store.get("openai")).toBeNull();
        await store.delete("openai");
      } finally {
        await store.delete("openai");
      }
    },
    60000
  );
});
