import { afterEach, describe, expect, it, vi } from "vitest";
import type {
  CredentialProcess,
  CredentialSubprocess
} from "../../src/platform/credential-command.js";
import { createSystemCredentialStore } from "../../src/platform/system-credential-store.js";

function installHost(platform: string, environment: Record<string, string> = {}) {
  const writes: string[] = [];
  const replies: string[] = [];
  const call = vi.fn<CredentialSubprocess["call"]>(() => {
    const reply = replies.shift() ?? "{}";
    const child: CredentialProcess = {
      stdin: {
        write: (value) => {
          writes.push(value);
          return Promise.resolve();
        },
        close: () => Promise.resolve()
      },
      stdout: {
        read: vi
          .fn()
          .mockResolvedValueOnce(new TextEncoder().encode(reply).buffer)
          .mockResolvedValue(new ArrayBuffer(0))
      },
      stderr: { read: () => Promise.resolve(new ArrayBuffer(0)) },
      wait: () => Promise.resolve({ exitCode: 0 }),
      kill: () => Promise.resolve()
    };
    return Promise.resolve(child);
  });
  const subprocess = { call, getEnvironment: () => environment };
  const importESModule = vi.fn((spec: string) => {
    if (spec === "resource://gre/modules/Subprocess.sys.mjs") return { Subprocess: subprocess };
    if (spec === "resource://gre/modules/AppConstants.sys.mjs")
      return { AppConstants: { platform } };
    throw new Error("Unexpected module");
  });
  vi.stubGlobal("ChromeUtils", { importESModule });
  return { call, writes, replies, importESModule };
}

afterEach(() => vi.unstubAllGlobals());

describe("system credential store", () => {
  it.each([
    { SystemRoot: "C:\\Windows", PATH: "C:\\untrusted" },
    { SYSTEMROOT: "D:\\Windows\\", PATH: "C:\\untrusted" }
  ])("uses the Windows system PowerShell path and only stdin for keys", async (environment) => {
    const host = installHost("win", environment);
    const { credentials, storageName } = createSystemCredentialStore();
    await credentials.set("openai", "synthetic-secret");
    const input = host.call.mock.calls[0]?.[0];
    expect(input?.command).toBe(
      `${(environment.SystemRoot ?? environment.SYSTEMROOT).replace(/\\$/u, "")}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe`
    );
    expect(input?.arguments.join(" ")).not.toContain("synthetic-secret");
    expect(JSON.parse(host.writes[0] ?? "")).toEqual({
      operation: "set",
      provider: "openai",
      secret: "synthetic-secret"
    });
    expect(storageName).toBe("Windows Credential Manager");
  });

  it("uses the absolute Linux helper and verifies the stored secret", async () => {
    const host = installHost("linux", { PATH: "/untrusted" });
    host.replies.push("", "synthetic-secret");
    const { credentials, storageName } = createSystemCredentialStore();
    await credentials.set("anthropic", "synthetic-secret");
    expect(host.call).toHaveBeenCalledTimes(2);
    for (const [input] of host.call.mock.calls) {
      expect(input.command).toBe("/usr/bin/secret-tool");
      expect(input.arguments).not.toContain("synthetic-secret");
    }
    expect(host.writes).toEqual(["synthetic-secret"]);
    expect(storageName).toContain("Linux");
  });

  it.each([{}, { SystemRoot: "relative" }, { SystemRoot: "\\\\host\\share" }])(
    "rejects missing or nonlocal Windows installation paths",
    async (environment) => {
      const host = installHost("win", environment);
      await expect(
        createSystemCredentialStore().credentials.set("openai", "synthetic-secret")
      ).rejects.toThrow("could not be initialized");
      expect(host.call).not.toHaveBeenCalled();
    }
  );

  it("rejects unsupported platforms without invoking a helper", async () => {
    const host = installHost("macosx");
    await expect(
      createSystemCredentialStore().credentials.set("openai", "synthetic-secret")
    ).rejects.toThrow("supports Windows and Linux only");
    expect(host.call).not.toHaveBeenCalled();
  });

  it("rejects unavailable host modules without revealing native errors", async () => {
    vi.stubGlobal("ChromeUtils", {
      importESModule: () => {
        throw new Error("sensitive-native-error");
      }
    });
    await expect(
      createSystemCredentialStore().credentials.set("openai", "synthetic-secret")
    ).rejects.toThrow(/^Secure credential storage could not be initialized\./);
  });

  it("rejects a host without ChromeUtils", async () => {
    vi.stubGlobal("ChromeUtils", undefined);
    const store = createSystemCredentialStore().credentials;
    await expect(store.set("openai", "synthetic-secret")).rejects.toThrow(
      "unavailable in this host"
    );
    await expect(store.get("openai")).rejects.toThrow();
    await expect(store.delete("openai")).rejects.toThrow();
  });
});
