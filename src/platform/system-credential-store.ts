import type { CredentialStore } from "../secrets/credential-store.js";
import { createWindowsCredentialStore } from "../secrets/windows-credential-store.js";
import { createLinuxCredentialStore } from "../secrets/linux-credential-store.js";
import { createCredentialCommandRunner, type CredentialSubprocess } from "./credential-command.js";

export function createSystemCredentialStore(): {
  readonly credentials: CredentialStore;
  readonly storageName: string;
} {
  const unavailable = (message: string) => {
    const fail = (): Promise<never> => Promise.reject(new Error(message));
    return {
      credentials: { get: fail, set: fail, delete: fail },
      storageName: "the OS credential store"
    };
  };
  type Imports = {
    importESModule(spec: "resource://gre/modules/Subprocess.sys.mjs"): {
      Subprocess: CredentialSubprocess;
    };
    importESModule(spec: "resource://gre/modules/AppConstants.sys.mjs"): {
      AppConstants: { platform: string };
    };
  };
  const chrome = (globalThis as unknown as { ChromeUtils?: Imports }).ChromeUtils;
  if (chrome === undefined)
    return unavailable("Secure credential storage is unavailable in this host.");
  try {
    const { Subprocess } = chrome.importESModule("resource://gre/modules/Subprocess.sys.mjs");
    const { AppConstants } = chrome.importESModule("resource://gre/modules/AppConstants.sys.mjs");
    const run = createCredentialCommandRunner(Subprocess);
    if (AppConstants.platform === "win") {
      // Use the Windows installation directory instead of trusting user PATH.
      const env = (
        Subprocess as CredentialSubprocess & {
          getEnvironment(): Record<string, string | undefined>;
        }
      ).getEnvironment();
      const root = env.SystemRoot ?? env.SYSTEMROOT;
      if (root === undefined || !/^[A-Za-z]:\\/u.test(root)) {
        return unavailable("Windows Credential Manager could not be initialized.");
      }
      const command = `${root.replace(/\\$/u, "")}\\System32\\WindowsPowerShell\\v1.0\\powershell.exe`;
      return {
        credentials: createWindowsCredentialStore((input) => run({ ...input, command })),
        storageName: "Windows Credential Manager"
      };
    }
    if (AppConstants.platform === "linux") {
      return {
        credentials: createLinuxCredentialStore(run),
        storageName: "your Linux Secret Service keyring"
      };
    }
    return unavailable("Secure API-key storage currently supports Windows and Linux only.");
  } catch {
    return unavailable(
      "Secure credential storage could not be initialized. Restart Zotero and try again."
    );
  }
}
