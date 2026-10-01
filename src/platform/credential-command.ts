import type { CredentialCommandRunner } from "../secrets/credential-store.js";

export type CredentialProcess = {
  readonly stdin: {
    readonly write: (value: string) => Promise<unknown>;
    readonly close: () => Promise<unknown>;
  };
  readonly stdout: { readonly read: () => Promise<ArrayBuffer> };
  readonly stderr: { readonly read: () => Promise<ArrayBuffer> };
  readonly wait: () => Promise<{ readonly exitCode: number }>;
  readonly kill: (timeout: number) => Promise<unknown>;
};

export type CredentialSubprocess = {
  readonly call: (options: {
    readonly command: string;
    readonly arguments: readonly string[];
    readonly stderr: "pipe";
  }) => Promise<CredentialProcess>;
};

export function createCredentialCommandRunner(
  subprocess: CredentialSubprocess,
  timeoutMs = 30000
): CredentialCommandRunner {
  return async (input) => {
    let child: CredentialProcess | undefined;
    let cancelled = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const work = async () => {
      child = await subprocess.call({
        command: input.command,
        arguments: input.args,
        stderr: "pipe"
      });
      const proc = child;
      if (cancelled) {
        await proc.kill(0);
        throw new Error("Credential operation cancelled.");
      }
      const drain = async (pipe: CredentialProcess["stdout"], keep: boolean) => {
        const decoder = new TextDecoder();
        let result = "";
        let size = 0;
        for (;;) {
          const chunk = await pipe.read();
          if (chunk.byteLength === 0) break;
          size += chunk.byteLength;
          if (size > 65536) throw new Error("Credential helper output exceeded its limit.");
          if (keep) result += decoder.decode(chunk, { stream: true });
        }
        return keep ? result + decoder.decode() : "";
      };
      const [stdout, , , status] = await Promise.all([
        drain(proc.stdout, true),
        drain(proc.stderr, false),
        (async () => {
          if (input.stdin !== undefined) await proc.stdin.write(input.stdin);
          await proc.stdin.close();
        })(),
        proc.wait()
      ]);
      return { exitCode: status.exitCode, stdout };
    };
    try {
      return await Promise.race([
        work(),
        new Promise<never>((_, reject) => {
          timer = setTimeout(() => {
            reject(new Error("Credential helper timed out."));
          }, timeoutMs);
        })
      ]);
    } catch {
      cancelled = true;
      if (child !== undefined) await child.kill(0).catch(() => undefined);
      // Native error text may contain stdin or credential output. Never expose it.
      throw new Error(
        "Secure credential storage is unavailable. Unlock your keyring and try again."
      );
    } finally {
      if (timer !== undefined) clearTimeout(timer);
    }
  };
}
