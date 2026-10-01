import { describe, expect, it, vi } from "vitest";
import {
  createCredentialCommandRunner,
  type CredentialProcess,
  type CredentialSubprocess
} from "../../src/platform/credential-command.js";

function fixture() {
  const child: CredentialProcess = {
    stdin: { write: vi.fn(() => Promise.resolve()), close: vi.fn(() => Promise.resolve()) },
    stdout: { read: vi.fn().mockResolvedValue(new ArrayBuffer(0)) },
    stderr: { read: vi.fn().mockResolvedValue(new ArrayBuffer(0)) },
    wait: vi.fn(() => Promise.resolve({ exitCode: 0 })),
    kill: vi.fn(() => Promise.resolve())
  };
  const subprocess: CredentialSubprocess = {
    call: vi.fn(() => Promise.resolve(child))
  };
  return { child, subprocess };
}

describe("credential subprocess transport", () => {
  it("writes only over stdin and drains both streams before returning", async () => {
    const f = fixture();
    vi.mocked(f.child.stdout.read).mockResolvedValueOnce(
      new TextEncoder().encode("retrieved-key").buffer
    );
    vi.mocked(f.child.stderr.read).mockResolvedValueOnce(
      new TextEncoder().encode("sensitive error").buffer
    );
    const result = await createCredentialCommandRunner(f.subprocess)({
      command: "/trusted/helper",
      args: ["get"],
      stdin: "new-key"
    });
    expect(result).toEqual({ exitCode: 0, stdout: "retrieved-key" });
    expect(f.subprocess.call).toHaveBeenCalledWith({
      command: "/trusted/helper",
      arguments: ["get"],
      stderr: "pipe"
    });
    expect(f.child.stdin.write).toHaveBeenCalledWith("new-key");
    expect(f.child.stdin.close).toHaveBeenCalled();
  });

  it("kills timed-out helpers without exposing their errors", async () => {
    const f = fixture();
    vi.mocked(f.child.wait).mockReturnValue(new Promise(() => undefined));
    await expect(
      createCredentialCommandRunner(f.subprocess, 10)({ command: "/helper", args: [] })
    ).rejects.toThrow("Secure credential storage is unavailable");
    expect(f.child.kill).toHaveBeenCalledWith(0);
  });

  it("kills a child that launches after the timeout", async () => {
    const f = fixture();
    let launch: ((child: CredentialProcess) => void) | undefined;
    vi.mocked(f.subprocess.call).mockReturnValue(
      new Promise((resolve) => {
        launch = resolve;
      })
    );
    await expect(
      createCredentialCommandRunner(f.subprocess, 10)({ command: "/helper", args: [] })
    ).rejects.toThrow();
    launch?.(f.child);
    await Promise.resolve();
    await Promise.resolve();
    expect(f.child.kill).toHaveBeenCalledWith(0);
    expect(f.child.stdin.write).not.toHaveBeenCalled();
  });

  it("bounds output and kills a helper on transport failure", async () => {
    const f = fixture();
    vi.mocked(f.child.stdout.read).mockResolvedValueOnce(new ArrayBuffer(65537));
    await expect(
      createCredentialCommandRunner(f.subprocess)({ command: "/helper", args: [] })
    ).rejects.toThrow();
    expect(f.child.kill).toHaveBeenCalledWith(0);
  });
});
