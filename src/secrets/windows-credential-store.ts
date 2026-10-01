import type {
  ApiKeyProvider,
  CredentialCommandRunner,
  CredentialStore
} from "./credential-store.js";

// A fixed script keeps provider names and secrets out of executable command text.
// CredWriteW uses generic credentials persisted for this user on this machine.
const SCRIPT = String.raw`
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
try {
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class ZoteroCredentials {
  [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
  public struct Credential {
    public uint Flags, Type;
    public string TargetName, Comment;
    public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
    public uint CredentialBlobSize;
    public IntPtr CredentialBlob;
    public uint Persist, AttributeCount;
    public IntPtr Attributes;
    public string TargetAlias, UserName;
  }
  [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool CredReadW(string target, uint type, uint flags, out IntPtr credential);
  [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool CredWriteW(ref Credential credential, uint flags);
  [DllImport("advapi32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
  public static extern bool CredDeleteW(string target, uint type, uint flags);
  [DllImport("advapi32.dll")]
  public static extern void CredFree(IntPtr credential);
}
'@
$request = [Console]::In.ReadToEnd() | ConvertFrom-Json
if ($request.provider -notin @('openai','anthropic','gemini')) { throw 'Invalid provider' }
$target = 'zotero-ai-explain/' + $request.provider
switch ($request.operation) {
  'get' {
    $pointer = [IntPtr]::Zero
    if (-not [ZoteroCredentials]::CredReadW($target, 1, 0, [ref]$pointer)) {
      if ([Runtime.InteropServices.Marshal]::GetLastWin32Error() -eq 1168) {
        [Console]::Out.Write('{"secret":null}')
      } else { throw 'Read failed' }
    } else {
      try {
        $credential = [Runtime.InteropServices.Marshal]::PtrToStructure($pointer, [type][ZoteroCredentials+Credential])
        if (($credential.CredentialBlobSize % 2) -ne 0) { throw 'Invalid credential' }
        $secret = [Runtime.InteropServices.Marshal]::PtrToStringUni($credential.CredentialBlob, [int]($credential.CredentialBlobSize / 2))
        [Console]::Out.Write((@{secret=$secret} | ConvertTo-Json -Compress))
      } finally { [ZoteroCredentials]::CredFree($pointer) }
    }
  }
  'set' {
    if ($request.secret -isnot [string] -or $request.secret.Length -eq 0) { throw 'Invalid secret' }
    $bytes = [Text.Encoding]::Unicode.GetBytes($request.secret)
    if ($bytes.Length -gt 2560) { throw 'Secret too long' }
    $blob = [Runtime.InteropServices.Marshal]::AllocHGlobal($bytes.Length)
    try {
      [Runtime.InteropServices.Marshal]::Copy($bytes, 0, $blob, $bytes.Length)
      $credential = New-Object ZoteroCredentials+Credential
      $credential.Type = 1
      $credential.TargetName = $target
      $credential.UserName = 'Zotero AI Explain'
      $credential.CredentialBlobSize = $bytes.Length
      $credential.CredentialBlob = $blob
      $credential.Persist = 2
      if (-not [ZoteroCredentials]::CredWriteW([ref]$credential, 0)) { throw 'Write failed' }
      [Console]::Out.Write('{}')
    } finally {
      for ($i = 0; $i -lt $bytes.Length; $i++) { [Runtime.InteropServices.Marshal]::WriteByte($blob, $i, 0) }
      [Runtime.InteropServices.Marshal]::FreeHGlobal($blob)
      [Array]::Clear($bytes, 0, $bytes.Length)
    }
  }
  'delete' {
    if (-not [ZoteroCredentials]::CredDeleteW($target, 1, 0)) {
      if ([Runtime.InteropServices.Marshal]::GetLastWin32Error() -ne 1168) { throw 'Delete failed' }
    }
    [Console]::Out.Write('{}')
  }
  default { throw 'Invalid operation' }
}
} catch {
  [Console]::Error.Write('Windows credential storage failed.')
  exit 1
}
`;

function encodePowerShellScript(script: string): string {
  const alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  const bytes = Array.from(script).flatMap((character) => [character.charCodeAt(0), 0]);
  let encoded = "";
  for (let index = 0; index < bytes.length; index += 3) {
    const first = bytes[index] ?? 0;
    const second = bytes[index + 1] ?? 0;
    const third = bytes[index + 2] ?? 0;
    encoded += alphabet.charAt(first >> 2);
    encoded += alphabet.charAt(((first & 3) << 4) | (second >> 4));
    encoded +=
      index + 1 < bytes.length ? alphabet.charAt(((second & 15) << 2) | (third >> 6)) : "=";
    encoded += index + 2 < bytes.length ? alphabet.charAt(third & 63) : "=";
  }
  return encoded;
}

const ENCODED_SCRIPT = encodePowerShellScript(SCRIPT);

export function createWindowsCredentialStore(run: CredentialCommandRunner): CredentialStore {
  async function request(
    operation: "get" | "set" | "delete",
    provider: ApiKeyProvider,
    secret?: string
  ): Promise<{ secret?: string | null }> {
    try {
      if (!["openai", "anthropic", "gemini"].includes(provider)) {
        throw new Error();
      }
      if (operation === "set" && (!secret || secret.length * 2 > 2560)) {
        throw new Error();
      }
      const result = await run({
        command: "powershell.exe",
        args: [
          "-NoLogo",
          "-NoProfile",
          "-NonInteractive",
          "-WindowStyle",
          "Hidden",
          "-EncodedCommand",
          ENCODED_SCRIPT
        ],
        stdin: JSON.stringify({ operation, provider, secret })
      });
      if (result.exitCode !== 0) throw new Error();
      const response: unknown = JSON.parse(result.stdout);
      if (typeof response !== "object" || response === null) throw new Error();
      if (operation === "get") {
        const value = (response as { secret?: unknown }).secret;
        if (value !== null && typeof value !== "string") throw new Error();
        return { secret: value };
      }
      return {};
    } catch {
      throw new Error(
        "Windows Credential Manager is unavailable or the credential operation failed."
      );
    }
  }
  return {
    async get(provider) {
      return (await request("get", provider)).secret ?? null;
    },
    async set(provider, secret) {
      await request("set", provider, secret);
    },
    async delete(provider) {
      await request("delete", provider);
    }
  };
}
