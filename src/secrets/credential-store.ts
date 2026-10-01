export type ApiKeyProvider = "openai" | "anthropic" | "gemini";

export type CredentialStore = {
  readonly get: (provider: ApiKeyProvider) => Promise<string | null>;
  readonly set: (provider: ApiKeyProvider, secret: string) => Promise<void>;
  readonly delete: (provider: ApiKeyProvider) => Promise<void>;
};

/** Secrets must travel over stdin/stdout, never argv, logs, or temporary files. */
export type CredentialCommandRunner = (input: {
  readonly command: string;
  readonly args: readonly string[];
  readonly stdin?: string;
}) => Promise<{ readonly exitCode: number; readonly stdout: string }>;
