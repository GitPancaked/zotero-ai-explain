import { ollamaSettingsToProfile } from "../preferences/ollama-profile.js";
import type { ProviderProfileSettings } from "../preferences/provider-profile.js";
import { createClaudeApiProvider } from "./adapters/claude-api.js";
import { createOpenAIChatProvider } from "./adapters/openai-chat.js";
import type { ModelProvider } from "./provider-types.js";

/** Snapshot saved chat settings for each new stream, including existing conversations. */
export function createLiveChatProvider(deps: {
  readonly fetch: (input: string, init: RequestInit) => Promise<Response>;
  readonly readProviderProfile: () => ProviderProfileSettings;
  readonly ollamaProvider: ModelProvider;
}): ModelProvider {
  return {
    id: "configured-chat",
    displayName: "Configured chat provider",
    async *streamChat(request, signal) {
      const settings = deps.readProviderProfile();
      let provider: ModelProvider;
      switch (settings.chatProvider) {
        case "codex-api":
          provider = createOpenAIChatProvider({
            fetch: deps.fetch,
            getApiKey: () => settings.openaiApiKey || null
          });
          break;
        case "claude-api":
          provider = createClaudeApiProvider({
            fetch: deps.fetch,
            getApiKey: () => settings.anthropicApiKey || null
          });
          break;
        default:
          provider = deps.ollamaProvider;
      }
      yield* provider.streamChat(
        { ...request, profile: ollamaSettingsToProfile(settings.ollama) },
        signal
      );
    }
  };
}
