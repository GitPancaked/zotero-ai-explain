import { describe, expect, it, vi } from "vitest";
import { createDefaultProviderProfileSettings } from "../../src/preferences/provider-profile.js";
import { ollamaSettingsToProfile } from "../../src/preferences/ollama-profile.js";
import { createLiveChatProvider } from "../../src/providers/live-chat-provider.js";
import type { ChatEvent, ChatRequest, ModelProvider } from "../../src/providers/provider-types.js";

describe("live chat settings", () => {
  it("switches an existing conversation from Ollama to OpenAI to Anthropic and back", async () => {
    let settings = createDefaultProviderProfileSettings();
    const localRequests: ChatRequest[] = [];
    const local: ModelProvider = {
      id: "ollama",
      displayName: "Ollama",
      async *streamChat(request) {
        await Promise.resolve();
        localRequests.push(request);
        yield { type: "message_end" };
      }
    };
    const fetch = vi.fn<(input: string, init: RequestInit) => Promise<Response>>(() =>
      Promise.resolve(new Response("", { status: 401 }))
    );
    const provider = createLiveChatProvider({
      fetch,
      readProviderProfile: () => settings,
      ollamaProvider: local
    });
    // Reuse the old conversation's request/profile after every settings save.
    const request: ChatRequest = {
      selection: {
        quote: "text",
        source: { itemKey: null, itemTitle: null, attachmentKey: null, pageLabel: null },
        anchor: null
      },
      messages: [{ role: "user", content: "Explain" }],
      profile: ollamaSettingsToProfile(settings.ollama)
    };
    const run = async () => {
      const events: ChatEvent[] = [];
      for await (const event of provider.streamChat(request, new AbortController().signal))
        events.push(event);
      return events;
    };
    await run();
    settings = {
      ...settings,
      chatProvider: "codex-api",
      openaiApiKey: "synthetic-openai",
      ollama: {
        ...settings.ollama,
        chatBaseUrl: "https://api.openai.com/v1",
        chatModel: "new-chat-model"
      }
    };
    await run();
    expect(fetch.mock.calls[0]?.[0]).toBe("https://api.openai.com/v1/chat/completions");
    const openaiInit = fetch.mock.calls[0]?.[1];
    if (openaiInit === undefined) throw new Error("OpenAI request was not sent");
    expect(new Headers(openaiInit.headers).get("Authorization")).toBe("Bearer synthetic-openai");
    expect(JSON.parse(openaiInit.body as string)).toMatchObject({ model: "new-chat-model" });
    settings = {
      ...settings,
      chatProvider: "claude-api",
      anthropicApiKey: "synthetic-anthropic",
      ollama: {
        ...settings.ollama,
        chatBaseUrl: "https://api.anthropic.com",
        chatModel: "claude-model"
      }
    };
    await run();
    expect(fetch.mock.calls[1]?.[0]).toBe("https://api.anthropic.com/v1/messages");
    const anthropicInit = fetch.mock.calls[1]?.[1];
    if (anthropicInit === undefined) throw new Error("Anthropic request was not sent");
    expect(new Headers(anthropicInit.headers).get("x-api-key")).toBe("synthetic-anthropic");
    expect(new Headers(anthropicInit.headers).get("Authorization")).toBeNull();
    settings = {
      ...settings,
      chatProvider: "ollama",
      ollama: {
        ...settings.ollama,
        chatBaseUrl: "http://localhost:11434",
        chatModel: "local-model"
      }
    };
    await run();
    expect(fetch).toHaveBeenCalledTimes(2);
    expect(localRequests).toHaveLength(2);
    expect(localRequests[1]?.profile.baseUrl).toBe("http://localhost:11434");
    expect(localRequests[1]?.profile.model).toBe("local-model");
  });

  it("does not fall back to Ollama when the newly selected API provider has no key", async () => {
    const settings = {
      ...createDefaultProviderProfileSettings(),
      chatProvider: "codex-api" as const
    };
    const fetch = vi.fn();
    const streamChat = vi.fn();
    const provider = createLiveChatProvider({
      fetch,
      readProviderProfile: () => settings,
      ollamaProvider: { id: "ollama", displayName: "Ollama", streamChat }
    });
    const request: ChatRequest = {
      selection: {
        quote: "text",
        source: { itemKey: null, itemTitle: null, attachmentKey: null, pageLabel: null },
        anchor: null
      },
      messages: [],
      profile: ollamaSettingsToProfile(settings.ollama)
    };
    const events = [];
    for await (const event of provider.streamChat(request, new AbortController().signal))
      events.push(event);
    expect(events.at(-1)?.type).toBe("error");
    expect(fetch).not.toHaveBeenCalled();
    expect(streamChat).not.toHaveBeenCalled();
  });
});
