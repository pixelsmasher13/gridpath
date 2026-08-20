import OpenAI from 'openai';
import Anthropic from '@anthropic-ai/sdk';
// @ts-ignore — subpath exports require moduleResolution "bundler" or "node16"; Vercel bundles correctly at deploy
import { createVertex } from '@ai-sdk/google-vertex/edge';
// @ts-ignore
import { generateText } from 'ai';

export interface LLMMessage {
  role: 'system' | 'user' | 'assistant';
  content: string;
}

export interface LLMResponse {
  content: string;
  usage: {
    prompt_tokens: number;
    completion_tokens: number;
    total_tokens: number;
  };
}

export interface LLMRequest {
  provider: 'openai' | 'claude' | 'grok' | 'gemini' | 'pro' | 'ultra';
  messages: LLMMessage[];
  maxTokens: number;
  sessionId: string;
}

// Initialize clients
let openaiClient: OpenAI | null = null;
let anthropicClient: Anthropic | null = null;

function getOpenAIClient(): OpenAI {
  if (!openaiClient) {
    const apiKey = process.env.OPENAI_API_KEY;
    if (!apiKey) {
      throw new Error('OPENAI_API_KEY environment variable is required');
    }
    openaiClient = new OpenAI({ apiKey });
  }
  return openaiClient;
}

function getAnthropicClient(): Anthropic {
  if (!anthropicClient) {
    const apiKey = process.env.CLAUDE_API_KEY;
    if (!apiKey) {
      throw new Error('CLAUDE_API_KEY environment variable is required');
    }
    anthropicClient = new Anthropic({ apiKey });
  }
  return anthropicClient;
}

export async function callOpenAI(messages: LLMMessage[], maxTokens: number): Promise<LLMResponse> {
  const client = getOpenAIClient();
  
  const completion = await client.chat.completions.create({
    model: 'gpt-5.4',
    messages: messages as any,
    max_completion_tokens: maxTokens,
    temperature: 0.1,
  });

  const choice = completion.choices[0];
  if (!choice?.message?.content) {
    throw new Error('No response from OpenAI');
  }

  return {
    content: choice.message.content,
    usage: {
      prompt_tokens: completion.usage?.prompt_tokens || 0,
      completion_tokens: completion.usage?.completion_tokens || 0,
      total_tokens: completion.usage?.total_tokens || 0,
    },
  };
}

// Anthropic call shared by callClaude, callProModel, and callUltraModel.
// Parametrized on model so all three benefit from the same ephemeral system-prompt cache.
async function callAnthropicModel(
  messages: LLMMessage[],
  maxTokens: number,
  model: string
): Promise<LLMResponse> {
  const client = getAnthropicClient();

  // Convert messages to Claude format
  const systemMessage = messages.find(m => m.role === 'system')?.content || '';
  const conversationMessages = messages.filter(m => m.role !== 'system');

  // Mark the system prompt as cacheable. Anthropic auto-hits on prefix match
  // within the 5-min ephemeral TTL (auto-refreshed on every hit).
  // SDK 0.17 types predate cache_control on system — cast bypasses the old types;
  // the runtime API supports this shape fine.
  const systemField: any = systemMessage
    ? [{ type: 'text', text: systemMessage, cache_control: { type: 'ephemeral' } }]
    : undefined;

  const response = await client.messages.create({
    model,
    max_tokens: maxTokens,
    system: systemField,
    messages: conversationMessages.map(msg => ({
      role: msg.role as 'user' | 'assistant',
      content: msg.content,
    })),
  });

  const content = response.content[0];
  if (content.type !== 'text') {
    throw new Error('Unexpected response type from Claude');
  }

  // Surface cache telemetry so we can see hit rate in logs post-deploy
  const cacheWrite = (response.usage as any).cache_creation_input_tokens || 0;
  const cacheRead = (response.usage as any).cache_read_input_tokens || 0;
  if (cacheWrite > 0 || cacheRead > 0) {
    console.log(`⚡ ${model} cache: ${cacheRead} read, ${cacheWrite} write, ${response.usage.input_tokens} total input`);
  }

  return {
    content: content.text,
    usage: {
      prompt_tokens: response.usage.input_tokens,
      completion_tokens: response.usage.output_tokens,
      total_tokens: response.usage.input_tokens + response.usage.output_tokens,
    },
  };
}

export async function callClaude(messages: LLMMessage[], maxTokens: number): Promise<LLMResponse> {
  return callAnthropicModel(messages, maxTokens, 'claude-sonnet-4-6');
}

export async function callGrok(messages: LLMMessage[], maxTokens: number): Promise<LLMResponse> {
  const apiKey = process.env.GROK_API_KEY;
  console.log('Grok API Key present:', !!apiKey, apiKey ? `${apiKey.substring(0, 10)}...` : 'MISSING');

  if (!apiKey) {
    throw new Error('GROK_API_KEY environment variable is required');
  }

  const requestBody = {
    model: 'grok-3',
    messages: messages,
    max_tokens: maxTokens,
    temperature: 0.1,
  };

  console.log('Grok request:', JSON.stringify(requestBody, null, 2));

  const response = await fetch('https://api.x.ai/v1/chat/completions', {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${apiKey}`,
    },
    body: JSON.stringify(requestBody),
  });

  console.log('Grok response status:', response.status, response.statusText);

  if (!response.ok) {
    const errorText = await response.text();
    console.error('Grok API error details:', {
      status: response.status,
      statusText: response.statusText,
      headers: Object.fromEntries(response.headers.entries()),
      body: errorText,
      requestUrl: 'https://api.x.ai/v1/chat/completions',
      apiKeyPresent: !!apiKey,
      apiKeyPrefix: apiKey ? apiKey.substring(0, 4) : 'NONE'
    });
    throw new Error(`Grok API error: ${response.status} ${response.statusText} - ${errorText}`);
  }

  const data: any = await response.json();
  console.log('Grok response data:', JSON.stringify(data, null, 2));

  const choice = data.choices?.[0];

  if (!choice?.message?.content) {
    throw new Error('No response from Grok - invalid response structure');
  }

  return {
    content: choice.message.content,
    usage: {
      prompt_tokens: data.usage?.prompt_tokens || 0,
      completion_tokens: data.usage?.completion_tokens || 0,
      total_tokens: data.usage?.total_tokens || 0,
    },
  };
}

// Vertex AI provider (lazy-initialized, only when GOOGLE_VERTEX_API_KEY is set)
let vertexProvider: ReturnType<typeof createVertex> | null = null;

function getVertexProvider() {
  if (!vertexProvider) {
    const apiKey = process.env.GOOGLE_VERTEX_API_KEY;
    if (!apiKey) return null;
    vertexProvider = createVertex({ apiKey });
  }
  return vertexProvider;
}

async function callGeminiViaVertex(messages: LLMMessage[], maxTokens: number, model: string, thinkingBudget: number = 1024): Promise<LLMResponse> {
  const vertex = getVertexProvider();
  if (!vertex) throw new Error('GOOGLE_VERTEX_API_KEY not configured');

  const systemMessage = messages.find(m => m.role === 'system')?.content;
  const conversationMessages = messages
    .filter(m => m.role !== 'system')
    .map(msg => ({ role: msg.role as 'user' | 'assistant', content: msg.content }));

  const result = await generateText({
    model: vertex(model),
    system: systemMessage,
    messages: conversationMessages,
    maxOutputTokens: maxTokens,
    temperature: 0.1,
    providerOptions: {
      vertex: {
        thinkingConfig: {
          thinkingBudget,
        },
      },
    },
  });

  const inputTokens = result.usage?.inputTokens || 0;
  const outputTokens = result.usage?.outputTokens || 0;

  return {
    content: result.text,
    usage: {
      prompt_tokens: inputTokens,
      completion_tokens: outputTokens,
      total_tokens: inputTokens + outputTokens,
    },
  };
}

async function callGeminiViaAIStudio(messages: LLMMessage[], maxTokens: number, model: string, thinkingBudget: number = 1024): Promise<LLMResponse> {
  const apiKey = process.env.GEMINI_API_KEY;

  if (!apiKey) {
    throw new Error('GEMINI_API_KEY environment variable is required');
  }

  const systemMessage = messages.find(m => m.role === 'system');
  const conversationMessages = messages.filter(m => m.role !== 'system');

  const contents = conversationMessages.map(msg => ({
    role: msg.role === 'assistant' ? 'model' : 'user',
    parts: [{ text: msg.content }]
  }));

  const requestBody: any = {
    contents,
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: maxTokens,
      thinkingConfig: {
        thinkingBudget,
      },
    },
  };

  if (systemMessage) {
    requestBody.systemInstruction = {
      parts: [{ text: systemMessage.content }]
    };
  }

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(requestBody),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Gemini API error (${model}): ${response.status} - ${errorText}`);
  }

  const data: any = await response.json();

  if (!data.candidates || data.candidates.length === 0) {
    console.error('Gemini response structure:', JSON.stringify(data, null, 2));
    if (data.promptFeedback?.blockReason) {
      throw new Error(`Gemini blocked the prompt: ${data.promptFeedback.blockReason}`);
    }
    throw new Error('Gemini API returned no candidates - may have hit safety filters or rate limits');
  }

  const candidate = data.candidates[0];

  if (candidate.finishReason === 'SAFETY') {
    throw new Error('Gemini response blocked by safety filters');
  }

  if (candidate.finishReason === 'MAX_TOKENS') {
    console.warn('Gemini response was truncated due to token limit');
  }

  if (!candidate.content?.parts || candidate.content.parts.length === 0) {
    console.error('Gemini candidate structure:', JSON.stringify(candidate, null, 2));
    throw new Error('Gemini API returned response with no content parts');
  }

  const responseText = candidate.content.parts[0]?.text;
  if (!responseText) {
    throw new Error('Gemini API returned empty text response');
  }

  const usageMetadata = data.usageMetadata;
  const promptTokens = usageMetadata?.promptTokenCount ||
    Math.ceil(messages.reduce((sum, msg) => sum + msg.content.length, 0) / 4);
  const completionTokens = usageMetadata?.candidatesTokenCount ||
    usageMetadata?.totalTokenCount ? (usageMetadata.totalTokenCount - (usageMetadata?.promptTokenCount || 0)) :
    Math.ceil(responseText.length / 4);

  return {
    content: responseText,
    usage: {
      prompt_tokens: promptTokens,
      completion_tokens: completionTokens,
      total_tokens: promptTokens + completionTokens,
    },
  };
}

// ─── Gemini Explicit Context Caching (AI Studio) ───
// Caches the system prompt once per execution so it isn't re-processed on every step.
// Requires GEMINI_API_KEY. Min cached content: 1024 tokens (Flash models).
const GEMINI_CACHE_MODEL = 'gemini-3-flash-preview';

export async function createGeminiCache(
  systemPromptText: string,
  ttlSeconds: number = 1800
): Promise<string | null> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    console.warn('⚡ Gemini caching: GEMINI_API_KEY not set, skipping');
    return null;
  }

  try {
    const response = await fetch(
      `https://generativelanguage.googleapis.com/v1beta/cachedContents?key=${apiKey}`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          model: `models/${GEMINI_CACHE_MODEL}`,
          systemInstruction: {
            parts: [{ text: systemPromptText }]
          },
          ttl: `${ttlSeconds}s`,
        }),
      }
    );

    if (!response.ok) {
      const errorText = await response.text();
      console.warn(`⚡ Gemini cache creation failed (${response.status}): ${errorText}`);
      return null;
    }

    const data: any = await response.json();
    const cacheName = data.name;
    console.log(`⚡ Gemini cache created: ${cacheName} (TTL: ${ttlSeconds}s)`);
    return cacheName;
  } catch (error: any) {
    console.warn(`⚡ Gemini cache creation error: ${error.message}`);
    return null;
  }
}

export async function callGeminiWithCache(
  cacheName: string,
  userMessages: LLMMessage[],
  maxTokens: number
): Promise<LLMResponse> {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) throw new Error('GEMINI_API_KEY required for cached calls');

  const contents = userMessages
    .filter(m => m.role !== 'system')
    .map(msg => ({
      role: msg.role === 'assistant' ? 'model' : 'user',
      parts: [{ text: msg.content }]
    }));

  const requestBody: any = {
    contents,
    cachedContent: cacheName,
    generationConfig: {
      temperature: 0.1,
      maxOutputTokens: maxTokens,
    },
  };

  const response = await fetch(
    `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_CACHE_MODEL}:generateContent?key=${apiKey}`,
    {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(requestBody),
    }
  );

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Gemini cached call error (${response.status}): ${errorText}`);
  }

  const data: any = await response.json();

  if (!data.candidates || data.candidates.length === 0) {
    if (data.promptFeedback?.blockReason) {
      throw new Error(`Gemini blocked the prompt: ${data.promptFeedback.blockReason}`);
    }
    throw new Error('Gemini API returned no candidates');
  }

  const candidate = data.candidates[0];
  if (candidate.finishReason === 'SAFETY') {
    throw new Error('Gemini response blocked by safety filters');
  }

  const responseText = candidate.content?.parts?.[0]?.text;
  if (!responseText) throw new Error('Gemini API returned empty text response');

  const usageMetadata = data.usageMetadata;
  const promptTokens = usageMetadata?.promptTokenCount || 0;
  const completionTokens = usageMetadata?.candidatesTokenCount || 0;
  const cachedTokens = usageMetadata?.cachedContentTokenCount || 0;

  if (cachedTokens > 0) {
    console.log(`⚡ Cache hit: ${cachedTokens} cached tokens, ${promptTokens - cachedTokens} new tokens`);
  }

  return {
    content: responseText,
    usage: {
      prompt_tokens: promptTokens,
      completion_tokens: completionTokens,
      total_tokens: promptTokens + completionTokens,
    },
  };
}

export async function deleteGeminiCache(cacheName: string | null): Promise<void> {
  if (!cacheName) return;
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) return;

  try {
    await fetch(
      `https://generativelanguage.googleapis.com/v1beta/${cacheName}?key=${apiKey}`,
      { method: 'DELETE' }
    );
    console.log(`⚡ Gemini cache deleted: ${cacheName}`);
  } catch (error: any) {
    console.warn(`⚡ Gemini cache deletion failed (non-fatal): ${error.message}`);
  }
}

async function callGeminiWithModel(messages: LLMMessage[], maxTokens: number, model: string, thinkingBudget: number = 1024): Promise<LLMResponse> {
  // Try Vertex AI first (lower latency), fall back to AI Studio
  if (process.env.GOOGLE_VERTEX_API_KEY) {
    try {
      return await callGeminiViaVertex(messages, maxTokens, model, thinkingBudget);
    } catch (e: any) {
      console.warn(`Vertex AI failed, falling back to AI Studio: ${e.message}`);
    }
  }
  return callGeminiViaAIStudio(messages, maxTokens, model, thinkingBudget);
}

export async function callGemini(messages: LLMMessage[], maxTokens: number, thinkingBudget: number = 1024): Promise<LLMResponse> {
  return callGeminiWithModel(messages, maxTokens, 'gemini-3-flash-preview', thinkingBudget);
}

// Lightweight model for simple classification tasks; thinking disabled (0 = off)
export async function callGeminiLite(messages: LLMMessage[], maxTokens: number): Promise<LLMResponse> {
  return callGeminiWithModel(messages, maxTokens, 'gemini-3.1-flash-lite-preview', 0);
}

// "Pro" tier — Claude Sonnet 4.6.
export async function callProModel(messages: LLMMessage[], maxTokens: number): Promise<LLMResponse> {
  return callAnthropicModel(messages, maxTokens, 'claude-sonnet-4-6');
}

// "Ultra" tier — Claude Opus 4.7. For the hardest reasoning steps.
export async function callUltraModel(messages: LLMMessage[], maxTokens: number): Promise<LLMResponse> {
  return callAnthropicModel(messages, maxTokens, 'claude-opus-4-7');
}

export async function callLLM(request: LLMRequest): Promise<LLMResponse> {
  const { provider, messages, maxTokens } = request;

  switch (provider) {
    case 'openai':
      return callOpenAI(messages, maxTokens);
    case 'claude':
      return callClaude(messages, maxTokens);
    case 'grok':
      return callGrok(messages, maxTokens);
    case 'gemini':
      return callGemini(messages, maxTokens);
    case 'pro':
      return callProModel(messages, maxTokens);
    case 'ultra':
      return callUltraModel(messages, maxTokens);
    default:
      throw new Error(`Unsupported LLM provider: ${provider}`);
  }
}