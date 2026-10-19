/**
 * The server half of the demo's agent loop: one Claude call per request.
 * The key stays here (ANTHROPIC_API_KEY, or an `ant auth login` profile);
 * the browser sends the conversation, runs the model's tool calls against
 * the live grid with `sheet.ai.run()`, and posts the results back.
 *
 * The model, system prompt and tool list are fixed on the server. Taking
 * any of them from the request would turn this route into a general
 * proxy for your key. In a real app, also check the user's session and
 * rate-limit here.
 */
import Anthropic from "@anthropic-ai/sdk";
import { AGENT_INSTRUCTIONS, agentTools } from "gridpath/agent";

let client: Anthropic | null = null;
const tools = agentTools() as Anthropic.Beta.BetaTool[];

export type AgentTurn = { content: Anthropic.Beta.BetaContentBlock[]; stop_reason: string | null };

export async function agentTurn(body: unknown): Promise<AgentTurn> {
  const messages = (body as { messages?: Anthropic.Beta.BetaMessageParam[] } | null)?.messages;
  if (!Array.isArray(messages) || messages.length === 0) throw Object.assign(new Error("messages required"), { status: 400 });
  client ??= new Anthropic();
  const response = await client.beta.messages.create({
    model: "claude-opus-5-5",
    max_tokens: 16000,
    output_config: { effort: "high" },
    system: AGENT_INSTRUCTIONS,
    tools,
    // The system prompt and tools are identical on every call: cache them.
    cache_control: { type: "ephemeral" },
    // A declined request is re-run on a fallback model in the same call.
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default",
    messages,
  });
  return { content: response.content, stop_reason: response.stop_reason };
}
