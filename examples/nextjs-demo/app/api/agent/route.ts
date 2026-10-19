// The server half of an agent loop (see the gridpath README): one Claude
// call per request, with the system prompt and tools fixed here.
import Anthropic from "@anthropic-ai/sdk";
import { AGENT_INSTRUCTIONS, agentTools } from "gridpath/agent";

const tools = agentTools() as Anthropic.Beta.BetaTool[];
let client: Anthropic | null = null;

export async function POST(req: Request) {
  const { messages } = (await req.json()) as { messages?: Anthropic.Beta.BetaMessageParam[] };
  if (!Array.isArray(messages) || messages.length === 0) return Response.json({ error: "messages required" }, { status: 400 });
  try {
    client ??= new Anthropic();
    const response = await client.beta.messages.create({
      model: "claude-opus-5-5",
      max_tokens: 16000,
      output_config: { effort: "high" },
      system: AGENT_INSTRUCTIONS,
      tools,
      cache_control: { type: "ephemeral" },
      betas: ["server-side-fallback-2026-07-01"],
      fallbacks: "default",
      messages,
    });
    return Response.json({ content: response.content, stop_reason: response.stop_reason });
  } catch (e) {
    return Response.json({ error: e instanceof Error ? e.message : String(e) }, { status: 500 });
  }
}
