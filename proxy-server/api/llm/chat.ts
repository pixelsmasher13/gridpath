import { VercelRequest, VercelResponse } from '@vercel/node';
import { withAuth, AuthenticatedRequest } from '../../lib/middleware';
import { callLLM, LLMRequest } from '../../lib/llm-clients';

async function chatHandler(req: AuthenticatedRequest, res: VercelResponse): Promise<void> {
  if (req.method !== 'POST') {
    res.status(405).json({ error: 'Method not allowed' });
    return;
  }

  try {
    const { provider, messages, maxTokens, sessionId } = req.body;

    // Validate request
    if (!provider || !messages || !Array.isArray(messages)) {
      res.status(400).json({ error: 'Invalid request format' });
      return;
    }

    if (!['openai', 'claude', 'grok', 'gemini', 'pro', 'ultra'].includes(provider)) {
      res.status(400).json({ error: 'Unsupported provider' });
      return;
    }

    // Prepare LLM request
    const llmRequest: LLMRequest = {
      provider,
      messages,
      maxTokens: maxTokens || 200,
      sessionId: sessionId || 'default',
    };

    // Call LLM
    console.log(`Processing ${provider} request for user ${req.user.email}`);
    const response = await callLLM(llmRequest);

    // Log usage for billing (you can extend this to save to database)
    console.log(`LLM Usage - User: ${req.user.email}, Provider: ${provider}, Tokens: ${response.usage.total_tokens}`);

    // Return response
    res.status(200).json({
      success: true,
      content: response.content,
      usage: response.usage,
      provider,
    });
  } catch (error) {
    console.error('LLM proxy error:', error);
    
    // Return appropriate error message
    if (error instanceof Error) {
      if (error.message.includes('API key')) {
        res.status(500).json({ error: 'LLM service configuration error' });
      } else if (error.message.includes('rate limit')) {
        res.status(429).json({ error: 'Rate limit exceeded' });
      } else {
        res.status(500).json({ error: 'LLM request failed' });
      }
    } else {
      res.status(500).json({ error: 'Unknown error occurred' });
    }
  }
}

export default withAuth(chatHandler); 