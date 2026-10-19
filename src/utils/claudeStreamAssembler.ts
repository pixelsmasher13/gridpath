/**
 * Claude CLI Stream Assembler
 * 
 * Manages state for Claude CLI streaming output, accumulating thinking,
 * content, and tool calls as events arrive from the Rust backend.
 */

// Types matching the Rust ClaudeOutputType enum
export type ClaudeOutputType =
  | 'thinking'
  | 'thinking_delta'
  | 'content'
  | 'content_delta'
  | 'tool_start'
  | 'tool_input_delta'
  | 'tool_end'
  | 'tool_result'
  | 'message_start'
  | 'message_end'
  | 'error'
  | 'unknown';

// Tool call status
export type ToolCallStatus = 'started' | 'running' | 'completed' | 'error';

// Tool call state
export interface ToolCall {
  id: string;
  name: string;
  input?: unknown;
  status: ToolCallStatus;
  result?: string;
  startedAt: number;
  updatedAt: number;
}

// Claude stream event from backend
export interface ClaudeStreamEvent {
  process_id: string;
  output_type: ClaudeOutputType;
  block_index?: number;
  content: string;
  tool_name?: string;
  tool_id?: string;
  tool_input?: unknown;
  is_delta: boolean;
  timestamp: number;
}

// Specific event types
export interface ClaudeThinkingEvent {
  process_id: string;
  thinking: string;
  is_delta: boolean;
  timestamp: number;
}

export interface ClaudeContentEvent {
  process_id: string;
  content: string;
  is_delta: boolean;
  timestamp: number;
}

export interface ClaudeToolEvent {
  process_id: string;
  tool_id: string;
  tool_name: string;
  status: string;
  input?: unknown;
  result?: string;
  timestamp: number;
}

export interface ClaudeCliDetectedEvent {
  process_id: string;
  command: string;
  timestamp: number;
}

// Display state for the UI
export interface ClaudeDisplayState {
  isClaudeCli: boolean;
  thinking: string;
  content: string;
  toolCalls: Map<string, ToolCall>;
  hasError: boolean;
  errorMessage?: string;
  isComplete: boolean;
  lastUpdated: number;
}

/**
 * Assembles streaming Claude CLI output into a coherent state
 */
export class ClaudeStreamAssembler {
  private processId: string;
  private state: ClaudeDisplayState;
  private listeners: Set<(state: ClaudeDisplayState) => void> = new Set();

  constructor(processId: string) {
    this.processId = processId;
    this.state = this.createInitialState();
  }

  private createInitialState(): ClaudeDisplayState {
    return {
      isClaudeCli: false,
      thinking: '',
      content: '',
      toolCalls: new Map(),
      hasError: false,
      errorMessage: undefined,
      isComplete: false,
      lastUpdated: Date.now(),
    };
  }

  /**
   * Get the current display state
   */
  getState(): ClaudeDisplayState {
    return { ...this.state, toolCalls: new Map(this.state.toolCalls) };
  }

  /**
   * Subscribe to state changes
   */
  subscribe(listener: (state: ClaudeDisplayState) => void): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  /**
   * Notify all listeners of state change
   */
  private notifyListeners() {
    const state = this.getState();
    this.listeners.forEach(listener => listener(state));
  }

  /**
   * Handle Claude CLI detected event
   */
  handleCliDetected(event: ClaudeCliDetectedEvent): void {
    if (event.process_id !== this.processId) return;
    
    this.state.isClaudeCli = true;
    this.state.lastUpdated = event.timestamp;
    this.notifyListeners();
  }

  /**
   * Handle thinking update event
   */
  handleThinking(event: ClaudeThinkingEvent): void {
    if (event.process_id !== this.processId) return;

    if (event.is_delta) {
      this.state.thinking += event.thinking;
    } else {
      // If not a delta, it's the full thinking text so far
      // But we accumulate deltas, so just append
      this.state.thinking += event.thinking;
    }
    this.state.lastUpdated = event.timestamp;
    this.notifyListeners();
  }

  /**
   * Handle content update event
   */
  handleContent(event: ClaudeContentEvent): void {
    if (event.process_id !== this.processId) return;

    if (event.is_delta) {
      this.state.content += event.content;
    } else {
      this.state.content += event.content;
    }
    this.state.lastUpdated = event.timestamp;
    this.notifyListeners();
  }

  /**
   * Handle tool event
   */
  handleTool(event: ClaudeToolEvent): void {
    if (event.process_id !== this.processId) return;

    const existingTool = this.state.toolCalls.get(event.tool_id);
    
    const toolCall: ToolCall = {
      id: event.tool_id,
      name: event.tool_name,
      input: event.input ?? existingTool?.input,
      status: event.status as ToolCallStatus,
      result: event.result ?? existingTool?.result,
      startedAt: existingTool?.startedAt ?? event.timestamp,
      updatedAt: event.timestamp,
    };

    this.state.toolCalls.set(event.tool_id, toolCall);
    this.state.lastUpdated = event.timestamp;
    this.notifyListeners();
  }

  /**
   * Handle generic stream event
   */
  handleStreamEvent(event: ClaudeStreamEvent): void {
    if (event.process_id !== this.processId) return;

    this.state.isClaudeCli = true;

    switch (event.output_type) {
      case 'thinking':
      case 'thinking_delta':
        if (event.content) {
          this.state.thinking += event.content;
        }
        break;

      case 'content':
      case 'content_delta':
        if (event.content) {
          this.state.content += event.content;
        }
        break;

      case 'tool_start':
        if (event.tool_id && event.tool_name) {
          this.state.toolCalls.set(event.tool_id, {
            id: event.tool_id,
            name: event.tool_name,
            input: event.tool_input,
            status: 'started',
            startedAt: event.timestamp,
            updatedAt: event.timestamp,
          });
        }
        break;

      case 'tool_end':
        if (event.tool_id) {
          const tool = this.state.toolCalls.get(event.tool_id);
          if (tool) {
            tool.status = 'running';
            tool.input = event.tool_input ?? tool.input;
            tool.updatedAt = event.timestamp;
          }
        }
        break;

      case 'tool_result':
        if (event.tool_id) {
          const tool = this.state.toolCalls.get(event.tool_id);
          if (tool) {
            tool.status = 'completed';
            tool.result = event.content;
            tool.updatedAt = event.timestamp;
          }
        }
        break;

      case 'message_start':
        // Reset for new message
        this.state.thinking = '';
        this.state.content = '';
        this.state.toolCalls.clear();
        this.state.isComplete = false;
        break;

      case 'message_end':
        this.state.isComplete = true;
        break;

      case 'error':
        this.state.hasError = true;
        this.state.errorMessage = event.content;
        break;
    }

    this.state.lastUpdated = event.timestamp;
    this.notifyListeners();
  }

  /**
   * Reset the assembler state
   */
  reset(): void {
    this.state = this.createInitialState();
    this.notifyListeners();
  }

  /**
   * Mark a tool as completed with result
   */
  completeTool(toolId: string, result: string): void {
    const tool = this.state.toolCalls.get(toolId);
    if (tool) {
      tool.status = 'completed';
      tool.result = result;
      tool.updatedAt = Date.now();
      this.state.lastUpdated = Date.now();
      this.notifyListeners();
    }
  }

  /**
   * Mark a tool as errored
   */
  errorTool(toolId: string, error: string): void {
    const tool = this.state.toolCalls.get(toolId);
    if (tool) {
      tool.status = 'error';
      tool.result = error;
      tool.updatedAt = Date.now();
      this.state.lastUpdated = Date.now();
      this.notifyListeners();
    }
  }
}

/**
 * Create a new Claude stream assembler
 */
export function createClaudeStreamAssembler(processId: string): ClaudeStreamAssembler {
  return new ClaudeStreamAssembler(processId);
}
