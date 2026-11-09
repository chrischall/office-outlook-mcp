/**
 * The M365 Copilot transport: SignalR over WebSocket.
 *
 * Adapted from the sibling project `opencode-m365-copilot` (MIT, same owner):
 * https://github.com/chrischall/opencode-copilot-plugin — src/session.ts, which
 * reverse-engineered this protocol. Trimmed to a single mail-grounded turn (no
 * image upload, no declarative agents) and pointed at the Outlook side-pane
 * scenario rather than the full-screen one, because the side-pane scenario is
 * the surface that grounds on the signed-in mailbox.
 *
 * Load-bearing and non-obvious, all of which fail SILENTLY when wrong:
 *  - The access token goes in the WebSocket URL query string, not a header.
 *  - Node's built-in WebSocket is refused; `ws` plus a browser-shaped
 *    Origin/User-Agent is what the server accepts.
 *  - Frames are JSON terminated by 0x1E (record separator); one message can
 *    carry several.
 *  - A `Metrics` frame must ride in the SAME send as the chat frame. Omit it and
 *    the turn never produces output — no error, just silence.
 *  - Only bot messages with NO `messageType` carry the answer; everything with
 *    one is control traffic.
 *  - `messageType: "Disengaged"` is a refusal with empty content. It looks
 *    exactly like rate limiting and must not be retried into.
 */
import { WebSocket } from 'ws';
import { randomUUID } from 'node:crypto';

/** Record separator: every SignalR frame ends with this byte. */
const RS = '\x1e';

const DEFAULT_ENDPOINT = 'wss://substrate.office.com';

/** The upgrade is rejected without browser-shaped headers. */
const BROWSER_HEADERS = {
  Origin: 'https://m365copilotapp.svc.cloud.microsoft',
  'User-Agent':
    'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36',
};

/**
 * Feature flags and option sets copied from a captured Outlook side-pane
 * session. Cargo-culted on purpose: these are the values the real client sends;
 * which subset actually matters has never been established, and removing them is
 * untested. `WorkSearch` in the agent capabilities is what grounds the turn on
 * the user's mailbox.
 */
const VARIANTS = [
  'feature.OutlookSidepaneWorkCopilotToHelix',
  'feature.EnableStarterLicenseCheckBypass',
  'feature.IsStreamingModeInChatRequestEnabled',
  'feature.isExternalEmailEnabled',
  'IncludeSourceAttributionsConcise',
].join(',');

const OPTIONS_SETS = [
  'enterprise_flux_work',
  'enterprise_flux_web',
  'enterprise_pagination_support',
  'enable_search_result_progress_messages',
  'rich_responses',
  'enable_inferred_memory_read',
  'flux_v3_references',
];

/** The message types we tell the server we can render. */
const ALLOWED_MESSAGE_TYPES = [
  'Chat',
  'Suggestion',
  'Progress',
  'EndOfRequest',
  'Disengaged',
  'InternalSearchQuery',
  'InternalSearchResult',
  'RenderCardRequest',
  'ReferencesListComplete',
  'GeneratedCode',
];

/** Read the object id and tenant id out of a Sydney token — they address the hub. */
export function decodeJwt(token: string): { oid: string; tid: string } {
  const parts = token.split('.');
  if (parts.length < 2) throw new Error('Malformed Copilot token: expected a JWT');
  let claims: { oid?: string; tid?: string };
  try {
    claims = JSON.parse(Buffer.from(parts[1]!, 'base64url').toString('utf8'));
  } catch {
    throw new Error('Malformed Copilot token: payload is not valid base64url JSON');
  }
  if (!claims?.oid || !claims?.tid) throw new Error('Copilot token is missing oid/tid claims');
  return { oid: claims.oid, tid: claims.tid };
}

/**
 * Strip Copilot's inline citation control tokens from an answer.
 *
 * The real web UI renders these as citation chips; over the wire they arrive as
 * literal noise: private-use-area delimiters (U+E000–U+F8FF) wrapping
 * `cite…turn1search23…` runs, plus `【2-9c5ade】`-style attribution markers.
 * None of it is content — removed so the returned text reads as prose. Measured
 * against live answers 2026-11-09.
 */
export function cleanAnswer(text: string): string {
  return text
    .replace(/[\u{E000}-\u{F8FF}\u{F0000}-\u{FFFFD}\u{100000}-\u{10FFFD}]/gu, '')
    .replace(/cite(?:turn\d+[a-z]+\d+)+/gi, '')
    .replace(/【[^】]{1,60}】/g, '')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/[ \t]{2,}/g, ' ')
    .trim();
}

/** M365 refused to engage with the prompt. Not a rate limit — do not retry into it. */
export class DisengagedError extends Error {
  readonly disengaged = true;
  constructor(message = 'M365 Copilot declined to answer this prompt') {
    super(message);
    this.name = 'DisengagedError';
  }
}

export interface ThrottleState {
  current: number;
  max: number;
}

export interface TurnResult {
  text: string;
  /** Per-conversation quota, when the server reported it. */
  throttle?: ThrottleState;
  /** Which backend answered — `DeepLeo` is the reasoning pipeline. */
  contentOrigin?: string;
}

export interface CopilotSessionOptions {
  /** Supplies a Sydney token; called once per turn so refresh is transparent. */
  getToken: () => Promise<string>;
  /** Override the wss origin. Tests point this at a stub. */
  endpoint?: string;
}

export interface TurnOptions {
  /** Selects the model. There is no model parameter in this protocol. */
  tone?: string;
  signal?: AbortSignal;
  onDelta?: (chunk: string) => void;
}

interface Frame {
  type?: number;
  target?: string;
  arguments?: unknown[];
  item?: unknown;
  error?: unknown;
  protocol?: string;
}

/**
 * One M365 conversation.
 *
 * Stateful across turns: the conversation and session ids persist so the server
 * keeps its own context, and follow-up turns only carry the new message. The
 * 600-message cap is per conversation, which is why we reuse rather than opening
 * a fresh one each turn.
 */
export class CopilotSession {
  conversationId = randomUUID();
  sessionId = randomUUID();
  turnIndex = 0;

  private readonly options: CopilotSessionOptions;

  constructor(options: CopilotSessionOptions) {
    this.options = options;
  }

  /** Abandon the server-side context and start a fresh conversation. */
  reset(): void {
    this.conversationId = randomUUID();
    this.sessionId = randomUUID();
    this.turnIndex = 0;
  }

  async run(text: string, turn: TurnOptions = {}): Promise<TurnResult> {
    if (turn.signal?.aborted) throw abortError();
    const token = await this.options.getToken();
    const { oid, tid } = decodeJwt(token);
    const url = this.buildUrl(token, oid, tid);
    const socket = new WebSocket(url, { headers: BROWSER_HEADERS });
    try {
      const result = await this.pump(socket, text, turn);
      this.turnIndex += 1;
      return result;
    } finally {
      closeSocket(socket);
    }
  }

  private buildUrl(token: string, oid: string, tid: string): string {
    const origin = this.options.endpoint ?? DEFAULT_ENDPOINT;
    const params = new URLSearchParams({
      access_token: token,
      ConversationId: this.conversationId,
      chatsessionid: this.sessionId,
      'X-SessionId': this.sessionId,
      clientrequestid: randomUUID(),
      source: '"owahub"',
      product: 'OwaHub',
      agentHost: 'Bizchat.Sidepane',
      scenario: 'OutlookSidepaneWorkCopilot',
      agent: 'work',
      variants: VARIANTS,
    });
    return `${origin}/m365Copilot/Chathub/${oid}@${tid}?${params.toString()}`;
  }

  /** Drive one turn to completion. */
  private pump(socket: WebSocket, text: string, turn: TurnOptions): Promise<TurnResult> {
    return new Promise<TurnResult>((resolve, reject) => {
      let deltaText = '';
      let snapshotText = '';
      let handshakeDone = false;
      let settled = false;
      const result: TurnResult = { text: '' };

      const finish = (error?: Error) => {
        if (settled) return;
        settled = true;
        turn.signal?.removeEventListener('abort', onAbort);
        if (error) {
          reject(error);
          return;
        }
        // The server sends the answer twice — incrementally and as a snapshot —
        // and either can be the more complete one depending on where it ended.
        result.text = cleanAnswer(deltaText.length >= snapshotText.length ? deltaText : snapshotText);
        resolve(result);
      };

      const onAbort = () => {
        try {
          socket.send(
            `${JSON.stringify({ arguments: [{}], invocationId: '1', target: 'stop', type: 1 })}${RS}`,
          );
        } catch {
          /* the socket may already be gone */
        }
        finish(abortError());
      };
      turn.signal?.addEventListener('abort', onAbort, { once: true });

      socket.on('error', (error) => finish(error instanceof Error ? error : new Error(String(error))));
      // Only a completion frame (type 3, or a clean type 7) ends a turn. A close
      // before one is a dropped transport: resolving with whatever arrived would
      // hand back a truncated — or empty — answer that looks like a dead agent.
      socket.on('close', (code, reason) => {
        const why = reason.toString();
        finish(new Error(`M365 connection closed before the turn completed (code ${code}${why ? `: ${why}` : ''})`));
      });

      socket.on('open', () => {
        socket.send(`${JSON.stringify({ protocol: 'json', version: 1 })}${RS}`);
      });

      socket.on('message', (raw: unknown) => {
        for (const frame of splitFrames(String(raw))) {
          // The handshake ack is `{}`; the chat frame goes out immediately after.
          if (!handshakeDone) {
            handshakeDone = true;
            this.sendChat(socket, text, turn);
            if (Object.keys(frame).length === 0) continue;
          }

          switch (frame.type) {
            case 1:
              if (frame.target === 'update') {
                for (const argument of frame.arguments ?? []) {
                  const applied = applyUpdate(argument, result, turn);
                  if (applied.disengaged) {
                    finish(new DisengagedError());
                    return;
                  }
                  deltaText += applied.delta;
                  if (applied.snapshot.length > snapshotText.length) snapshotText = applied.snapshot;
                }
              }
              break;
            case 2:
              applyStreamItem(frame.item, result);
              break;
            case 3:
              if (frame.error) {
                finish(new Error(`M365 rejected the turn: ${JSON.stringify(frame.error)}`));
                return;
              }
              finish();
              return;
            case 6:
              // Keepalive. Silence here gets the connection dropped mid-answer.
              socket.send(`${JSON.stringify({ type: 6 })}${RS}`);
              break;
            case 7:
              if (frame.error) {
                finish(new Error(`M365 closed the connection: ${JSON.stringify(frame.error)}`));
                return;
              }
              finish();
              return;
            default:
              break;
          }
        }
      });
    });
  }

  /** The chat invocation and its mandatory Metrics companion, in one send. */
  private sendChat(socket: WebSocket, text: string, turn: TurnOptions): void {
    const now = new Date().toISOString();
    const argument = {
      source: 'owahub',
      optionsSets: OPTIONS_SETS,
      streamingMode: 'ConciseWithPadding',
      allowedMessageTypes: ALLOWED_MESSAGE_TYPES,
      isStartOfSession: this.turnIndex === 0,
      conversationId: this.conversationId,
      message: {
        author: 'user',
        inputMethod: 'Keyboard',
        text,
        entityAnnotationTypes: ['People', 'File', 'Event', 'Email', 'TeamsMessage'],
        locale: 'en-us',
        messageType: 'Chat',
        experienceType: 'Default',
        timestamp: now,
        messageId: randomUUID(),
      },
      // The side-pane agent. `WorkSearch` is the capability that grounds the turn
      // on the signed-in mailbox; `WebSearch` lets it fall back to Bing.
      gpts: [
        {
          id: 'bizchat-as-gpt-scenario',
          source: 'BuiltInAgents',
          clientOverrides: {
            capabilities: [{ name: 'WebSearch' }, { name: 'WorkSearch' }],
          },
        },
      ],
      plugins: [{ Id: 'BingWebSearch', Source: 'BuiltIn' }],
      tone: turn.tone ?? 'Magic',
      clientInfo: {
        clientPlatform: 'OwaHub-web',
        clientAppName: 'OwaHub',
        clientAppType: 'Web',
        productEntryPoint: 'Email',
      },
      traceId: randomUUID(),
    };

    const chat = { arguments: [argument], invocationId: '0', target: 'chat', type: 4 };
    const metrics = {
      arguments: [
        {
          Timestamps: {
            ConnectionStart: now,
            UserInputStart: now,
            ConnectionEstablished: now,
            UserInputSubmit: now,
          },
        },
      ],
      target: 'Metrics',
      type: 1,
    };
    // Both frames, one send. This is the part that silently breaks everything.
    socket.send(`${JSON.stringify(chat)}${RS}${JSON.stringify(metrics)}${RS}`);
  }
}

/** Fold one `update` argument into the accumulating result. */
function applyUpdate(
  argument: unknown,
  result: TurnResult,
  turn: TurnOptions,
): { delta: string; snapshot: string; disengaged: boolean } {
  const arg = argument as {
    writeAtCursor?: unknown;
    throttling?: { numUserMessagesInConversation?: number; maxNumUserMessagesInConversation?: number };
    messages?: { author?: string; messageType?: string; text?: unknown; contentOrigin?: unknown }[];
  };
  let delta = '';
  let snapshot = '';

  if (typeof arg.writeAtCursor === 'string') {
    delta = arg.writeAtCursor;
    turn.onDelta?.(delta);
  }
  if (arg.throttling) {
    result.throttle = {
      current: arg.throttling.numUserMessagesInConversation ?? 0,
      max: arg.throttling.maxNumUserMessagesInConversation ?? 0,
    };
  }
  for (const message of arg.messages ?? []) {
    if (message?.author !== 'bot') continue;
    if (message.messageType === 'Disengaged') return { delta, snapshot, disengaged: true };
    if (message.messageType) continue; // a messageType means control traffic, never the answer
    if (typeof message.text === 'string' && message.text.length > snapshot.length) snapshot = message.text;
    if (typeof message.contentOrigin === 'string') result.contentOrigin = message.contentOrigin;
  }
  return { delta, snapshot, disengaged: false };
}

function applyStreamItem(item: unknown, result: TurnResult): void {
  const it = item as {
    throttling?: { numUserMessagesInConversation?: number; maxNumUserMessagesInConversation?: number };
    messages?: { author?: string; messageType?: string; contentOrigin?: unknown }[];
  } | null;
  if (!it) return;
  if (it.throttling) {
    result.throttle = {
      current: it.throttling.numUserMessagesInConversation ?? 0,
      max: it.throttling.maxNumUserMessagesInConversation ?? 0,
    };
  }
  for (const message of it.messages ?? []) {
    if (message?.author === 'bot' && !message.messageType && typeof message.contentOrigin === 'string') {
      result.contentOrigin = message.contentOrigin;
    }
  }
}

/** One WebSocket message can carry several 0x1E-terminated frames. */
function splitFrames(raw: string): Frame[] {
  const frames: Frame[] = [];
  for (const chunk of raw.split(RS)) {
    if (chunk.trim() === '') continue;
    try {
      frames.push(JSON.parse(chunk) as Frame);
    } catch {
      /* an unparseable frame is skipped rather than killing the turn */
    }
  }
  return frames;
}

/**
 * Shut a turn's socket down without losing anything still buffered.
 *
 * `terminate()` drops queued frames — which is how a pong or the stop frame can
 * be sent and still never arrive. `close()` flushes first; the timer is the
 * backstop for a server that never completes the handshake.
 */
function closeSocket(socket: WebSocket): void {
  if (socket.readyState === WebSocket.CLOSED || socket.readyState === WebSocket.CLOSING) return;
  try {
    socket.close();
  } catch {
    socket.terminate();
    return;
  }
  const timer = setTimeout(() => socket.terminate(), 1000);
  timer.unref?.();
  socket.once('close', () => clearTimeout(timer));
}

function abortError(): Error {
  const error = new Error('The M365 Copilot turn was aborted');
  error.name = 'AbortError';
  return error;
}
