/**
 * A stand-in for M365's SignalR endpoint, enough to drive CopilotSession end to
 * end: the 0x1E framing, the handshake, ping frames, and the update / completion
 * frames a turn is made of. Every client frame is captured so a test can assert
 * on things the real server punishes silently — above all the Metrics frame.
 *
 * Trimmed from the reference stub in opencode-m365-copilot (test/stub-copilot.ts).
 */
import { createServer } from 'node:http';
import type { AddressInfo } from 'node:net';
import { WebSocketServer, type WebSocket } from 'ws';

const RS = '\x1e';

export interface StubConn {
  url: string;
  frames: Record<string, unknown>[];
  socket: WebSocket;
}

export interface StubOptions {
  /** Frames to send after the chat frame arrives. Default: two deltas + completion. */
  respond?: (chat: Record<string, unknown>) => unknown[];
  /** Send a keepalive ping before responding. */
  pingFirst?: boolean;
}

export interface StubServer {
  origin: string;
  connections: StubConn[];
  last(): StubConn;
  close(): Promise<void>;
}

export async function startStub(options: StubOptions = {}): Promise<StubServer> {
  const connections: StubConn[] = [];
  const http = createServer();
  const wss = new WebSocketServer({ server: http });
  await new Promise<void>((resolve) => http.listen(0, '127.0.0.1', resolve));
  const { port } = http.address() as AddressInfo;

  wss.on('connection', (socket, request) => {
    const conn: StubConn = { url: request.url ?? '', frames: [], socket };
    connections.push(conn);
    socket.on('message', (raw: unknown) => {
      for (const chunk of String(raw).split(RS)) {
        if (chunk.trim() === '') continue;
        let frame: Record<string, unknown>;
        try {
          frame = JSON.parse(chunk);
        } catch {
          continue;
        }
        conn.frames.push(frame);
        if (frame.protocol === 'json') {
          socket.send(`{}${RS}`);
          continue;
        }
        if (frame.type === 4 && frame.target === 'chat') {
          if (options.pingFirst) socket.send(`${JSON.stringify({ type: 6 })}${RS}`);
          const out = options.respond?.(frame) ?? [delta('Hello'), delta(' world'), completion()];
          for (const f of out) socket.send(`${JSON.stringify(f)}${RS}`);
        }
      }
    });
  });

  return {
    origin: `ws://127.0.0.1:${port}`,
    connections,
    last: () => connections[connections.length - 1]!,
    close: () =>
      new Promise<void>((resolve) => {
        for (const c of connections) c.socket.terminate();
        wss.close(() => {
          http.closeAllConnections();
          http.close(() => resolve());
        });
      }),
  };
}

export const delta = (text: string) => ({
  type: 1,
  target: 'update',
  arguments: [{ writeAtCursor: text }],
});

export const botMessage = (text: string, extra: Record<string, unknown> = {}) => ({
  type: 1,
  target: 'update',
  arguments: [{ messages: [{ author: 'bot', text, ...extra }] }],
});

export const disengaged = () => ({
  type: 1,
  target: 'update',
  arguments: [{ messages: [{ author: 'bot', messageType: 'Disengaged' }] }],
});

export const throttling = (current: number, max = 600) => ({
  type: 1,
  target: 'update',
  arguments: [{ throttling: { numUserMessagesInConversation: current, maxNumUserMessagesInConversation: max } }],
});

export const completion = (error?: string) => (error ? { type: 3, error } : { type: 3 });

/** A JWT whose payload carries oid/tid, for CopilotSession.buildUrl/decodeJwt. */
export function fakeToken(oid = 'oid-1', tid = 'tid-1'): string {
  const payload = Buffer.from(JSON.stringify({ oid, tid })).toString('base64url');
  return `eyJhbGciOiJub25lIn0.${payload}.sig`;
}
