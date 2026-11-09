/**
 * The Copilot client: a Sydney token source plus one persistent conversation.
 *
 * Built once in `index.ts` and handed to `registerCopilotTools`, the same way
 * the OutlookClient is — `runMcp` reconstructs registrars per connection, so any
 * state (the MSAL cache handle, the conversation id) must live here, not in the
 * registrar. Nothing here touches the network or MSAL until the first ask, so a
 * server never asked to reach Copilot pays nothing for it.
 */
import { createCopilotTokenClient, type CopilotTokenClient } from './auth.js';
import { CopilotSession, type TurnOptions, type TurnResult } from './session.js';

export interface CopilotClientOptions {
  env?: NodeJS.ProcessEnv;
  /**
   * Override the Copilot WebSocket origin. Tests point this at a stub; the
   * shipped server never sets it, so it is a constructor option rather than an
   * env var the server reads (which the surface lint would require declaring).
   */
  endpoint?: string;
  /** Injected in tests; defaults to the real MSAL-backed client. */
  tokenClient?: CopilotTokenClient;
  /** Injected in tests; defaults to a real session over the token client. */
  session?: CopilotSession;
}

export class CopilotClient {
  readonly #env: NodeJS.ProcessEnv;
  readonly #endpoint: string | undefined;
  #tokenClient: CopilotTokenClient | undefined;
  #session: CopilotSession | undefined;

  constructor(opts: CopilotClientOptions = {}) {
    this.#env = opts.env ?? process.env;
    this.#endpoint = opts.endpoint;
    this.#tokenClient = opts.tokenClient;
    this.#session = opts.session;
  }

  #tokens(): CopilotTokenClient {
    this.#tokenClient ??= createCopilotTokenClient({ env: this.#env });
    return this.#tokenClient;
  }

  #sessionOrNew(): CopilotSession {
    this.#session ??= new CopilotSession({
      getToken: () => this.#tokens().getToken(),
      endpoint: this.#endpoint,
    });
    return this.#session;
  }

  /** The current conversation id, for reporting which thread an answer is in. */
  get conversationId(): string {
    return this.#sessionOrNew().conversationId;
  }

  /** Ask Copilot, continuing the current conversation. */
  ask(prompt: string, opts: TurnOptions = {}): Promise<TurnResult> {
    return this.#sessionOrNew().run(prompt, opts);
  }

  /** Start a fresh conversation; the next ask has no prior context. */
  reset(): void {
    this.#sessionOrNew().reset();
  }
}
