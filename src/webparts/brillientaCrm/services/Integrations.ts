/* eslint-disable @typescript-eslint/no-explicit-any */
import { MSGraphClientFactory, MSGraphClientV3 } from '@microsoft/sp-http-msgraph';
import { IWebhook } from '../models/settings';

export interface IMeetingResult {
  joinUrl: string;
  webLink: string;
  eventId: string;
}

/**
 * Microsoft Graph integration. Requires the API permissions declared in
 * config/package-solution.json to be approved by a tenant admin
 * (SharePoint admin center > Advanced > API access).
 */
export class GraphService {
  private client?: MSGraphClientV3;

  constructor(private readonly factory?: MSGraphClientFactory) {}

  public get available(): boolean {
    return !!this.factory;
  }

  private async graph(): Promise<MSGraphClientV3> {
    if (!this.factory) {
      throw new Error('Microsoft Graph is not available in this host.');
    }
    if (!this.client) {
      this.client = await this.factory.getClient('3');
    }
    return this.client;
  }

  private explain(e: any): Error {
    const code = e?.statusCode || e?.code;
    if (code === 403 || /Authorization_RequestDenied|AccessDenied|consent/i.test(String(e?.message))) {
      return new Error('Microsoft Graph permission not granted. A SharePoint admin must approve the Brillienta CRM API access requests (SharePoint admin center > Advanced > API access).');
    }
    return e instanceof Error ? e : new Error(String(e?.message || e));
  }

  public async sendMail(to: string[], subject: string, body: string, cc: string[] = [], html: boolean = false): Promise<void> {
    try {
      const client = await this.graph();
      await client.api('/me/sendMail').post({
        message: {
          subject,
          body: { contentType: html ? 'HTML' : 'Text', content: body },
          toRecipients: to.filter(Boolean).map(a => ({ emailAddress: { address: a } })),
          ccRecipients: cc.filter(Boolean).map(a => ({ emailAddress: { address: a } }))
        },
        saveToSentItems: true
      });
    } catch (e) {
      throw this.explain(e);
    }
  }

  /** Creates an Outlook calendar event with a Teams meeting and invites the attendees. */
  public async createTeamsMeeting(subject: string, startIso: string, minutes: number, attendees: string[], body: string): Promise<IMeetingResult> {
    try {
      const client = await this.graph();
      const start = new Date(startIso);
      const end = new Date(start.getTime() + minutes * 60000);
      const ev: any = await client.api('/me/events').post({
        subject,
        body: { contentType: 'Text', content: body },
        start: { dateTime: start.toISOString().replace('Z', ''), timeZone: 'UTC' },
        end: { dateTime: end.toISOString().replace('Z', ''), timeZone: 'UTC' },
        attendees: attendees.filter(Boolean).map(a => ({ emailAddress: { address: a }, type: 'required' })),
        isOnlineMeeting: true,
        onlineMeetingProvider: 'teamsForBusiness',
        allowNewTimeProposals: true
      });
      return { joinUrl: ev?.onlineMeeting?.joinUrl || ev?.onlineMeetingUrl || '', webLink: ev?.webLink || '', eventId: ev?.id };
    } catch (e) {
      throw this.explain(e);
    }
  }

  public async photoUrl(email: string): Promise<string | undefined> {
    try {
      const client = await this.graph();
      const blob: Blob = await client.api(`/users/${encodeURIComponent(email)}/photos/48x48/$value`).responseType('blob' as any).get();
      return URL.createObjectURL(blob);
    } catch {
      return undefined;
    }
  }
}

/**
 * Outgoing webhooks (Power Automate, Logic Apps, Zapier, Make, n8n, custom APIs)
 * and Teams channel notifications (Teams "Workflows" incoming webhook URL).
 * Sent as fire-and-forget text/plain POSTs so they work from the browser without CORS pre-flight.
 */
export class WebhookService {
  constructor(private readonly getHooks: () => IWebhook[], private readonly getTeamsUrl: () => string, private readonly tenant: string) {}

  public emit(event: string, data: Record<string, unknown>, summary: string): void {
    const payload = { event, occurredAt: new Date().toISOString(), tenant: this.tenant, summary, data };
    this.getHooks()
      .filter(h => h.active && h.url && (h.events.indexOf(event) >= 0 || h.events.indexOf('*') >= 0))
      .forEach(h => { this.post(h.url, payload).catch(() => undefined); });

    const teams = this.getTeamsUrl();
    if (teams && /\.(won|requested|created|breached|accepted|converted)$/i.test(event) && event !== 'account.created' && event !== 'contact.created') {
      this.post(teams, {
        type: 'message',
        attachments: [{
          contentType: 'application/vnd.microsoft.card.adaptive',
          content: {
            $schema: 'http://adaptivecards.io/schemas/adaptive-card.json', type: 'AdaptiveCard', version: '1.4',
            body: [
              { type: 'TextBlock', text: 'Brillienta CRM', weight: 'Bolder', size: 'Small', isSubtle: true },
              { type: 'TextBlock', text: summary, wrap: true, weight: 'Bolder' },
              { type: 'TextBlock', text: event, isSubtle: true, size: 'Small' }
            ]
          }
        }]
      }).catch(() => undefined);
    }
  }

  public async test(url: string): Promise<void> {
    await this.post(url, { event: 'test', occurredAt: new Date().toISOString(), tenant: this.tenant, summary: 'Test event from Brillienta CRM', data: {} });
  }

  private async post(url: string, body: unknown): Promise<void> {
    if (!/^https:\/\//i.test(url)) {
      return;
    }
    await fetch(url, { method: 'POST', mode: 'no-cors', headers: { 'Content-Type': 'text/plain' }, body: JSON.stringify(body) });
  }
}
