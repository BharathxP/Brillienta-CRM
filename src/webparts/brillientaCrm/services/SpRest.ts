import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

/* eslint-disable @typescript-eslint/no-explicit-any */

const JSON_HEADERS = {
  Accept: 'application/json;odata=nometadata',
  'Content-Type': 'application/json;odata=nometadata',
  'odata-version': ''
};

export class SpRest {
  constructor(private readonly client: SPHttpClient, public readonly webUrl: string) {}

  public listUrl(title: string): string {
    return `${this.webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(title.replace(/'/g, "''"))}')`;
  }

  public async get<T = any>(url: string): Promise<T> {
    const res = await this.client.get(this.abs(url), SPHttpClient.configurations.v1, { headers: JSON_HEADERS });
    return this.handle<T>(res);
  }

  /** Follows odata.nextLink so lists above 5,000 items are read completely. */
  public async getAll<T = any>(url: string, maxItems: number = 50000): Promise<T[]> {
    const results: T[] = [];
    let next: string | undefined = url;
    while (next && results.length < maxItems) {
      const page: any = await this.get(next);
      results.push(...(page.value || []));
      next = page['odata.nextLink'] || page['@odata.nextLink'];
    }
    return results;
  }

  public async post<T = any>(url: string, body?: any, extraHeaders: Record<string, string> = {}): Promise<T> {
    const options: ISPHttpClientOptions = {
      headers: { ...JSON_HEADERS, ...extraHeaders },
      body: body === undefined ? undefined : JSON.stringify(body)
    };
    const res = await this.client.post(this.abs(url), SPHttpClient.configurations.v1, options);
    return this.handle<T>(res);
  }

  public async postVerbose<T = any>(url: string, body: any, extraHeaders: Record<string, string> = {}): Promise<T> {
    const options: ISPHttpClientOptions = {
      headers: {
        Accept: 'application/json;odata=verbose',
        'Content-Type': 'application/json;odata=verbose',
        'odata-version': '',
        ...extraHeaders
      },
      body: JSON.stringify(body)
    };
    const res = await this.client.post(this.abs(url), SPHttpClient.configurations.v1, options);
    return this.handle<T>(res);
  }

  public merge(url: string, body: any): Promise<void> {
    return this.post(url, body, { 'X-HTTP-Method': 'MERGE', 'IF-MATCH': '*' });
  }

  public delete(url: string): Promise<void> {
    return this.post(url, undefined, { 'X-HTTP-Method': 'DELETE', 'IF-MATCH': '*' });
  }

  private abs(url: string): string {
    return url.indexOf('http') === 0 ? url : `${this.webUrl}${url}`;
  }

  private async handle<T>(res: SPHttpClientResponse): Promise<T> {
    if (!res.ok) {
      let message = `${res.status} ${res.statusText}`;
      try {
        const err: any = await res.json();
        message = err?.error?.message?.value || err?.['odata.error']?.message?.value || err?.error?.message || message;
      } catch {
        /* ignore body parse errors */
      }
      const e: any = new Error(message);
      e.status = res.status;
      throw e;
    }
    if (res.status === 204) {
      return undefined as unknown as T;
    }
    const text = await res.text();
    return (text ? JSON.parse(text) : undefined) as T;
  }
}
