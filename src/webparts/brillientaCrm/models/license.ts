import { LICENSE_PUBLIC_KEY } from './licensePublicKey';

export type Plan = 'trial' | 'standard' | 'enterprise';
export type LicenseStatus = 'trial' | 'active' | 'expired' | 'trialExpired' | 'invalid' | 'wrongTenant';

export interface ILicensePayload {
  id: string;
  customer: string;
  tenantId: string;
  plan: 'standard' | 'enterprise';
  seats: number;
  issued: string;
  expires: string;
}

export interface ILicenseInfo {
  status: LicenseStatus;
  plan: Plan;
  /** Features of this plan are available. */
  features: 'standard' | 'enterprise';
  /** False when the trial/licence has lapsed: the CRM becomes read-only. */
  canWrite: boolean;
  daysLeft?: number;
  payload?: ILicensePayload;
  message: string;
}

export const TRIAL_DAYS = 30;

function b64uToBytes(s: string): Uint8Array {
  const b64 = s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4);
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) {
    out[i] = bin.charCodeAt(i);
  }
  return out;
}

function utf8Decode(bytes: Uint8Array): string {
  let s = '';
  for (let i = 0; i < bytes.length; i++) {
    s += String.fromCharCode(bytes[i]);
  }
  return decodeURIComponent(escape(s));
}

function asciiBytes(s: string): Uint8Array {
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) {
    out[i] = s.charCodeAt(i);
  }
  return out;
}

const DAY = 86400000;

export async function evaluateLicense(key: string, tenantId: string, installDate: string): Promise<ILicenseInfo> {
  const trial = (): ILicenseInfo => {
    const started = new Date(installDate).getTime() || Date.now();
    const left = Math.ceil((started + TRIAL_DAYS * DAY - Date.now()) / DAY);
    return left > 0
      ? { status: 'trial', plan: 'trial', features: 'enterprise', canWrite: true, daysLeft: left, message: `Trial: ${left} day(s) left` }
      : { status: 'trialExpired', plan: 'trial', features: 'enterprise', canWrite: false, daysLeft: 0, message: 'Trial expired. The CRM is read-only until a licence key is entered.' };
  };

  const trimmed = (key || '').trim();
  if (!trimmed) {
    return trial();
  }
  const parts = trimmed.split('.');
  if (parts.length !== 2) {
    return { ...trial(), status: 'invalid', message: 'Licence key format is invalid.' };
  }
  try {
    const pub = await crypto.subtle.importKey('jwk', LICENSE_PUBLIC_KEY, { name: 'ECDSA', namedCurve: 'P-256' }, false, ['verify']);
    const ok = await crypto.subtle.verify({ name: 'ECDSA', hash: { name: 'SHA-256' } }, pub, b64uToBytes(parts[1]) as BufferSource, asciiBytes(parts[0]) as BufferSource);
    if (!ok) {
      return { ...trial(), status: 'invalid', message: 'Licence key signature is invalid.' };
    }
    const payload: ILicensePayload = JSON.parse(utf8Decode(b64uToBytes(parts[0])));
    if (payload.tenantId && tenantId && payload.tenantId.toLowerCase() !== tenantId.toLowerCase()) {
      return { ...trial(), status: 'wrongTenant', payload, message: 'This licence was issued for a different Microsoft 365 tenant.' };
    }
    const exp = new Date(`${payload.expires}T23:59:59Z`).getTime();
    const left = Math.ceil((exp - Date.now()) / DAY);
    if (left <= 0) {
      return { status: 'expired', plan: payload.plan, features: payload.plan, canWrite: false, daysLeft: 0, payload, message: `Licence expired on ${payload.expires}. The CRM is read-only.` };
    }
    return {
      status: 'active', plan: payload.plan, features: payload.plan, canWrite: true, daysLeft: left, payload,
      message: `${payload.plan === 'enterprise' ? 'Enterprise' : 'Standard'} licence for ${payload.customer}, valid until ${payload.expires}`
    };
  } catch (e) {
    return { ...trial(), status: 'invalid', message: `Licence key could not be verified (${(e as Error).message}).` };
  }
}
