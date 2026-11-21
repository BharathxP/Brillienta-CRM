import type { IToolSchema } from '../shared/schema';

export interface ICrmRenewalsCopilotComponentProperties {
  days?: number;
  onlyMine?: boolean;
  includeAutoRenew?: boolean;
}

const schema: IToolSchema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  additionalProperties: false,
  properties: {
    days: { type: 'integer', description: 'Renewal window in days from today. Default 60; "this month" is about 30, "this quarter" about 90.' },
    onlyMine: { type: 'boolean', description: 'True when the user asks for their own subscriptions ("my renewals").' },
    includeAutoRenew: { type: 'boolean', description: 'False to show only subscriptions that do not renew automatically. Default true.' }
  }
};

export default schema;
