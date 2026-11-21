import type { IToolSchema } from '../shared/schema';

export interface ICrmSlaRiskCasesCopilotComponentProperties {
  windowHours?: number;
  priority?: string;
  onlyMine?: boolean;
}

const schema: IToolSchema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  additionalProperties: false,
  properties: {
    windowHours: { type: 'integer', description: 'Also include cases whose SLA deadline is within this many hours. Default 24; "today" is about 8, "this week" about 120.' },
    priority: { type: 'string', enum: ['Low', 'Normal', 'High', 'Critical'], description: 'Only cases with this priority.' },
    onlyMine: { type: 'boolean', description: 'True when the user asks for their own cases.' }
  }
};

export default schema;
