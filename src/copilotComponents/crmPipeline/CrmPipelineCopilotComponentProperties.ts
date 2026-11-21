import type { IToolSchema } from '../shared/schema';

export interface ICrmPipelineCopilotComponentProperties {
  scope?: string;
  period?: string;
  health?: string;
  minAmount?: number;
}

const schema: IToolSchema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  additionalProperties: false,
  properties: {
    scope: { type: 'string', enum: ['mine', 'team', 'all'], description: 'Whose deals: mine (default), team (the user\'s sales team) or all deals the user can see.' },
    period: { type: 'string', enum: ['thisQuarter', 'nextQuarter', 'thisMonth', 'overdue', 'all'], description: 'Expected close period: thisQuarter (default, fiscal quarter), nextQuarter, thisMonth, overdue (close date passed) or all.' },
    health: { type: 'string', enum: ['any', 'atRisk', 'watch', 'healthy'], description: 'Only deals with this AI health: atRisk, watch, healthy, or any (default).' },
    minAmount: { type: 'number', description: 'Only deals worth at least this amount in the company currency.' }
  }
};

export default schema;
