import type { IToolSchema } from '../shared/schema';

export interface ICrmOverdueInvoicesCopilotComponentProperties {
  minDaysOverdue?: number;
  accountName?: string;
  onlyMine?: boolean;
}

const schema: IToolSchema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  additionalProperties: false,
  properties: {
    minDaysOverdue: { type: 'integer', description: 'Only invoices at least this many days past their due date, for example 30. Default 1.' },
    accountName: { type: 'string', description: 'Only invoices of customers whose name contains this text.' },
    onlyMine: { type: 'boolean', description: 'True when the user asks for invoices they own ("my overdue invoices").' }
  }
};

export default schema;
