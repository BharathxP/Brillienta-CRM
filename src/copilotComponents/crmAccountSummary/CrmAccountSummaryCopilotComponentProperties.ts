import type { IToolSchema } from '../shared/schema';

export interface ICrmAccountSummaryCopilotComponentProperties {
  accountName?: string;
}

const schema: IToolSchema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  additionalProperties: false,
  properties: {
    accountName: { type: 'string', description: 'Customer company name as the user wrote it, for example "Contoso". Partial names are matched; the user picks when several accounts match.' }
  }
};

export default schema;
