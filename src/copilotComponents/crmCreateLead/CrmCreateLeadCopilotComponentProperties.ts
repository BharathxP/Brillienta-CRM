import type { IToolSchema } from '../shared/schema';

export interface ICrmCreateLeadCopilotComponentProperties {
  firstName?: string;
  lastName?: string;
  email?: string;
  phone?: string;
  company?: string;
  jobTitle?: string;
  country?: string;
  leadSource?: string;
  rating?: string;
  description?: string;
}

const schema: IToolSchema = {
  $schema: 'http://json-schema.org/draft-07/schema#',
  type: 'object',
  additionalProperties: false,
  properties: {
    firstName: { type: 'string', description: 'Lead first name.' },
    lastName: { type: 'string', description: 'Lead last name.' },
    email: { type: 'string', description: 'Business email address.' },
    phone: { type: 'string', description: 'Phone number, with country code when given.' },
    company: { type: 'string', description: 'Company the lead works for.' },
    jobTitle: { type: 'string', description: 'Job title, for example CTO or Procurement manager.' },
    country: { type: 'string', description: 'Country of the lead or company.' },
    leadSource: { type: 'string', description: 'Where the lead came from, for example Event, Referral, Website, Webinar, Partner or Cold outreach. Matched to the CRM\'s lead sources.' },
    rating: { type: 'string', enum: ['Hot', 'Warm', 'Cold'], description: 'Hot, Warm (default) or Cold, only when the user says how interested the lead is.' },
    description: { type: 'string', description: 'Short notes about the lead or the conversation, in the user\'s words.' }
  }
};

export default schema;
