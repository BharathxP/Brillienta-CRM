// Allowed values for tool parameters, used when prompt values are normalised (crmQueries.ts).
// The *Properties.ts schemas repeat them as literals (see schema.ts); toolSchemas.test.ts keeps both in sync.
export const PIPELINE_SCOPES = ['mine', 'team', 'all'] as const;
export const PIPELINE_PERIODS = ['thisQuarter', 'nextQuarter', 'thisMonth', 'overdue', 'all'] as const;
export const HEALTH_FILTERS = ['any', 'atRisk', 'watch', 'healthy'] as const;
export const LEAD_RATINGS = ['Hot', 'Warm', 'Cold'] as const;
