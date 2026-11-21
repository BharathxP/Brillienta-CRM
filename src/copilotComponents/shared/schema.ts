// Tool parameter schemas for the Copilot Components, as plain JSON Schema (draft-07) in the portable
// Microsoft 365 plugin v2.4 subset: optional top-level properties of type string, integer, number or
// boolean, with descriptions and string enums. No min/max/length keywords: bounds are enforced when the
// values are normalised (crmQueries.ts).
//
// The build loads each *Properties.js file on its own in Node (ES module, no bundler) to generate the tool
// descriptions, so those files may only use `import type` from here; their values must be inline literals.
// toolSchemas.test.ts checks that the inline enum values match the ones the code uses.
export interface IToolParameter {
  type: 'string' | 'integer' | 'number' | 'boolean';
  description: string;
  enum?: string[];
}

export interface IToolSchema {
  $schema: 'http://json-schema.org/draft-07/schema#';
  type: 'object';
  properties: Record<string, IToolParameter>;
  additionalProperties: false;
}
