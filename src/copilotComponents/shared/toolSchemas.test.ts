/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-require-imports */
import { CASE_PRIORITIES } from '../../webparts/brillientaCrm/models/model';
import { HEALTH_FILTERS, LEAD_RATINGS, PIPELINE_PERIODS, PIPELINE_SCOPES } from './toolEnums';
import pipeline from '../crmPipeline/CrmPipelineCopilotComponentProperties';
import account from '../crmAccountSummary/CrmAccountSummaryCopilotComponentProperties';
import renewals from '../crmRenewals/CrmRenewalsCopilotComponentProperties';
import invoices from '../crmOverdueInvoices/CrmOverdueInvoicesCopilotComponentProperties';
import cases from '../crmSlaRiskCases/CrmSlaRiskCasesCopilotComponentProperties';
import lead from '../crmCreateLead/CrmCreateLeadCopilotComponentProperties';

// Jest runs these tests in Node; the project has no @types/node, so type the two Node APIs used here.
interface INodeFs { readFileSync(file: string, encoding: 'utf8'): string; readdirSync(dir: string): string[] }
interface INodePath { resolve(...parts: string[]): string; join(...parts: string[]): string }
declare const __dirname: string;
const fs: INodeFs = require('fs');
const path: INodePath = require('path');

// Tests run from lib-commonjs/copilotComponents/shared.
const ROOT = path.resolve(__dirname, '../../..');
const readJson = (p: string): any => JSON.parse(fs.readFileSync(path.join(ROOT, p), 'utf8').replace(/^\s*\/\/.*$/gm, ''));

const SCHEMAS = { pipeline, account, renewals, invoices, cases, lead };
const PARAM_KEYS = ['type', 'description', 'enum'];
const TYPES = ['string', 'integer', 'number', 'boolean'];

describe('Copilot tool parameter schemas', () => {
  Object.keys(SCHEMAS).forEach(name => {
    const schema = (SCHEMAS as any)[name];
    it(`${name} stays in the portable plugin v2.4 subset`, () => {
      expect(Object.keys(schema).sort()).toEqual(['$schema', 'additionalProperties', 'properties', 'type']);
      expect(schema.type).toBe('object');
      Object.keys(schema.properties).forEach(key => {
        const p = schema.properties[key];
        expect(key).toMatch(/^[A-Za-z0-9_]+$/);
        Object.keys(p).forEach(k => expect(PARAM_KEYS).toContain(k));
        expect(TYPES).toContain(p.type);
        expect(p.description.length).toBeGreaterThan(10);
        expect(p.description.length).toBeLessThan(400);
      });
    });
  });

  it('repeats exactly the enum values the code accepts', () => {
    expect(pipeline.properties.scope.enum).toEqual([...PIPELINE_SCOPES]);
    expect(pipeline.properties.period.enum).toEqual([...PIPELINE_PERIODS]);
    expect(pipeline.properties.health.enum).toEqual([...HEALTH_FILTERS]);
    expect(lead.properties.rating.enum).toEqual([...LEAD_RATINGS]);
    expect(cases.properties.priority.enum).toEqual([...CASE_PRIORITIES]);
  });
});

describe('Copilot agent registration', () => {
  const folder = path.join(ROOT, 'src', 'copilotComponents');
  const manifests = fs.readdirSync(folder)
    .filter(d => d !== 'shared')
    .map(d => {
      const file = fs.readdirSync(path.join(folder, d)).find(f => f.endsWith('CopilotComponent.manifest.json'))!;
      return { dir: d, file: `src/copilotComponents/${d}/${file}`, json: readJson(`src/copilotComponents/${d}/${file}`) };
    });

  it('has six Copilot Components with unique ids, aliases and tool names', () => {
    expect(manifests.length).toBe(6);
    const tools = manifests.map(m => m.json.tools[0].name);
    expect(new Set(manifests.map(m => m.json.id)).size).toBe(6);
    expect(new Set(manifests.map(m => m.json.alias)).size).toBe(6);
    expect(tools.sort()).toEqual(['CreateLead', 'GetAccountSummary', 'GetMyPipeline', 'GetOverdueInvoices', 'GetRenewals', 'GetSlaRiskCases']);
    manifests.forEach(m => {
      expect(m.json.componentType).toBe('CopilotComponent');
      expect(m.json.copilotType).toBe('Ux');
      expect(m.json.tools[0].description.default.length).toBeLessThan(1000);
    });
  });

  it('registers every component with the agent and in the Copilot bundle', () => {
    const agent = readJson('config/copilot-agent.json').agents[0];
    expect([...agent.components].sort()).toEqual(manifests.map(m => m.json.id).sort());
    const bundle = readJson('config/config.json').bundles['brillienta-crm-copilot'].components.map((c: any) => c.manifest).sort();
    expect(bundle).toEqual(manifests.map(m => `./${m.file}`).sort());
  });

  it('names every tool in the agent instructions', () => {
    const instructions = fs.readFileSync(path.join(ROOT, 'copilot', 'instruction.txt'), 'utf8');
    manifests.forEach(m => expect(instructions).toContain(m.json.tools[0].name));
    expect(instructions.length).toBeLessThan(8000);
  });

  it('ships one Teams app with the CRM tabs and the Copilot agent (Sync to Teams publishes only this app)', () => {
    const app = readJson('copilot/manifest.json');
    const webPartId = (fs.readFileSync(path.join(ROOT, 'src/webparts/brillientaCrm/BrillientaCrmWebPart.manifest.json'), 'utf8').match(/"id"\s*:\s*"([0-9a-f-]{36})"/i) || [])[1];
    expect(webPartId).toBeTruthy();
    expect(app.copilotAgents.declarativeAgents.length).toBe(1);
    expect(app.staticTabs.some((t: any) => String(t.contentUrl || '').indexOf(`componentId=${webPartId}`) >= 0)).toBe(true);
    expect(app.configurableTabs.some((t: any) => String(t.configurationUrl || '').indexOf(`componentId=${webPartId}`) >= 0)).toBe(true);
    expect(app.webApplicationInfo.id).toBe('00000003-0000-0ff1-ce00-000000000000');
  });
});
