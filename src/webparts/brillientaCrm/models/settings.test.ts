import { ENTITIES, SYSTEM_LISTS } from './model';
import { mergeSettings } from './settings';

describe('Brillienta brand names', () => {
  it('names every CRM list BrillCRM_<Name> without spaces', () => {
    const titles = [...ENTITIES.map(e => e.listTitle), SYSTEM_LISTS.settings.listTitle, SYSTEM_LISTS.audit.listTitle];
    titles.forEach(t => expect(t).toMatch(/^BrillCRM_[A-Za-z]+$/));
    expect(new Set(titles).size).toBe(titles.length);
  });

  it('names every CRM security group BrillCRM_<Role> by default', () => {
    const g = mergeSettings(undefined).groups;
    expect(g).toEqual({ admins: 'BrillCRM_Administrators', managers: 'BrillCRM_SalesManagers', sales: 'BrillCRM_Sales', support: 'BrillCRM_Support', readers: 'BrillCRM_Readers' });
  });
});
