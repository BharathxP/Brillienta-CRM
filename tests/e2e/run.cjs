// End-to-end journeys for Brillienta CRM (npm run e2e). Each journey drives the real app in the local preview as a
// different person and checks what was saved in the (in-memory) SharePoint lists. Prints PASS / FAIL per check,
// saves screenshots to tests/e2e/output and exits 1 when anything fails or the browser reported an error.
const { check, start, summary } = require('./helpers.cjs');

(async () => {
  const { browser, h } = await start();
  const { p } = h;
  const step = async (name, fn) => {
    try { await fn(); } catch (e) { check(`${name}: finished without an error`, false, e.message.split('\n')[0]); await h.shot('error-' + name.replace(/\W+/g, '-')); }
  };

  // ---------------------------------------------------------------- 1. setup
  await step('Setup', async () => {
    await h.go('asha', 'dashboard', '&reset=1');
    check('Fresh site shows the setup screen with BrillCRM_ lists', await h.visible('text=BrillCRM_Opportunities'));
    await h.click('button:has-text("Set up CRM")');
    await p.waitForSelector('text=AI BRIEFING', { timeout: 30000 });
    const s = await h.state();
    const lists = Object.keys(s.lists);
    check('Setup creates all 15 BrillCRM_ lists', lists.length === 15 && lists.every((t) => /^BrillCRM_/.test(t)), lists.join(', '));
    check('Settings and audit lists are kept out of search', s.lists.BrillCRM_Settings.noCrawl && s.lists.BrillCRM_AuditLog.noCrawl);
    check('Settings saved with the install date (trial clock)', s.lists.BrillCRM_Settings.items.length === 1 && /installDate/.test(s.lists.BrillCRM_Settings.items[0].Value));
    await h.shot('01-dashboard-empty');
    await h.go('asha', 'setup');
    const t = await h.text('.crm-main');
    check('Setup & health: every list is ready after setup', !/Will be created|missing/i.test(t), t.slice(0, 200));
  });

  // ---------------------------------------------------------------- 2. products
  await step('Products', async () => {
    await h.go('asha', 'entity:products');
    for (const [name, sku, price, cycle] of [['HR Suite', 'HR-01', 120, 'Annual'], ['Onboarding service', 'SRV-01', 2500, 'One-time']]) {
      await h.click('.crm-main button:has-text("New product")');
      await h.fill('Product name', name); await h.fill('SKU', sku); await h.fill('Unit price', price);
      await h.choose('Billing', cycle);
      await h.save();
    }
    const items = await h.items('BrillCRM_Products');
    check('Two products saved with price and billing cycle', items.length === 2 && items.some((x) => x.UnitPrice === 120 && x.BillingCycle === 'Annual'), JSON.stringify(items.map((x) => [x.Title, x.UnitPrice, x.BillingCycle])));
  });

  // ---------------------------------------------------------------- 3. leads
  await step('Leads', async () => {
    await h.go('asha', 'entity:leads');
    await h.click('.crm-main button:has-text("New lead")');
    await h.fill('First name', 'Priya'); await h.fill('Last name', 'Shah'); await h.fill('Email', 'priya@contoso.com');
    await h.fill('Phone', '+91 98450 12345'); await h.fill('Company', 'Contoso Ltd'); await h.fill('Job title', 'CTO'); await h.fill('Country', 'India');
    await h.choose('Lead source', 'Event'); await h.choose('Rating', 'Hot');
    await h.save();
    const [lead] = await h.items('BrillCRM_Leads');
    check('Lead saved: name, score, owner and WhatsApp number', lead && lead.Title === 'Priya Shah' && lead.LeadScore > 0 && lead.LeadOwnerId === 7 && lead.Whatsapp === '919845012345', JSON.stringify(lead));
    check('Lead website filled from the business email', lead && lead.Website === 'https://contoso.com');
    // duplicate warning
    await h.click('.crm-main button:has-text("New lead")');
    await h.fill('First name', 'Priya'); await h.fill('Last name', 'Shah'); await h.fill('Email', 'PRIYA@contoso.com');
    await p.waitForTimeout(800);
    const dupText = await h.text('.ms-Panel.is-open');
    check('Duplicate lead is flagged while typing', /duplicate|already exists|possible match/i.test(dupText), dupText.slice(0, 160));
    await h.panel().locator('button:has-text("Cancel")').click();
    await p.waitForTimeout(500);
    await h.shot('02-leads');
  });

  // ---------------------------------------------------------------- 4. convert
  await step('Convert lead', async () => {
    await h.go('asha', 'entity:leads');
    await h.openRecord('Priya Shah');
    await h.click('.ms-Panel button:has-text("Convert")');
    const dlg = p.locator('.ms-Dialog-main');
    await dlg.getByLabel('Opportunity name', { exact: true }).fill('Contoso HR portal');
    await dlg.getByLabel('Estimated amount (USD)', { exact: true }).fill('48000');
    await dlg.locator('button:has-text("Convert")').click();
    await p.waitForTimeout(2500);
    const s = await h.state();
    const acc = s.lists.BrillCRM_Accounts.items[0];
    const con = s.lists.BrillCRM_Contacts.items[0];
    const opp = s.lists.BrillCRM_Opportunities.items[0];
    const lead = s.lists.BrillCRM_Leads.items[0];
    check('Conversion creates account, contact and opportunity', acc && con && opp, `${acc && acc.Title} / ${con && con.Title} / ${opp && opp.Title}`);
    check('Contact and opportunity link to the account', con && opp && con.AccountRef === acc.Id && opp.AccountRef === acc.Id && opp.ContactRef === con.Id);
    check('Lead marked Converted with references', lead.Status === 'Converted' && lead.AccountRef === acc.Id && lead.OpportunityRef === opp.Id, JSON.stringify({ s: lead.Status, a: lead.AccountRef, o: lead.OpportunityRef }));
    check('Opportunity gets amount, stage and probability', opp && opp.Amount === 48000 && !!opp.Stage && opp.Probability > 0, JSON.stringify({ a: opp && opp.Amount, s: opp && opp.Stage, p: opp && opp.Probability }));
  });

  // ---------------------------------------------------------------- 5. opportunity stages
  await step('Opportunity stages', async () => {
    await h.go('asha', 'entity:opportunities');
    await h.openRecord('Contoso HR portal');
    await h.choose('Stage', 'Negotiation');
    await h.save('Save & close');
    let [opp] = await h.items('BrillCRM_Opportunities');
    check('Changing the stage updates probability and forecast category', opp.Stage === 'Negotiation' && opp.Probability === 75 && opp.ForecastCategory === 'Commit', JSON.stringify({ s: opp.Stage, p: opp.Probability, f: opp.ForecastCategory }));
    await h.openRecord('Contoso HR portal');
    await h.panel().locator('button:has-text("Lost")').first().click();
    const dlg = p.getByRole('dialog', { name: 'Why was this deal lost?' });
    await dlg.waitFor({ timeout: 5000 }).catch(() => undefined);
    check('Marking a deal lost asks for the lost reason', await p.isVisible('text=Why was this deal lost?'));
    check('"Mark as lost" stays disabled until a reason is chosen', await dlg.locator('button:has-text("Mark as lost")').isDisabled());
    await h.choose('Lost reason', 'Price', dlg);
    await dlg.getByLabel(/^Competitor/).fill('Rival CRM');
    await dlg.locator('button:has-text("Mark as lost")').click();
    await p.waitForTimeout(1800);
    [opp] = await h.items('BrillCRM_Opportunities');
    check('Lost deal saved with reason, competitor, 0% and Omitted', opp.Stage === 'Closed Lost' && opp.LostReason === 'Price' && /Rival CRM/.test(opp.Competitors || '') && opp.Probability === 0 && opp.ForecastCategory === 'Omitted', JSON.stringify({ s: opp.Stage, r: opp.LostReason, c: opp.Competitors, p: opp.Probability, f: opp.ForecastCategory }));
    if (await h.panel().isVisible()) { await h.panel().locator('button:has-text("Cancel")').click(); await p.waitForTimeout(400); }
    check('Lost deal gets an actual close date', !!opp.ActualCloseDate);
  });

  // ---------------------------------------------------------------- 6. quote to cash
  await step('Quote to cash', async () => {
    await h.go('asha', 'entity:opportunities');
    await h.click('.crm-main button:has-text("New opportunity")');
    await h.fill('Deal name', 'Contoso expansion'); await h.fill('Amount', '20000');
    await h.save();
    await h.openRecord('Contoso expansion');
    await h.command('New quote');
    await h.fill('Quote name', 'Contoso expansion quote');
    await h.panel().locator('button:has-text("Line items")').click(); await p.waitForTimeout(500);
    await h.panel().locator('button:has-text("Add line")').click(); await p.waitForTimeout(400);
    const product = h.panel().getByRole('combobox', { name: 'Product' }).last();
    await product.click(); await p.waitForTimeout(300);
    await p.getByRole('option', { name: /HR Suite/ }).first().click(); await p.waitForTimeout(300);
    await h.panel().getByLabel('Quantity').last().fill('50');
    await h.panel().getByLabel('Discount percent').last().fill('20');
    await p.waitForTimeout(300);
    const totals = await h.text('.ms-Panel.is-open');
    check('Line totals: 50 × $120 − 20% = $4,800', /Total:\s*\$4,800/.test(totals), (totals.match(/Subtotal[^C]*/) || [''])[0]);
    await h.save('Create');
    let [q] = await h.items('BrillCRM_Quotes');
    check('Quote numbered Q-<year>-0001 and linked to the deal', q && /^Q-\d{4}-0001$/.test(q.QuoteNumber) && q.OpportunityRef > 0, q && q.QuoteNumber);
    check('Quote totals saved', q && q.Subtotal === 6000 && q.DiscountTotal === 1200 && q.GrandTotal === 4800, JSON.stringify({ s: q.Subtotal, d: q.DiscountTotal, g: q.GrandTotal }));
    const pt = await h.text('.ms-Panel.is-open');
    check('20% discount is over the 15% limit: approval needed', /above the 15% limit/.test(pt));
    check('"Accepted" is disabled until the discount is approved', await h.commandDisabled('Accepted'));
    await h.command('Request approval');
    [q] = await h.items('BrillCRM_Quotes');
    check('Requesting approval sets Approval = Pending', q.ApprovalStatus === 'Pending');
    // The "Quote discount approval" flow records the manager's decision.
    await h.flow('BrillCRM_Quotes', q.Id, { ApprovalStatus: 'Approved', ApprovalNote: 'OK for a 3-year deal' });
    await h.go('asha', 'entity:quotes');
    await h.openRecord('Contoso expansion quote');
    check('After approval, "Accepted" is enabled', !(await h.commandDisabled('Accepted')));
    await h.command('Accepted'); await p.waitForTimeout(1500);
    const s = await h.state();
    q = s.lists.BrillCRM_Quotes.items[0];
    const opp = s.lists.BrillCRM_Opportunities.items.find((x) => x.Title === 'Contoso expansion');
    const subs = s.lists.BrillCRM_Subscriptions.items;
    check('Accepting the quote marks it Accepted', q.QuoteStatus === 'Accepted');
    check('…closes the deal as won with the quote total', opp.Stage === 'Closed Won' && opp.Amount === 4800 && !!opp.ActualCloseDate, JSON.stringify({ s: opp.Stage, a: opp.Amount }));
    check('…and creates the subscription for the annual line', subs.length === 1 && subs[0].Quantity === 50 && subs[0].SubStatus === 'Active' && !!subs[0].RenewalDate, JSON.stringify(subs.map((x) => [x.Title, x.Quantity, x.SubStatus, x.RenewalDate, x.Mrr, x.Arr])));
    check('Subscription MRR/ARR from the discounted price', subs[0] && subs[0].Arr === 4800 && subs[0].Mrr === 400, subs[0] && `${subs[0].Mrr} / ${subs[0].Arr}`);
    await h.command('Create invoice'); await p.waitForTimeout(1000);
    let [inv] = await h.items('BrillCRM_Invoices');
    check('Invoice created from the quote: INV-<year>-0001, total $4,800', inv && /^INV-\d{4}-0001$/.test(inv.InvoiceNumber) && inv.GrandTotal === 4800 && inv.QuoteRef === q.Id, inv && JSON.stringify({ n: inv.InvoiceNumber, t: inv.GrandTotal, d: inv.DueDate }));
    check('Invoice due date follows the payment terms', inv && inv.DueDate && inv.IssueDate && new Date(inv.DueDate) > new Date(inv.IssueDate));
    await h.command('Mark sent');
    h.answer('1800');
    await h.command('Record payment'); await p.waitForTimeout(500);
    [inv] = await h.items('BrillCRM_Invoices');
    check('Part payment: $1,800 paid, invoice still open', inv.AmountPaid === 1800 && inv.InvoiceStatus !== 'Paid', `${inv.AmountPaid} ${inv.InvoiceStatus}`);
    h.answer('3000');
    await h.command('Record payment'); await p.waitForTimeout(500);
    [inv] = await h.items('BrillCRM_Invoices');
    check('Full payment marks the invoice Paid with a paid date', inv.AmountPaid === 4800 && inv.InvoiceStatus === 'Paid' && !!inv.PaidDate, `${inv.AmountPaid} ${inv.InvoiceStatus} ${inv.PaidDate}`);
    await h.shot('03-invoice');
    await h.panel().locator('button:has-text("Cancel")').click().catch(() => undefined);
  });

  // ---------------------------------------------------------------- 7. cases and SLA
  await step('Cases', async () => {
    await h.go('asha', 'entity:cases');
    await h.click('.crm-main button:has-text("New case")');
    await h.fill('Subject', 'SSO login fails for new users');
    await h.choose('Priority', 'High');
    await h.save();
    let [c] = await h.items('BrillCRM_Cases');
    check('Case numbered CS-<year>-<id>', c && /^CS-\d{4}-\d{4,}$/.test(c.CaseNumber), c && c.CaseNumber);
    check('High priority case gets first-response and resolution SLA deadlines', c && c.SlaDue && c.FirstResponseDue && new Date(c.SlaDue) > new Date(c.FirstResponseDue), c && `${c.FirstResponseDue} / ${c.SlaDue}`);
    await h.openRecord('SSO login fails for new users');
    await h.command('Waiting on customer');
    [c] = await h.items('BrillCRM_Cases');
    check('Waiting on customer pauses the SLA clock', c.CaseStatus === 'Waiting on customer' && !!c.SlaPausedAt);
    await h.go('asha', 'entity:cases');
    await h.openRecord('SSO login fails for new users');
    await h.choose('Status', 'In progress');
    await h.save('Save & close');
    [c] = await h.items('BrillCRM_Cases');
    check('Resuming moves the SLA deadline on and clears the pause', c.CaseStatus === 'In progress' && !c.SlaPausedAt, `${c.CaseStatus} ${c.SlaPausedAt}`);
    await h.openRecord('SSO login fails for new users');
    await h.command('Resolve');
    check('Resolve asks for the resolution first', /Describe the resolution before resolving/.test(await h.text('.ms-Panel.is-open')));
    await h.fill('Resolution', 'Reset the SSO certificate.');
    await h.command('Resolve'); await p.waitForTimeout(500);
    [c] = await h.items('BrillCRM_Cases');
    check('Resolve saves status, resolved date and the resolution text', c.CaseStatus === 'Resolved' && !!c.ResolvedDate && c.Resolution === 'Reset the SSO certificate.', `${c.CaseStatus} ${c.ResolvedDate} ${c.Resolution}`);
    await h.panel().locator('button:has-text("Cancel")').click().catch(() => undefined);
  });

  // ---------------------------------------------------------------- 8. activities and My work
  await step('Activities', async () => {
    await h.go('asha', 'entity:activities');
    await h.click('.crm-main button:has-text("New activity")');
    await h.fill('Subject', 'Send renewal proposal to Contoso');
    await h.choose('Type', 'Task');
    await h.save();
    const acts = await h.items('BrillCRM_Activities');
    check('Task saved and assigned to me', acts.some((a) => a.Title === 'Send renewal proposal to Contoso' && a.ActivityType === 'Task' && a.TaskOwnerId === 7), JSON.stringify(acts.map((a) => [a.Title, a.ActivityType, a.TaskOwnerId])));
    await h.go('asha', 'mywork');
    check('My work lists the open task', /Send renewal proposal to Contoso/.test(await h.text('.crm-main')));
    await h.shot('04-mywork');
  });

  // ---------------------------------------------------------------- 9. roles and security
  await step('Security and roles', async () => {
    await h.go('asha', 'settings');
    await h.click('.crm-main button:has-text("Roles & permissions")');
    await h.click('.crm-main button:has-text("Create groups & apply list permissions")');
    await p.waitForTimeout(6000);
    let st = await h.state();
    const groups = Object.values(st.groups).map((g) => g.Title).sort();
    check('Creates the five BrillCRM_ security groups', JSON.stringify(groups) === JSON.stringify(['BrillCRM_Administrators', 'BrillCRM_Readers', 'BrillCRM_Sales', 'BrillCRM_SalesManagers', 'BrillCRM_Support']), groups.join(', '));
    check('Gives every CRM list its own permissions', Object.values(st.lists).every((l) => l.unique));
    // In a real tenant an admin now adds people to the groups in Site permissions.
    await p.evaluate(() => window.__crm.seedGroupMembers());

    await h.go('chen', 'entity:leads');
    check('Sales rep sees the CRM with the Sales role', !(await h.visible('text=read-only')));
    await h.click('.crm-main button:has-text("New lead")');
    await h.fill('First name', 'Arjun'); await h.fill('Last name', 'Mehta'); await h.fill('Company', 'Northwind Traders'); await h.fill('Email', 'arjun@northwind.com');
    await h.save();
    const arjun = (await h.items('BrillCRM_Leads')).find((x) => x.Title === 'Arjun Mehta');
    check('Sales rep creates a lead owned by them', arjun && arjun.LeadOwnerId === 9 && arjun.AuthorId === 9, arjun && `${arjun.LeadOwnerId}`);
    await h.openRecord('Arjun Mehta');
    check('Sales rep cannot delete records', !(await h.panel().locator('button:has-text("Delete")').isVisible()));
    await h.panel().locator('button:has-text("Cancel")').click(); await p.waitForTimeout(300);

    await h.go('eli', 'entity:leads');
    check('Reader has no "New lead" button', !(await h.visible('.crm-main button:has-text("New lead")')));
    await h.openRecord('Arjun Mehta');
    check('Reader sees records as read-only', /read-only access/.test(await h.text('.ms-Panel.is-open')));
    await h.panel().locator('button:has-text("Cancel"), button:has-text("Close")').first().click().catch(() => undefined);

    await h.go('dana', 'entity:cases');
    check('Support can open cases', await h.visible('.crm-main button:has-text("New case")'));
    await h.go('dana', 'entity:opportunities');
    check('Support cannot create deals', !(await h.visible('.crm-main button:has-text("New opportunity")')));

    // Record-level security
    await h.go('asha', 'settings');
    await h.click('.crm-main button:has-text("Roles & permissions")');
    await p.locator('.crm-main').getByRole('switch', { name: /Protect each record/ }).click();
    await p.waitForTimeout(300);
    await p.locator('.crm-main button:has-text("Save")').first().click(); await p.waitForTimeout(2000);
    check('After saving, existing records can be protected ("Apply now")', await h.visible('.crm-main button:has-text("Apply now")'));
    await h.click('.crm-main button:has-text("Apply now")');
    await p.waitForTimeout(8000);
    st = await h.state();
    const leads = st.lists.BrillCRM_Leads.items;
    const secured = leads.filter((x) => x.__unique);
    check('Protect each record gives every lead its own permissions', secured.length === leads.length && leads.length > 0, `${secured.length}/${leads.length}`);
    const a = leads.find((x) => x.Title === 'Arjun Mehta');
    check('…and records who it is secured for', a && String(a.SecuredOwner) === '9', a && a.SecuredOwner);
    await h.go('maya', 'entity:leads');
    const mayaSees = await h.text('.crm-main');
    check('Another sales rep cannot see Chen\'s lead', !/Arjun Mehta/.test(mayaSees));
    await h.go('chen', 'entity:leads');
    check('The owner still sees it', /Arjun Mehta/.test(await h.text('.crm-main')));
    await h.go('ben', 'entity:leads');
    check('Sales managers see every lead', /Arjun Mehta/.test(await h.text('.crm-main')));
    await h.go('ben', 'entity:accounts');
    check('…and every account', /Contoso Ltd/.test(await h.text('.crm-main')));
  });

  // ---------------------------------------------------------------- 10. settings, search, audit
  await step('Settings, search and audit', async () => {
    await h.go('asha', 'settings');
    await h.click('.crm-main button:has-text("Company profile")');
    await p.locator('.crm-main').getByLabel(/^Company name/).first().fill('Brillienta Demo Ltd');
    await p.waitForTimeout(400);
    const saveBtn = p.locator('.crm-main button:has-text("Save")').first();
    if (await saveBtn.isVisible()) { await saveBtn.click(); }
    await p.waitForTimeout(2000);
    const s = (await h.items('BrillCRM_Settings'))[0];
    check('Company profile is saved to BrillCRM_Settings', /Brillienta Demo Ltd/.test(s.Value));
    await h.go('asha', 'dashboard');
    const search = p.getByPlaceholder(/Search leads/).first();
    await search.fill('Contoso'); await p.waitForTimeout(1500);
    const hits = await p.$$eval('.crm-search-hit', (b) => b.map((x) => x.innerText.replace(/\s+/g, ' ')));
    check('Global search finds the account, contact and deals', hits.some((x) => /Contoso Ltd/.test(x)) && hits.some((x) => /Contoso expansion/.test(x)), hits.join(' | '));
    await search.fill('');
    await h.go('asha', 'audit');
    const audit = await h.items('BrillCRM_AuditLog');
    check('Audit log records creates and updates', audit.some((x) => /Created/.test(x.Title)) && audit.some((x) => /Updated/.test(x.Title)), `${audit.length} entries`);
    check('Audit log page lists them', /Contoso/.test(await h.text('.crm-main')));
  });

  // ---------------------------------------------------------------- 11. import, export, custom fields
  await step('Import, export and custom fields', async () => {
    await h.go('asha', 'entity:leads');
    await h.click('.crm-main button[aria-label="More"]');
    await p.getByRole('menuitem', { name: 'Import from CSV' }).first().click();
    await p.waitForTimeout(600);
    const csv = 'First name,Last name,Email,Company,Lead source,Rating\nNina,Patel,nina@fabrikam.com,Fabrikam,Webinar,Warm\nOmar,Haddad,omar@adatum.com,Adatum,Referral,Hot\n';
    await p.setInputFiles('input[aria-label="CSV file"]', { name: 'leads.csv', mimeType: 'text/csv', buffer: Buffer.from(csv) });
    await p.waitForTimeout(800);
    await h.shot('05-import-mapping');
    await p.click('.ms-Dialog-main button:has-text("Import")');
    await p.waitForTimeout(4000);
    const leads = await h.items('BrillCRM_Leads');
    const nina = leads.find((x) => x.Title === 'Nina Patel');
    const omar = leads.find((x) => x.Title === 'Omar Haddad');
    check('CSV import creates leads with mapped columns', nina && omar && nina.Email === 'nina@fabrikam.com' && omar.Company === 'Adatum' && omar.Rating === 'Hot', JSON.stringify([nina && [nina.Email, nina.Company, nina.LeadSource], omar && [omar.Rating, omar.LeadSource]]));
    check('Imported leads are scored', nina && nina.LeadScore > 0 && omar.LeadScore > nina.LeadScore, nina && `${nina.LeadScore} / ${omar && omar.LeadScore}`);
    if (await p.isVisible('.ms-Dialog-main button:has-text("Close")')) { await p.click('.ms-Dialog-main button:has-text("Close")'); }
    await p.keyboard.press('Escape'); await p.waitForTimeout(300);
    // export
    await h.go('asha', 'entity:leads');
    await h.click('.crm-main button[aria-label="More"]');
    const [dl] = await Promise.all([p.waitForEvent('download', { timeout: 10000 }), p.getByRole('menuitem', { name: /Export to CSV/ }).first().click()]);
    const file = require('path').join(require('./helpers.cjs').OUT, dl.suggestedFilename());
    await dl.saveAs(file);
    const text = require('fs').readFileSync(file, 'utf8');
    check('Export to CSV downloads the leads', /Nina Patel/.test(text) && /Omar Haddad/.test(text), dl.suggestedFilename());
    // custom field
    await h.go('asha', 'settings');
    await h.click('.crm-main button:has-text("Custom fields")');
    await h.choose('Record type', 'Lead', p.locator('.crm-main'));
    await p.locator('.crm-main').getByLabel('Field label').fill('Budget (INR)');
    await h.choose('Field type', 'Currency', p.locator('.crm-main'));
    await h.click('.crm-main button:has-text("Add field")');
    await p.locator('.crm-main button:has-text("Save")').first().click();
    await p.waitForTimeout(4000);
    let st = await h.state();
    const col = Object.keys(st.lists.BrillCRM_Leads.fields).find((k) => /^cf/.test(k));
    check('Adding a custom field creates the SharePoint column', !!col && st.lists.BrillCRM_Leads.fields[col].type === 'Currency', col);
    await h.go('asha', 'entity:leads');
    await h.openRecord('Nina Patel');
    await h.fill('Budget (INR)', '250000');
    await h.save('Save & close');
    st = await h.state();
    const nina2 = st.lists.BrillCRM_Leads.items.find((x) => x.Title === 'Nina Patel');
    check('Custom field value is saved on the record', nina2 && col && nina2[col] === 250000, nina2 && col && nina2[col]);
  });

  // ---------------------------------------------------------------- 12. billing maintenance
  await step('Renewals and overdue invoices', async () => {
    let st = await h.state();
    const sub = st.lists.BrillCRM_Subscriptions.items[0];
    // Make the subscription renew in 20 days and the invoice overdue (as time passing would).
    const soon = new Date(Date.now() + 20 * 86400000).toISOString();
    await h.flow('BrillCRM_Subscriptions', sub.Id, { RenewalDate: soon });
    await h.go('asha', 'entity:invoices');
    await h.click('.crm-main button:has-text("New invoice")');
    await h.fill('Invoice title', 'Consulting March');
    const acc = h.panel().getByRole('combobox', { name: 'Account' }).first();
    await acc.fill('Contoso'); await p.waitForTimeout(600);
    await p.getByRole('option', { name: /Contoso Ltd/ }).first().click().catch(() => undefined);
    await h.save();
    st = await h.state();
    const inv = st.lists.BrillCRM_Invoices.items.find((x) => x.Title === 'Consulting March');
    check('Manual invoice saved with an account', inv && inv.AccountRef > 0, inv && JSON.stringify({ a: inv.AccountRef, n: inv.InvoiceNumber }));
    if (inv) { await h.flow('BrillCRM_Invoices', inv.Id, { InvoiceStatus: 'Sent', GrandTotal: 900, DueDate: new Date(Date.now() - 10 * 86400000).toISOString() }); }
    // Maintenance runs at most every 6 hours per browser; start the clock again as if hours had passed.
    const rerun = () => p.evaluate(() => Object.keys(localStorage).filter((k) => /maintenance/.test(k)).forEach((k) => localStorage.removeItem(k)));
    await rerun();
    await h.go('ben', 'dashboard');
    await p.waitForTimeout(6000);
    st = await h.state();
    const renewal = st.lists.BrillCRM_Opportunities.items.find((x) => x.SubscriptionRef === sub.Id);
    check('Maintenance creates a renewal opportunity before the renewal date', !!renewal && renewal.DealType === 'Renewal', renewal && JSON.stringify({ t: renewal.Title, a: renewal.Amount, s: renewal.Stage, c: renewal.ExpectedCloseDate }));
    const inv2 = inv && st.lists.BrillCRM_Invoices.items.find((x) => x.Id === inv.Id);
    check('Maintenance marks past-due invoices Overdue', inv2 && inv2.InvoiceStatus === 'Overdue', inv2 && inv2.InvoiceStatus);
    await rerun();
    await h.go('ben', 'dashboard');
    await p.waitForTimeout(5000);
    st = await h.state();
    check('Running maintenance again does not duplicate the renewal', st.lists.BrillCRM_Opportunities.items.filter((x) => x.SubscriptionRef === sub.Id).length === 1);
    // A rep moves the renewal's close date by a week; the next run must still not create a second renewal.
    if (renewal) { await h.flow('BrillCRM_Opportunities', renewal.Id, { ExpectedCloseDate: new Date(Date.now() + 27 * 86400000).toISOString() }); }
    await rerun();
    await h.go('ben', 'dashboard');
    await p.waitForTimeout(5000);
    st = await h.state();
    check('…not even after its close date was moved', st.lists.BrillCRM_Opportunities.items.filter((x) => x.SubscriptionRef === sub.Id).length === 1, st.lists.BrillCRM_Opportunities.items.filter((x) => x.SubscriptionRef === sub.Id).map((x) => x.ExpectedCloseDate).join(', '));
    await h.shot('06-dashboard');
  });

  // ---------------------------------------------------------------- 13. knowledge base, campaigns, demo form
  await step('Knowledge base, campaigns and demo form', async () => {
    await h.go('dana', 'entity:articles');
    await h.click('.crm-main button:has-text("New article")');
    await h.fill('Question or title', 'How do I reset single sign-on (SSO)?');
    await h.fill('Keywords', 'sso login certificate');
    await h.choose('Status', 'Published');
    await h.fill('Answer', 'Open Settings > Security and renew the SSO certificate.');
    await h.save();
    await h.go('dana', 'entity:cases');
    await h.click('.crm-main button:has-text("New case")');
    await h.fill('Subject', 'Cannot log in with SSO');
    await h.save('Create');
    const t = await h.text('.ms-Panel.is-open');
    check('A new case suggests the matching knowledge article', /How do I reset single sign-on/.test(t));
    await h.panel().locator('button:has-text("Cancel")').click().catch(() => undefined);

    await h.go('asha', 'entity:campaigns');
    await h.click('.crm-main button:has-text("New campaign")');
    await h.fill('Campaign name', 'Expo 2026');
    await h.choose('Type', 'Event');
    await h.fill('Budget', '5000');
    await h.save();
    const camp = (await h.items('BrillCRM_Campaigns')).find((x) => x.Title === 'Expo 2026');
    check('Campaign saved with type and budget', camp && camp.CampaignType === 'Event' && camp.Budget === 5000);

    await p.goto(`${require('./helpers.cjs').URL}?as=eli&mode=demoForm`);
    await p.waitForTimeout(2500);
    await h.shot('07-demo-form');
    const form = p.locator('body');
    check('Demo form: "Request demo" disabled until the form is complete', await p.locator('button:has-text("Request demo")').isDisabled());
    await form.getByLabel(/^First name/).fill('Lena');
    await form.getByLabel(/^Last name/).fill('Gomez');
    await form.getByLabel(/^Work email/).fill('lena@woodgrove.com');
    await form.getByLabel(/^Company/).fill('Woodgrove Bank');
    await form.getByLabel(/^What would you like to see/).fill('Quotes and renewals');
    await p.click('text=I agree to be contacted about this request.');
    await p.click('button:has-text("Request demo")');
    await p.waitForTimeout(2500);
    check('A reader (no write access) gets a plain-language message, not a raw SharePoint error', /do not have permission to add items on this site/.test(await h.text('body')));
    // Someone with access to the site (here a sales rep) sends the same request.
    await p.goto(`${require('./helpers.cjs').URL}?as=chen&mode=demoForm`);
    await p.waitForTimeout(2500);
    await form.getByLabel(/^First name/).fill('Lena');
    await form.getByLabel(/^Last name/).fill('Gomez');
    await form.getByLabel(/^Work email/).fill('lena@woodgrove.com');
    await form.getByLabel(/^Company/).fill('Woodgrove Bank');
    await form.getByLabel(/^What would you like to see/).fill('Quotes and renewals');
    await p.click('text=I agree to be contacted about this request.');
    await p.click('button:has-text("Request demo")');
    await p.waitForTimeout(2500);
    await h.shot('08-demo-form-sent');
    const st = await h.state();
    const demos = st.lists.BrillCRM_DemoRequests.items;
    check('Public demo form creates a demo request', demos.some((x) => x.ContactName === 'Lena Gomez' && x.Company === 'Woodgrove Bank' && x.DemoStatus === 'Requested'), JSON.stringify(demos.map((x) => [x.ContactName, x.Company, x.DemoStatus])));
    const lena = st.lists.BrillCRM_Leads.items.find((x) => x.Email === 'lena@woodgrove.com');
    check('…and a lead from the demo form, linked to the demo request', lena && lena.LeadSource === 'Website Demo Form' && demos.some((x) => x.LeadRef === lena.Id), lena && JSON.stringify({ s: lena.LeadSource, demoLead: demos.map((x) => x.LeadRef) }));
    check('Agreeing to be contacted about the request is not taken as marketing consent (GDPR)', lena && !lena.MarketingConsent);
  });

  // ---------------------------------------------------------------- 14. every page for every role
  await step('All pages', async () => {
    const pages = ['dashboard', 'mywork', 'calendar', 'forecast', 'reports', 'audit', 'settings', 'setup',
      'entity:leads', 'entity:accounts', 'entity:contacts', 'entity:opportunities', 'entity:quotes', 'entity:subscriptions', 'entity:invoices',
      'entity:demos', 'entity:campaigns', 'entity:cases', 'entity:articles', 'entity:products', 'entity:activities'];
    for (const who of ['asha', 'ben', 'chen', 'dana', 'eli']) {
      const before = h.errors.length;
      const bad = [];
      for (const pg of pages) {
        await p.goto(`${require('./helpers.cjs').URL}?as=${who}&page=${encodeURIComponent(pg)}`);
        await p.waitForSelector('.crm-nav-inner', { timeout: 20000 });
        await p.waitForTimeout(700);
        const t = await h.text('.crm-main');
        if (/Not implemented in the preview|does not exist|Something went wrong|TypeError|undefined is not|NaN/.test(t)) { bad.push(`${pg}: ${(t.match(/(Not implemented[^.]*|does not exist[^.]*|Something went wrong[^.]*|TypeError[^.]*|NaN[^ ]*)/) || [''])[0]}`); }
      }
      check(`Every page opens cleanly for ${who}`, !bad.length && h.errors.length === before, bad.concat(h.errors.slice(before)).join(' | '));
    }
  });

  // ---------------------------------------------------------------- 15. privacy, merge, pipeline board
  await step('Privacy, merge and pipeline board', async () => {
    await h.go('asha', 'entity:leads');
    await h.openRecord('Omar Haddad');
    await h.command('Privacy');
    h.answer('ERASE');
    await p.getByRole('menuitem', { name: /Erase personal data/ }).first().click();
    await p.waitForTimeout(3000);
    const omar = (await h.items('BrillCRM_Leads')).find((x) => x.Id && /Omar|Erased|\[erased\]/i.test(x.Title || '') || x.Company === 'Adatum');
    check('GDPR erase removes the person\'s name, email and phone', omar && !/Omar/.test(omar.Title || '') && !omar.Email && !omar.Phone, omar && JSON.stringify({ t: omar.Title, e: omar.Email, f: omar.FirstName }));
    const audit = await h.items('BrillCRM_AuditLog');
    check('…takes the erased lead out of follow-up (Unqualified, no next follow-up)', omar && omar.Status === 'Unqualified' && !omar.NextFollowUp, omar && omar.Status);
    check('…and deletes its change history', omar && !audit.some((x) => x.EntityType === 'leads' && x.RecordId === omar.Id && /Omar/.test(x.Changes || x.RecordTitle || '')));
    await h.panel().locator('button:has-text("Cancel")').click().catch(() => undefined);

    // merge: a second contact for the same person
    await h.go('asha', 'entity:contacts');
    await h.click('.crm-main button:has-text("New contact")');
    await h.fill('First name', 'Priya'); await h.fill('Last name', 'Shah'); await h.fill('Email', 'priya@contoso.com'); await h.fill('Job title', 'Chief Technology Officer');
    await h.save();
    let cons = await h.items('BrillCRM_Contacts');
    check('Duplicate contact can be created (with a warning)', cons.filter((x) => x.Email === 'priya@contoso.com').length === 2);
    const keep = cons.find((x) => x.Email === 'priya@contoso.com' && x.AccountRef);
    await p.click(`.crm-main >> text="Priya Shah" >> nth=0`); await p.waitForTimeout(1200);
    const mergeLink = h.panel().locator('text=merge into this').first();
    check('Duplicate warning offers "merge into this"', await mergeLink.isVisible());
    await mergeLink.click(); await p.waitForTimeout(800);
    await p.locator('.ms-Dialog-main button:has-text("Merge")').first().click();
    await p.waitForTimeout(4000);
    cons = await h.items('BrillCRM_Contacts');
    const left = cons.filter((x) => x.Email === 'priya@contoso.com');
    check('Merging leaves one contact', left.length === 1, left.map((x) => x.Id).join(','));
    const opps = await h.items('BrillCRM_Opportunities');
    check('…and the deals still point to the remaining contact', left.length === 1 && opps.filter((x) => x.ContactRef).every((x) => cons.some((c) => c.Id === x.ContactRef)), JSON.stringify(opps.map((x) => x.ContactRef)));
    check('…keeping the account link', left.length === 1 && left[0].AccountRef === (keep && keep.AccountRef));
    await h.panel().locator('button:has-text("Cancel")').click().catch(() => undefined);

    // pipeline board drag and drop
    await h.go('ben', 'entity:opportunities');
    const card = p.locator('.crm-card-item').first();
    const title = (await card.innerText()).split('\n')[0];
    const lane = p.locator('.crm-main >> text=Negotiation').first();
    await card.dragTo(lane); await p.waitForTimeout(2000);
    const moved = (await h.items('BrillCRM_Opportunities')).find((x) => x.Title === title);
    check('Dragging a deal to another stage saves the stage and probability', moved && moved.Stage === 'Negotiation' && moved.Probability === 75, moved && `${title}: ${moved.Stage} ${moved.Probability}`);
    await h.shot('09-pipeline-board');
  });

  // ---------------------------------------------------------------- 16. layout: phone width and dark theme
  await step('Layout', async () => {
    await p.setViewportSize({ width: 390, height: 844 });
    const over = [];
    for (const pg of ['dashboard', 'entity:leads', 'entity:opportunities', 'mywork', 'settings']) {
      await h.go('chen', pg);
      const w = await p.evaluate(() => [document.documentElement.scrollWidth, window.innerWidth]);
      if (w[0] > w[1] + 1) { over.push(`${pg} ${w[0]}>${w[1]}`); }
    }
    check('Phone width (390 px): no page scrolls sideways', !over.length, over.join(', '));
    await h.go('chen', 'entity:leads');
    await h.openRecord('Arjun Mehta');
    const pw = await p.evaluate(() => { const el = document.querySelector('.ms-Panel-main'); return el ? el.getBoundingClientRect().width : 0; });
    check('Record panel fits the phone screen', pw > 0 && pw <= 390, pw);
    await h.shot('10-phone');
    await p.setViewportSize({ width: 1366, height: 900 });
    await h.go('asha', 'dashboard', '&dark=1');
    const bg = await p.evaluate(() => { let el = document.querySelector('.crm-main'); while (el && /rgba\(0, 0, 0, 0\)|transparent/.test(getComputedStyle(el).backgroundColor)) { el = el.parentElement; } return el ? getComputedStyle(el).backgroundColor : ''; });
    check('Dark theme renders a dark background', /rgb\((\d+), (\d+), (\d+)\)/.test(bg) && Number(bg.match(/\d+/)[0]) < 80, bg);
    await h.shot('11-dark');
  });

  check('No browser errors', h.errors.length === 0, h.errors.join(' | '));
  await browser.close();
  process.exit(summary() ? 1 : 0);
})();
