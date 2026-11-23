/* Brillienta CRM for Outlook — task pane.
 * Shows the CRM record of the email's sender (contact or lead) with open deals and cases, logs the email
 * on the record ("Track email"), and creates a lead for unknown senders. Talks to the customer's own
 * SharePoint CRM site with the signed-in user's permissions (Microsoft Entra, nested app authentication). */
(function () {
  'use strict';
  var CFG = window.BRILLIENTA_ADDIN_CONFIG || {};
  var app = document.getElementById('app');
  var pca, account, siteUrl, pageUrl;
  var FREE = ['gmail.com', 'googlemail.com', 'yahoo.com', 'hotmail.com', 'outlook.com', 'live.com', 'icloud.com', 'aol.com', 'proton.me', 'protonmail.com', 'zoho.com', 'gmx.com', 'rediffmail.com'];

  function el(tag, attrs, children) {
    var e = document.createElement(tag);
    Object.keys(attrs || {}).forEach(function (k) { if (k === 'text') { e.textContent = attrs[k]; } else if (k.indexOf('on') === 0) { e.addEventListener(k.slice(2), attrs[k]); } else { e.setAttribute(k, attrs[k]); } });
    (children || []).forEach(function (c) { if (c) { e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); } });
    return e;
  }
  function show(nodes) { app.innerHTML = ''; nodes.forEach(function (n) { if (n) { app.appendChild(n); } }); }
  function msg(text, cls) { return el('p', { 'class': cls || 'muted', text: text }); }
  function esc(s) { return String(s == null ? '' : s).replace(/'/g, "''"); }

  // ---------------------------------------------------------------- settings
  function loadSettings() {
    var rs = Office.context.roamingSettings;
    siteUrl = (rs.get('crmSiteUrl') || '').replace(/\/+$/, '');
    pageUrl = rs.get('crmPageUrl') || '';
  }
  function settingsForm(reason) {
    var site = el('input', { id: 'site', type: 'url', placeholder: 'https://contoso.sharepoint.com/sites/crm', value: siteUrl || '' });
    var page = el('input', { id: 'page', type: 'url', placeholder: 'https://contoso.sharepoint.com/sites/crm/SitePages/CRM.aspx', value: pageUrl || '' });
    var status = msg('');
    show([
      el('h2', { text: 'Connect to your CRM' }), reason ? msg(reason) : null,
      el('label', { 'for': 'site', text: 'CRM site URL' }), site,
      el('label', { 'for': 'page', text: 'CRM page URL (optional)' }), page,
      el('div', { 'class': 'actions' }, [el('button', { type: 'button', text: 'Save', onclick: function () {
        if (!/^https:\/\/[^/]+\.sharepoint\.com\//i.test(site.value.trim())) { status.textContent = 'Enter the address of the SharePoint site that hosts Brillienta CRM.'; status.className = 'err'; return; }
        var rs = Office.context.roamingSettings;
        rs.set('crmSiteUrl', site.value.trim()); rs.set('crmPageUrl', page.value.trim());
        rs.saveAsync(function () { loadSettings(); render(); });
      } })]), status
    ]);
  }

  // ---------------------------------------------------------------- auth + SharePoint REST
  function authInit() {
    if (pca) { return Promise.resolve(); }
    var conf = { auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/organizations' }, cache: { cacheLocation: 'localStorage' } };
    var create = msal.createNestablePublicClientApplication ? msal.createNestablePublicClientApplication(conf) : Promise.reject(new Error('no NAA'));
    return create.then(function (c) { pca = c; }).catch(function () {
      pca = new msal.PublicClientApplication(conf);
      return pca.initialize();
    });
  }
  function token() {
    var scopes = [new URL(siteUrl).origin + '/AllSites.Write'];
    return authInit().then(function () {
      account = account || pca.getActiveAccount() || pca.getAllAccounts()[0];
      var req = { scopes: scopes, account: account };
      return (account ? pca.acquireTokenSilent(req) : Promise.reject(new Error('sign in'))).catch(function () {
        return pca.acquireTokenPopup({ scopes: scopes, loginHint: Office.context.mailbox.userProfile.emailAddress });
      });
    }).then(function (r) { account = r.account; if (pca.setActiveAccount) { pca.setActiveAccount(r.account); } return r.accessToken; });
  }
  function sp(path, method, body) {
    return token().then(function (t) {
      return fetch(siteUrl + path, {
        method: method || 'GET',
        headers: { Authorization: 'Bearer ' + t, Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
        body: body ? JSON.stringify(body) : undefined
      });
    }).then(function (res) {
      if (!res.ok) { return res.text().then(function (t) { throw new Error('SharePoint ' + res.status + ': ' + t.slice(0, 200)); }); }
      return res.status === 204 ? {} : res.json();
    });
  }
  function items(list, query) { return sp("/_api/web/lists/getbytitle('" + encodeURIComponent(list) + "')/items?" + query).then(function (r) { return r.value || []; }); }

  // ---------------------------------------------------------------- email context
  function sender() {
    var item = Office.context.mailbox.item;
    var f = item.from || item.sender || {};
    return { email: String(f.emailAddress || '').toLowerCase(), name: f.displayName || '', subject: item.subject || '', id: item.internetMessageId || item.itemId };
  }
  function bodyText() {
    return new Promise(function (resolve) {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (r) { resolve(r.status === Office.AsyncResultStatus.Succeeded ? String(r.value || '') : ''); });
    });
  }

  // ---------------------------------------------------------------- actions
  function track(kind, rec, s) {
    var ext = ('mail:' + s.id + '|' + (kind === 'Contact' ? 'contacts' : 'leads') + ':' + rec.Id).slice(0, 255);
    return items('BrillCRM_Activities', "$select=Id&$top=1&$filter=ExternalId eq '" + esc(ext) + "'").then(function (found) {
      if (found.length) { return 'already'; }
      return bodyText().then(function (text) {
        return sp("/_api/web/lists/getbytitle('BrillCRM_Activities')/items", 'POST', {
          Title: ('Received: ' + (s.subject || '(no subject)')).slice(0, 255), ActivityType: 'Email', RelatedType: kind, RelatedId: rec.Id,
          Completed: true, DueDate: new Date().toISOString(), ActivitySource: 'Outlook', ExternalId: ext,
          Details: ('From: ' + s.name + ' <' + s.email + '>\n\n' + text).slice(0, 60000)
        });
      }).then(function () { return 'tracked'; });
    });
  }
  function createLead(s) {
    var parts = (s.name || s.email.split('@')[0]).trim().split(/\s+/);
    var domain = s.email.split('@')[1] || '';
    var business = domain && FREE.indexOf(domain) < 0;
    var company = business ? domain.split('.')[0].charAt(0).toUpperCase() + domain.split('.')[0].slice(1) : '';
    return sp("/_api/web/lists/getbytitle('BrillCRM_Leads')/items", 'POST', {
      Title: s.name || s.email, FirstName: parts[0], LastName: parts.slice(1).join(' '), Email: s.email, Company: company,
      Website: business ? 'https://' + domain : '', LeadSource: 'Other', Status: 'New', Rating: 'Warm', EnrichStatus: business ? 'Requested' : null,
      Description: 'Created from Outlook: ' + s.subject
    });
  }

  // ---------------------------------------------------------------- view
  function render() {
    if (!CFG.clientId || /^0{8}-/.test(CFG.clientId)) { show([msg('This add-in is not configured yet (missing app registration). Please contact your administrator.', 'err')]); return; }
    if (!siteUrl) { settingsForm('Enter your Brillienta CRM site once; it is remembered for your mailbox.'); return; }
    var s = sender();
    if (!s.email) { show([msg('Open an email to see the sender in the CRM.')]); return; }
    show([msg('Looking up ' + s.email + '…')]);
    var sel = '$select=Id,Title,Email,JobTitle,Company,AccountRef,Status,Phone&$top=1&$filter=Email eq \'' + esc(s.email) + '\'';
    Promise.all([items('BrillCRM_Contacts', sel.replace(',Company', '').replace(',Status', '')), items('BrillCRM_Leads', sel.replace(',AccountRef', ''))]).then(function (r) {
      var contact = r[0][0], lead = r[1][0];
      var rec = contact || lead, kind = contact ? 'Contact' : 'Lead';
      var status = msg('');
      var head = [el('h2', { text: rec ? rec.Title : (s.name || s.email) }), msg(s.email)];
      if (!rec) {
        show(head.concat([
          msg('This sender is not in the CRM yet.'),
          el('div', { 'class': 'actions' }, [el('button', { type: 'button', text: 'Create lead and track email', onclick: function (e) {
            e.target.disabled = true; status.textContent = 'Creating…';
            createLead(s).then(function (l) { return track('Lead', l, s); }).then(function () { render(); }).catch(function (x) { status.textContent = x.message; status.className = 'err'; e.target.disabled = false; });
          } })]), status
        ]));
        return;
      }
      var details = el('div', { 'class': 'card' }, [
        el('div', { 'class': 'row' }, [el('span', { 'class': 'muted', text: 'Type' }), el('span', { 'class': 'pill', text: kind + (lead && lead.Status ? ' · ' + lead.Status : '') })]),
        rec.JobTitle ? el('div', { 'class': 'row' }, [el('span', { 'class': 'muted', text: 'Job title' }), el('span', { text: rec.JobTitle })]) : null,
        rec.Phone ? el('div', { 'class': 'row' }, [el('span', { 'class': 'muted', text: 'Phone' }), el('a', { href: 'tel:' + rec.Phone, text: rec.Phone })]) : null
      ]);
      var lists = el('div', {}, [msg('Loading deals and cases…')]);
      show(head.concat([details,
        el('div', { 'class': 'actions' }, [
          el('button', { type: 'button', text: 'Track email', onclick: function (e) {
            e.target.disabled = true; status.textContent = 'Saving…'; status.className = 'muted';
            track(kind, rec, s).then(function (res) { status.textContent = res === 'already' ? 'This email is already on the timeline.' : 'Email saved to the ' + kind.toLowerCase() + '’s timeline.'; status.className = 'ok'; })
              .catch(function (x) { status.textContent = x.message; status.className = 'err'; e.target.disabled = false; });
          } }),
          pageUrl ? el('button', { type: 'button', 'class': 'secondary', text: 'Open CRM', onclick: function () { window.open(pageUrl, '_blank'); } }) : null
        ]), status, lists]));
      if (kind !== 'Contact') { lists.innerHTML = ''; return; }
      Promise.all([
        items('BrillCRM_Opportunities', "$select=Id,Title,Stage,Amount,Currency,ExpectedCloseDate&$top=5&$orderby=Modified desc&$filter=ContactRef eq " + rec.Id),
        items('BrillCRM_Cases', "$select=Id,Title,CaseNumber,CaseStatus&$top=5&$orderby=Modified desc&$filter=ContactRef eq " + rec.Id)
      ]).then(function (x) {
        var opps = x[0].filter(function (o) { return !/^Closed/.test(o.Stage || ''); });
        var cases = x[1].filter(function (c) { return c.CaseStatus !== 'Closed' && c.CaseStatus !== 'Resolved'; });
        lists.innerHTML = '';
        lists.appendChild(el('h3', { text: 'Open deals' }));
        lists.appendChild(opps.length ? el('div', { 'class': 'card' }, opps.map(function (o) {
          return el('div', { 'class': 'row' }, [el('span', { text: o.Title }), el('span', { 'class': 'muted', text: (o.Amount != null ? Number(o.Amount).toLocaleString() + ' ' + (o.Currency || '') : '') + ' · ' + (o.Stage || '') })]);
        })) : msg('None'));
        lists.appendChild(el('h3', { text: 'Open cases' }));
        lists.appendChild(cases.length ? el('div', { 'class': 'card' }, cases.map(function (c) {
          return el('div', { 'class': 'row' }, [el('span', { text: (c.CaseNumber ? c.CaseNumber + ' ' : '') + c.Title }), el('span', { 'class': 'pill', text: c.CaseStatus || '' })]);
        })) : msg('None'));
      }).catch(function () { lists.innerHTML = ''; });
    }).catch(function (e) {
      show([msg(e.message, 'err'), el('div', { 'class': 'actions' }, [el('button', { type: 'button', 'class': 'secondary', text: 'Check settings', onclick: function () { settingsForm(); } }), el('button', { type: 'button', text: 'Try again', onclick: render })])]);
    });
  }

  Office.onReady(function () {
    loadSettings();
    document.getElementById('settings').addEventListener('click', function () { settingsForm(); });
    if (Office.context.mailbox.addHandlerAsync && Office.EventType.ItemChanged) {
      Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, function () { render(); });
    }
    render();
  });
})();
