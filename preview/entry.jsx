// Local preview: the real Brillienta CRM (compiled lib/) against an in-memory SharePoint site.
//   ?as=asha|ben|chen|dana|eli|maya   who is signed in (admin, sales manager, sales, sales, support, reader)
//   ?reset=1                          start from an empty site (Set up CRM)
//   ?page=<page>                      open a page, e.g. dashboard, entity:leads, settings
//   ?dark=1                           dark theme       ?mode=demoForm   the public demo-request form
import * as React from 'react';
import { createRoot } from 'react-dom/client';
import { createTheme, initializeIcons } from '@fluentui/react';
import BrillientaCrm from '../lib/webparts/brillientaCrm/components/BrillientaCrm';
import { USERS, WEB, createFakeHttp, patchItem, previewState, seedGroupMembers } from './fakeSharePoint';

const params = new URLSearchParams(window.location.search);
const as = USERS[params.get('as') || 'asha'] || USERS.asha;
initializeIcons(undefined, { disableWarnings: true });
window.__crm = { state: previewState, seedGroupMembers, patchItem, user: as };
if (params.get('reset')) {
  // Keep the address clean so reloads do not wipe the site again.
  params.delete('reset');
  window.history.replaceState(null, '', `${window.location.pathname}?${params.toString()}${window.location.hash}`);
}

// ?dark=1: the dark palette the web part uses in Microsoft Teams (SharePoint passes its own theme the same way).
const darkTheme = createTheme({
  isInverted: true,
  palette: {
    themePrimary: '#19c48a', themeLighterAlt: '#010805', themeLighter: '#04201a', themeLight: '#073b2c', themeTertiary: '#0f7654', themeSecondary: '#16ad7a',
    themeDarkAlt: '#2ec995', themeDark: '#4dd2a6', themeDarker: '#7bdfbf',
    neutralLighterAlt: '#292929', neutralLighter: '#2e2e2e', neutralLight: '#3d3d3d', neutralQuaternaryAlt: '#454545', neutralQuaternary: '#4c4c4c',
    neutralTertiaryAlt: '#6b6b6b', neutralTertiary: '#c8c8c8', neutralSecondary: '#d0d0d0', neutralPrimaryAlt: '#dadada', neutralPrimary: '#ffffff',
    neutralDark: '#f4f4f4', black: '#f8f8f8', white: '#1f1f1f'
  }
});
if (params.get('dark')) { document.body.style.background = '#1f1f1f'; }

createRoot(document.getElementById('root')).render(
  <BrillientaCrm
    spHttpClient={createFakeHttp()}
    graphFactory={undefined}
    webUrl={WEB}
    hostWebUrl={WEB}
    host="sharepoint"
    tenantSiteUrl=""
    foundSiteUrl=""
    tenantId="6f1c2a9e-0b7d-4c55-9e3a-2d8f41b07c13"
    culture="en-US"
    isSiteAdmin={as.role === 'admin'}
    listNames={{}}
    mode={params.get('mode') || 'crm'}
    title="Brillienta CRM"
    fullScreen={false}
    pageEditing={false}
    defaultPage={params.get('page') || 'dashboard'}
    formTitle="Request a demo"
    formIntro=""
    productName="Brillienta CRM"
    isDarkTheme={!!params.get('dark')}
    theme={params.get('dark') ? darkTheme : undefined}
  />
);
