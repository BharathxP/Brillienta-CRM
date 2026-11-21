// Jest stand-in for @microsoft/sp-http (mapped in config/jest.config.json). The real package needs
// Microsoft-internal modules that are not published, and the unit tests never call SharePoint.
export const SPHttpClient = { configurations: { v1: {} } };
