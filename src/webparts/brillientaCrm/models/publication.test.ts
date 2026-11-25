/* eslint-disable @typescript-eslint/no-explicit-any, @typescript-eslint/no-require-imports */
// Microsoft Marketplace / SharePoint Store details for both packages: what the listing validation checks.

interface INodeFs { readFileSync(file: string, encoding?: 'utf8'): any; existsSync(file: string): boolean }
interface INodePath { resolve(...parts: string[]): string; join(...parts: string[]): string }
declare const __dirname: string;
const fs: INodeFs = require('fs');
const path: INodePath = require('path');

// Tests run from lib-commonjs/webparts/brillientaCrm/models.
const ROOT = path.resolve(__dirname, '../../../..');
const json = (p: string): any => JSON.parse(fs.readFileSync(path.join(ROOT, p), 'utf8'));

/** Width and height of a PNG file (from its IHDR header). */
function pngSize(file: string): { w: number; h: number } {
  const b = fs.readFileSync(file);
  return { w: b.readUInt32BE(16), h: b.readUInt32BE(20) };
}

const PACKAGES = [
  { name: 'Brillienta CRM', config: 'config/package-solution.json', images: 'sharepoint/images', screenshots: 5 },
  { name: 'Brillienta CRM Copilot', config: 'copilot-solution/config/package-solution.json', images: 'copilot-solution/images', screenshots: 4 }
];

describe.each(PACKAGES)('$name package details', ({ name, config, images, screenshots }) => {
  const s = json(config).solution;

  it('has the listing name, Partner ID and brillienta.com links', () => {
    expect(s.name).toBe(name);
    expect(s.developer.name).toBe('Brillienta');
    expect(s.developer.mpnId).toMatch(/^\d{6,}$/);
    expect(s.developer.websiteUrl).toBe('https://brillienta.com/crm');
    expect(s.developer.privacyUrl).toBe('https://brillienta.com/privacy');
    expect(s.developer.termsOfUseUrl).toBe('https://brillienta.com/terms');
  });

  it('has descriptions within the limits, ending with Get started, Help and Contact us links', () => {
    const short = s.metadata.shortDescription.default;
    const long = s.metadata.longDescription.default;
    expect(short.length).toBeLessThanOrEqual(80);
    expect(long.length).toBeLessThanOrEqual(4000);
    ['Get started', 'Help', 'Contact us', 'https://brillienta.com/crm', 'https://brillienta.com/guides/crm/', 'https://brillienta.com/contact', 'support@brillienta.com']
      .forEach(t => expect(long).toContain(t));
    expect(s.metadata.categories.length).toBeGreaterThan(0);
    expect(s.metadata.categories.length).toBeLessThanOrEqual(3);
  });

  it('packages a 96 x 96 app icon and 1366 x 768 screenshots', () => {
    const icon = path.join(ROOT, 'sharepoint', s.iconPath);
    expect(fs.existsSync(icon)).toBe(true);
    expect(pngSize(icon)).toEqual({ w: 96, h: 96 });
    expect(s.metadata.screenshotPaths.length).toBe(screenshots);
    s.metadata.screenshotPaths.forEach((p: string) => {
      const file = path.join(ROOT, images, p.replace(/^images\//, ''));
      expect(fs.existsSync(file)).toBe(true);
      expect(pngSize(file)).toEqual({ w: 1366, h: 768 });
    });
  });

  it('keeps the feature version in step with the solution version', () => {
    expect(s.features[0].version).toBe(s.version);
  });
});

describe('Web part picker', () => {
  it('shows the app icon, not a font icon', () => {
    const text: string = fs.readFileSync(path.join(ROOT, 'src/webparts/brillientaCrm/BrillientaCrmWebPart.manifest.json'), 'utf8');
    const icon = 'data:image/png;base64,' + fs.readFileSync(path.join(ROOT, 'sharepoint/images/AppIcon.png')).toString('base64');
    expect(text).toContain(`"iconImageUrl": "${icon}"`);
    expect(text).not.toContain('officeFabricIconFontName');
  });
});
