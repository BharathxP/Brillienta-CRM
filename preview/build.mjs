// Builds the local preview (npm run preview): the compiled CRM from lib/ with SharePoint replaced by an in-memory
// fake (fakeSharePoint.js). Output: preview-dist/index.html (open it with file://).
import { build } from 'esbuild';
import { copyFileSync, mkdirSync, readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));
const OUT = join(HERE, '..', 'preview-dist');
mkdirSync(OUT, { recursive: true });

// SPFx resolves "[theme:slot, default: #hex]" tokens at run time; the preview uses the defaults.
const themeTokens = (css) => css.replace(/"\[theme:\s*[\w-]+\s*,\s*default:\s*([^\]]+?)\s*\]"/g, '$1');

await build({
  entryPoints: [join(HERE, 'entry.jsx')],
  bundle: true,
  outfile: join(OUT, 'bundle.js'),
  format: 'iife',
  jsx: 'automatic',
  target: 'es2020',
  sourcemap: false,
  logLevel: 'warning',
  define: { 'process.env.NODE_ENV': '"production"' },
  alias: { '@microsoft/sp-http': join(HERE, 'spHttpStub.js'), '@microsoft/sp-http-msgraph': join(HERE, 'spHttpStub.js') },
  plugins: [{
    name: 'spfx-scss-modules',
    setup(b) {
      b.onResolve({ filter: /\.module\.scss$/ }, (a) => ({ path: join(a.resolveDir, a.path) + '.css', namespace: 'scss' }));
      b.onLoad({ filter: /.*/, namespace: 'scss' }, (a) => ({ contents: themeTokens(readFileSync(a.path, 'utf8')), loader: 'local-css', resolveDir: dirname(a.path) }));
    }
  }]
});
copyFileSync(join(HERE, 'index.html'), join(OUT, 'index.html'));
console.log(`Preview built: ${join(OUT, 'index.html')}`);
