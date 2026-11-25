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
// Sass writes ".a :global .x, .a :global .y". SPFx's css-loader ends ":global" at each comma; esbuild carries it on,
// which leaves the second ".a" unhashed. Rewrite each selector to the ":global(.x)" form, which both read the same.
const globals = (css) => css.replace(/([^{}]+)\{/g, (m, sel) => (/:global\s/.test(sel) && !/^\s*@/.test(sel)
  ? sel.split(',').map((part) => part.replace(/:global\s+([^,]+?)\s*$/, (_, rest) => `:global(${rest})`)).join(',') + '{'
  : m));

const common = {
  bundle: true,
  format: 'iife',
  jsx: 'automatic',
  target: 'es2020',
  sourcemap: false,
  logLevel: 'warning',
  define: { 'process.env.NODE_ENV': '"production"' },
  alias: { '@microsoft/sp-http': join(HERE, 'spHttpStub.js'), '@microsoft/sp-http-msgraph': join(HERE, 'spHttpStub.js'), '@microsoft/sp-copilot-component': join(HERE, 'copilotStub.js') },
  plugins: [{
    name: 'spfx-scss-modules',
    setup(b) {
      b.onResolve({ filter: /\.module\.scss$/ }, (a) => ({ path: join(a.resolveDir, a.path) + '.css', namespace: 'scss' }));
      b.onLoad({ filter: /.*/, namespace: 'scss' }, (a) => ({ contents: globals(themeTokens(readFileSync(a.path, 'utf8'))), loader: 'local-css', resolveDir: dirname(a.path) }));
    }
  }]
};
// The CRM web part (index.html) and the Copilot cards (copilot.html).
await build({ ...common, entryPoints: [join(HERE, 'entry.jsx')], outfile: join(OUT, 'bundle.js') });
await build({ ...common, entryPoints: [join(HERE, 'copilot.jsx')], outfile: join(OUT, 'copilot.js') });
copyFileSync(join(HERE, 'index.html'), join(OUT, 'index.html'));
copyFileSync(join(HERE, 'copilot.html'), join(OUT, 'copilot.html'));
console.log(`Preview built: ${join(OUT, 'index.html')}`);
