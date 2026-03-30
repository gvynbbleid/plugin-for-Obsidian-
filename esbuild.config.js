const esbuild = require('esbuild');
const svelte = require('esbuild-svelte');

const production = process.argv.includes('production');

esbuild
  .build({
    entryPoints: ['src/main.ts'],
    bundle: true,
    external: ['obsidian', 'electron', 'fs', 'path'],
    format: 'cjs',
    target: 'es2018',
    logLevel: 'info',
    sourcemap: production ? false : 'inline',
    treeShaking: true,
    outfile: 'main.js',
    platform: 'node',
    plugins: [
      svelte({
        compilerOptions: {
          dev: !production,
          css: 'injected'
        }
      })
    ],
    define: {
      'process.env.NODE_ENV': production ? '"production"' : '"development"',
      'global': 'globalThis'
    }
  })
  .catch(() => process.exit(1));
