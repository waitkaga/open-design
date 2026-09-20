import { createHash } from 'node:crypto';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import ts from 'typescript';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = resolve(root, 'public/vendor');
const bundle = gunzipSync(readFileSync(resolve(root, '../desktop/vendor/dom-to-pptx/dom-to-pptx.bundle.js.gz')));
const expected = '0308535fd30c30fe78df78ed7d45f2d24e7393285bf09acbb2090189b4d50558';
if (createHash('sha256').update(bundle).digest('hex') !== expected) {
  throw new Error('Unexpected dom-to-pptx vendor hash; review the upstream version before updating');
}

const modules = ['client-pptx-protocol', 'pptx-export-normalizer', 'pptx-export-bridge'];
const printer = ts.createPrinter();
const texts = modules.map((name) => {
  const file = resolve(root, `src/runtime/${name}.ts`);
  const source = ts.createSourceFile(file, readFileSync(file, 'utf8'), ts.ScriptTarget.Latest, true);
  const transformed = ts.transform(source, [(context) => (node) => {
    const visit: ts.Visitor = (child) => {
      if (ts.isImportDeclaration(child)) return undefined;
      if (child.kind === ts.SyntaxKind.ExportKeyword) return undefined;
      return ts.visitEachChild(child, visit, context);
    };
    return ts.visitNode(node, visit) as ts.SourceFile;
  }]);
  try { return printer.printFile(transformed.transformed[0]!); }
  finally { transformed.dispose(); }
});
const compiled = ts.transpileModule(texts.join('\n') + '\ninstallPptxExportBridge();', {
  compilerOptions: { target: ts.ScriptTarget.ES2022, module: ts.ModuleKind.None, removeComments: true },
  reportDiagnostics: true,
});
if (compiled.diagnostics?.length) throw new Error(ts.formatDiagnosticsWithColorAndContext(compiled.diagnostics, {
  getCurrentDirectory: () => root, getCanonicalFileName: (file) => file, getNewLine: () => '\n',
}));
const marker = '// od-client-pptx:1;dom-to-pptx:2.0.1\n';
mkdirSync(output, { recursive: true });
writeFileSync(resolve(output, 'dom-to-pptx.bundle.js'), marker + bundle.toString('utf8'), 'utf8');
writeFileSync(resolve(output, 'pptx-export-bridge.js'), marker + `(function(){\n${compiled.outputText}\n})();`, 'utf8');
writeFileSync(resolve(output, 'client-pptx.json'), JSON.stringify({ enabled: true, version: 'od-client-pptx:1;dom-to-pptx:2.0.1' }) + '\n', 'utf8');
console.log('Built browser PPTX assets (dom-to-pptx 2.0.1, verified SHA-256)');
