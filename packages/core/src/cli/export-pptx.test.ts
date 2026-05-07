import { mkdir, mkdtemp, readFile, writeFile } from 'node:fs/promises';
import os from 'node:os';
import path from 'node:path';
import { strFromU8, unzipSync } from 'fflate';
import { describe, expect, it } from 'vitest';
import { exportPptx } from './export-pptx.ts';

describe('exportPptx CLI action', () => {
  it('loads a slide module pptx export and writes an editable pptx file', async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), 'open-slide-pptx-'));
    await mkdir(path.join(cwd, 'slides', 'demo'), { recursive: true });
    await writeFile(
      path.join(cwd, 'slides', 'demo', 'index.tsx'),
      `export default [];
export const pptx = {
  title: 'CLI PPTX',
  slides: [{ elements: [{ type: 'text', x: 64, y: 64, w: 800, h: 90, text: 'Demo from CLI', fontSize: 40 }] }],
};
`,
    );

    const output = path.join(cwd, 'demo.pptx');
    await exportPptx({ cwd, slideId: 'demo', output });

    const pptx = await readFile(output);
    const files = unzipSync(pptx);
    const slideXml = strFromU8(files['ppt/slides/slide1.xml']);
    expect(slideXml).toContain('Demo from CLI');
    expect(slideXml).toContain('<p:sp>');
    expect(slideXml).not.toContain('<p:pic>');
  });

  it('accepts a shared Slide AST export when explicit pptx is absent', async () => {
    const cwd = await mkdtemp(path.join(os.tmpdir(), 'open-slide-ast-pptx-'));
    await mkdir(path.join(cwd, 'slides', 'ast-demo'), { recursive: true });
    await writeFile(
      path.join(cwd, 'slides', 'ast-demo', 'index.tsx'),
      `export default [];
export const ast = {
  kind: 'open-slide.deck',
  title: 'AST CLI PPTX',
  slides: [
    {
      kind: 'open-slide.slide',
      children: [
        { kind: 'open-slide.text', x: 64, y: 64, w: 900, h: 90, text: 'Demo from AST', fontSize: 40 },
      ],
    },
  ],
};
`,
    );

    const output = path.join(cwd, 'ast-demo.pptx');
    await exportPptx({ cwd, slideId: 'ast-demo', output });

    const pptx = await readFile(output);
    const files = unzipSync(pptx);
    const slideXml = strFromU8(files['ppt/slides/slide1.xml']);
    expect(slideXml).toContain('Demo from AST');
    expect(slideXml).toContain('<p:sp>');
    expect(slideXml).not.toContain('<p:pic>');
  });
});
