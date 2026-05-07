import { strFromU8, unzipSync } from 'fflate';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import { createPptxBuffer } from '../pptx/exporter.ts';
import { defineDeck, line, rect, renderPptx, renderReact, slide, text } from './index.tsx';

function slideXml(buffer: Uint8Array): string {
  const files = unzipSync(buffer);
  return strFromU8(files['ppt/slides/slide1.xml']);
}

describe('Slide AST', () => {
  it('defines one deck that can render to React pages and editable PPTX', () => {
    const deck = defineDeck({
      title: 'Shared AST deck',
      author: 'Open Slide',
      slides: [
        slide({
          background: '#0F172A',
          children: [
            rect({ x: 96, y: 96, w: 480, h: 180, fill: '#22C55E', radius: 24 }),
            text({
              x: 128,
              y: 124,
              w: 900,
              h: 100,
              text: 'One AST, two renderers',
              fontSize: 44,
              bold: true,
              color: '#FFFFFF',
            }),
            line({ x: 128, y: 260, w: 640, h: 0, color: '#F97316', width: 4 }),
          ],
        }),
      ],
    });

    const pages = renderReact(deck);
    const html = renderToStaticMarkup(createElement(pages[0]));
    expect(html).toContain('One AST, two renderers');
    expect(html).toContain('position:absolute');
    expect(html).toContain('background:#0F172A');

    const pptxDeck = renderPptx(deck);
    expect(pptxDeck.title).toBe('Shared AST deck');
    expect(pptxDeck.slides[0].elements).toHaveLength(3);

    const xml = slideXml(createPptxBuffer(pptxDeck));
    expect(xml).toContain('One AST, two renderers');
    expect(xml).toContain('<p:sp>');
    expect(xml).not.toContain('<p:pic>');
  });

  it('keeps AST immutable enough for safe reuse across renderers', () => {
    const box = rect({ x: 10, y: 20, w: 30, h: 40, fill: '#111111' });
    const deck = defineDeck({ slides: [slide({ children: [box] })] });

    const pptxDeck = renderPptx(deck);
    pptxDeck.slides[0].elements[0].x = 999;

    expect(deck.slides[0].children[0]).toMatchObject({ x: 10, y: 20, w: 30, h: 40 });
  });
});
