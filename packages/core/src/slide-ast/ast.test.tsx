import { strFromU8, unzipSync } from 'fflate';
import { createElement } from 'react';
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import { createPptxBuffer } from '../pptx/exporter.ts';
import {
  defineDeck,
  group,
  line,
  list,
  rect,
  renderPptx,
  renderReact,
  slide,
  table,
  text,
} from './index.tsx';

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

  it('renders grouped structures, lists, and tables to React and editable PPTX primitives', () => {
    const deck = defineDeck({
      slides: [
        slide({
          children: [
            group({
              x: 100,
              y: 120,
              w: 700,
              h: 260,
              children: [
                rect({ x: 0, y: 0, w: 700, h: 260, fill: '#111827', radius: 18 }),
                list({
                  x: 40,
                  y: 42,
                  w: 560,
                  h: 140,
                  items: ['Shared structure', 'Editable output'],
                  fontSize: 32,
                  color: '#F9FAFB',
                }),
              ],
            }),
            table({
              x: 100,
              y: 440,
              w: 760,
              h: 220,
              rows: [
                ['Renderer', 'Output'],
                ['React', 'Page[]'],
                ['PPTX', 'Shapes'],
              ],
              fontSize: 24,
              color: '#111827',
              borderColor: '#CBD5E1',
              headerFill: '#E0F2FE',
            }),
          ],
        }),
      ],
    });

    const html = renderToStaticMarkup(createElement(renderReact(deck)[0]));
    expect(html).toContain('data-open-slide-ast="group"');
    expect(html).toContain('<ul');
    expect(html).toContain('<table');
    expect(html).toContain('Editable output');

    const pptxDeck = renderPptx(deck);
    expect(pptxDeck.slides[0].elements.length).toBeGreaterThan(10);
    expect(pptxDeck.slides[0].elements).toContainEqual(
      expect.objectContaining({
        type: 'text',
        x: 140,
        y: 162,
        text: '• Shared structure\n• Editable output',
      }),
    );
    expect(pptxDeck.slides[0].elements).toContainEqual(
      expect.objectContaining({ type: 'text', text: 'Renderer' }),
    );

    const xml = slideXml(createPptxBuffer(pptxDeck));
    expect(xml).toContain('Shared structure');
    expect(xml).toContain('Renderer');
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
