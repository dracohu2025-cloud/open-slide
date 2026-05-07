import type { CSSProperties } from 'react';
import type { Page } from '../app/lib/sdk.ts';
import type { HexColor, PptxDeck, PptxElement } from '../pptx/exporter.ts';

export type SlideAstBox = {
  x: number;
  y: number;
  w: number;
  h: number;
};

export type SlideAstText = SlideAstBox & {
  kind: 'open-slide.text';
  text: string;
  fontSize?: number;
  fontFace?: string;
  color?: HexColor;
  bold?: boolean;
  italic?: boolean;
  align?: 'left' | 'center' | 'right';
  valign?: 'top' | 'mid' | 'bottom';
  fill?: HexColor;
};

export type SlideAstRect = SlideAstBox & {
  kind: 'open-slide.rect';
  fill?: HexColor;
  color?: HexColor;
  radius?: number;
  line?: HexColor;
  lineWidth?: number;
};

export type SlideAstLine = Omit<SlideAstBox, 'h'> & {
  kind: 'open-slide.line';
  h?: number;
  color?: HexColor;
  width?: number;
};

export type SlideAstElement = SlideAstText | SlideAstRect | SlideAstLine;

export type SlideAstSlide = {
  kind: 'open-slide.slide';
  background?: HexColor;
  children: SlideAstElement[];
};

export type SlideAstDeck = {
  kind: 'open-slide.deck';
  title?: string;
  author?: string;
  subject?: string;
  company?: string;
  slides: SlideAstSlide[];
};

export type DefineDeckInput = Omit<SlideAstDeck, 'kind'>;
export type SlideInput = Omit<SlideAstSlide, 'kind'>;
export type TextInput = Omit<SlideAstText, 'kind'>;
export type RectInput = Omit<SlideAstRect, 'kind'>;
export type LineInput = Omit<SlideAstLine, 'kind'>;

export function defineDeck(input: DefineDeckInput): SlideAstDeck {
  return { kind: 'open-slide.deck', ...input, slides: input.slides.map(cloneSlide) };
}

export function slide(input: SlideInput): SlideAstSlide {
  return cloneSlide({ kind: 'open-slide.slide', ...input });
}

export function text(input: TextInput): SlideAstText {
  return { kind: 'open-slide.text', ...input };
}

export function rect(input: RectInput): SlideAstRect {
  return { kind: 'open-slide.rect', ...input };
}

export function line(input: LineInput): SlideAstLine {
  return { kind: 'open-slide.line', ...input };
}

export function isSlideAstDeck(value: unknown): value is SlideAstDeck {
  if (!value || typeof value !== 'object') return false;
  const candidate = value as { kind?: unknown; slides?: unknown };
  return (
    candidate.kind === 'open-slide.deck' &&
    Array.isArray(candidate.slides) &&
    candidate.slides.length > 0
  );
}

export function renderPptx(deck: SlideAstDeck): PptxDeck {
  return {
    title: deck.title,
    author: deck.author,
    subject: deck.subject,
    company: deck.company,
    slides: deck.slides.map((item) => ({
      background: item.background,
      elements: item.children.map(elementToPptx),
    })),
  };
}

export function renderReact(deck: SlideAstDeck): Page[] {
  return deck.slides.map((item) => {
    const PageComponent: Page = () => (
      <div
        data-open-slide-ast="slide"
        style={{
          position: 'relative',
          width: '100%',
          height: '100%',
          overflow: 'hidden',
          background: item.background ?? 'transparent',
        }}
      >
        {item.children.map((child, index) => renderReactElement(child, index))}
      </div>
    );
    return PageComponent;
  });
}

function cloneSlide(item: SlideAstSlide): SlideAstSlide {
  return { ...item, children: item.children.map((child) => ({ ...child })) };
}

function elementToPptx(element: SlideAstElement): PptxElement {
  if (element.kind === 'open-slide.text') {
    const { kind: _kind, ...rest } = element;
    return { type: 'text', ...rest };
  }
  if (element.kind === 'open-slide.rect') {
    const { kind: _kind, ...rest } = element;
    return { type: 'rect', ...rest };
  }
  const { kind: _kind, ...rest } = element;
  return { type: 'line', ...rest };
}

function renderReactElement(element: SlideAstElement, index: number) {
  if (element.kind === 'open-slide.text') return renderText(element, index);
  if (element.kind === 'open-slide.rect') return renderRect(element, index);
  return renderLine(element, index);
}

function baseStyle(element: { x: number; y: number; w: number; h?: number }): CSSProperties {
  return {
    position: 'absolute',
    left: element.x,
    top: element.y,
    width: element.w,
    height: element.h ?? 0,
    boxSizing: 'border-box',
  };
}

function renderText(element: SlideAstText, index: number) {
  const style: CSSProperties = {
    ...baseStyle(element),
    display: 'flex',
    alignItems: valignToFlex(element.valign),
    justifyContent: alignToFlex(element.align),
    padding: element.fill ? '0.15em 0.25em' : 0,
    background: element.fill,
    color: element.color ?? '#111827',
    fontFamily: element.fontFace,
    fontSize: element.fontSize,
    fontWeight: element.bold ? 700 : undefined,
    fontStyle: element.italic ? 'italic' : undefined,
    textAlign: element.align,
    whiteSpace: 'pre-wrap',
  };
  return (
    <div data-open-slide-ast="text" key={index} style={style}>
      {element.text}
    </div>
  );
}

function renderRect(element: SlideAstRect, index: number) {
  return (
    <div
      data-open-slide-ast="rect"
      key={index}
      style={{
        ...baseStyle(element),
        background: element.fill ?? 'transparent',
        color: element.color,
        borderRadius: element.radius,
        border: element.line ? `${element.lineWidth ?? 1}px solid ${element.line}` : undefined,
      }}
    />
  );
}

function renderLine(element: SlideAstLine, index: number) {
  const width = element.width ?? 2;
  return (
    <div
      data-open-slide-ast="line"
      key={index}
      style={{
        ...baseStyle({ ...element, h: Math.max(width, Math.abs(element.h ?? 0)) }),
        borderTop: `${width}px solid ${element.color ?? '#111827'}`,
      }}
    />
  );
}

function alignToFlex(align: SlideAstText['align']): CSSProperties['justifyContent'] {
  if (align === 'center') return 'center';
  if (align === 'right') return 'flex-end';
  return 'flex-start';
}

function valignToFlex(valign: SlideAstText['valign']): CSSProperties['alignItems'] {
  if (valign === 'mid') return 'center';
  if (valign === 'bottom') return 'flex-end';
  return 'flex-start';
}
