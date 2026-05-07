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

export type SlideAstGroup = SlideAstBox & {
  kind: 'open-slide.group';
  children: SlideAstElement[];
};

export type SlideAstList = SlideAstBox & {
  kind: 'open-slide.list';
  items: string[];
  ordered?: boolean;
  fontSize?: number;
  fontFace?: string;
  color?: HexColor;
  bold?: boolean;
  fill?: HexColor;
  gap?: number;
};

export type SlideAstTable = SlideAstBox & {
  kind: 'open-slide.table';
  rows: string[][];
  fontSize?: number;
  fontFace?: string;
  color?: HexColor;
  fill?: HexColor;
  headerFill?: HexColor;
  borderColor?: HexColor;
  borderWidth?: number;
  cellPadding?: number;
};

export type SlideAstElement =
  | SlideAstText
  | SlideAstRect
  | SlideAstLine
  | SlideAstGroup
  | SlideAstList
  | SlideAstTable;

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
export type GroupInput = Omit<SlideAstGroup, 'kind'>;
export type ListInput = Omit<SlideAstList, 'kind'>;
export type TableInput = Omit<SlideAstTable, 'kind'>;

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

export function group(input: GroupInput): SlideAstGroup {
  return { kind: 'open-slide.group', ...input, children: input.children.map(cloneElement) };
}

export function list(input: ListInput): SlideAstList {
  return { kind: 'open-slide.list', ...input, items: [...input.items] };
}

export function table(input: TableInput): SlideAstTable {
  return { kind: 'open-slide.table', ...input, rows: input.rows.map((row) => [...row]) };
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
      elements: item.children.flatMap((element) => elementToPptx(element)),
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
  return { ...item, children: item.children.map(cloneElement) };
}

function cloneElement<T extends SlideAstElement>(element: T): T {
  if (element.kind === 'open-slide.group') {
    return { ...element, children: element.children.map(cloneElement) } as T;
  }
  if (element.kind === 'open-slide.list') {
    return { ...element, items: [...element.items] } as T;
  }
  if (element.kind === 'open-slide.table') {
    return { ...element, rows: element.rows.map((row) => [...row]) } as T;
  }
  return { ...element };
}

function elementToPptx(element: SlideAstElement, offset = { x: 0, y: 0 }): PptxElement[] {
  if (element.kind === 'open-slide.text') {
    const { kind: _kind, ...rest } = element;
    return [{ type: 'text', ...withOffset(rest, offset) }];
  }
  if (element.kind === 'open-slide.rect') {
    const { kind: _kind, ...rest } = element;
    return [{ type: 'rect', ...withOffset(rest, offset) }];
  }
  if (element.kind === 'open-slide.line') {
    const { kind: _kind, ...rest } = element;
    return [{ type: 'line', ...withOffset(rest, offset) }];
  }
  if (element.kind === 'open-slide.group') {
    return element.children.flatMap((child) =>
      elementToPptx(child, { x: offset.x + element.x, y: offset.y + element.y }),
    );
  }
  if (element.kind === 'open-slide.list') return listToPptx(element, offset);
  return tableToPptx(element, offset);
}

function withOffset<T extends { x: number; y: number }>(
  element: T,
  offset: { x: number; y: number },
): T {
  return { ...element, x: element.x + offset.x, y: element.y + offset.y };
}

function listToPptx(element: SlideAstList, offset: { x: number; y: number }): PptxElement[] {
  const marker = (index: number) => (element.ordered ? `${index + 1}.` : '•');
  return [
    {
      type: 'text',
      x: element.x + offset.x,
      y: element.y + offset.y,
      w: element.w,
      h: element.h,
      text: element.items.map((item, index) => `${marker(index)} ${item}`).join('\n'),
      fontSize: element.fontSize,
      fontFace: element.fontFace,
      color: element.color,
      bold: element.bold,
      fill: element.fill,
    },
  ];
}

function tableToPptx(element: SlideAstTable, offset: { x: number; y: number }): PptxElement[] {
  const rowCount = element.rows.length;
  const colCount = Math.max(1, ...element.rows.map((row) => row.length));
  const cellW = element.w / colCount;
  const cellH = element.h / Math.max(1, rowCount);
  const padding = element.cellPadding ?? 12;
  const items: PptxElement[] = [];

  element.rows.forEach((row, rowIndex) => {
    for (let colIndex = 0; colIndex < colCount; colIndex += 1) {
      const x = element.x + offset.x + colIndex * cellW;
      const y = element.y + offset.y + rowIndex * cellH;
      items.push({
        type: 'rect',
        x,
        y,
        w: cellW,
        h: cellH,
        fill: rowIndex === 0 ? (element.headerFill ?? element.fill) : element.fill,
        line: element.borderColor ?? '#CBD5E1',
        lineWidth: element.borderWidth ?? 1,
      });
      items.push({
        type: 'text',
        x: x + padding,
        y: y + padding,
        w: Math.max(0, cellW - padding * 2),
        h: Math.max(0, cellH - padding * 2),
        text: row[colIndex] ?? '',
        fontSize: element.fontSize,
        fontFace: element.fontFace,
        color: element.color,
        bold: rowIndex === 0,
        valign: 'mid',
      });
    }
  });

  return items;
}

function renderReactElement(element: SlideAstElement, index: number) {
  if (element.kind === 'open-slide.text') return renderText(element, index);
  if (element.kind === 'open-slide.rect') return renderRect(element, index);
  if (element.kind === 'open-slide.line') return renderLine(element, index);
  if (element.kind === 'open-slide.group') return renderGroup(element, index);
  if (element.kind === 'open-slide.list') return renderList(element, index);
  return renderTable(element, index);
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

function renderGroup(element: SlideAstGroup, index: number) {
  return (
    <div data-open-slide-ast="group" key={index} style={baseStyle(element)}>
      {element.children.map((child, childIndex) => renderReactElement(child, childIndex))}
    </div>
  );
}

function renderList(element: SlideAstList, index: number) {
  const Tag = element.ordered ? 'ol' : 'ul';
  return (
    <Tag
      data-open-slide-ast="list"
      key={index}
      style={{
        ...baseStyle(element),
        margin: 0,
        paddingLeft: element.ordered ? 34 : 28,
        display: 'flex',
        flexDirection: 'column',
        gap: element.gap ?? 10,
        background: element.fill,
        color: element.color ?? '#111827',
        fontFamily: element.fontFace,
        fontSize: element.fontSize,
        fontWeight: element.bold ? 700 : undefined,
      }}
    >
      {element.items.map((item) => (
        <li key={item}>{item}</li>
      ))}
    </Tag>
  );
}

function renderTable(element: SlideAstTable, index: number) {
  return (
    <table
      data-open-slide-ast="table"
      key={index}
      style={{
        ...baseStyle(element),
        borderCollapse: 'collapse',
        tableLayout: 'fixed',
        background: element.fill,
        color: element.color ?? '#111827',
        fontFamily: element.fontFace,
        fontSize: element.fontSize,
      }}
    >
      <tbody>
        {element.rows.map((row, rowIndex) => (
          <tr key={row.join('|')}>
            {row.map((cell) => (
              <td
                key={cell}
                style={{
                  border: `${element.borderWidth ?? 1}px solid ${element.borderColor ?? '#CBD5E1'}`,
                  background: rowIndex === 0 ? element.headerFill : undefined,
                  padding: element.cellPadding ?? 12,
                  fontWeight: rowIndex === 0 ? 700 : undefined,
                }}
              >
                {cell}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
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
