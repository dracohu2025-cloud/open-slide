export type { ImagePlaceholderProps } from './app/components/image-placeholder.tsx';
export { ImagePlaceholder } from './app/components/image-placeholder.tsx';
export type {
  DesignFonts,
  DesignPalette,
  DesignSystem,
  DesignTypeScale,
} from './app/lib/design.ts';
export { cssVarsToString, defaultDesign, designToCssVars } from './app/lib/design.ts';
export type { Page, SlideMeta, SlideModule } from './app/lib/sdk.ts';
export { CANVAS_HEIGHT, CANVAS_WIDTH } from './app/lib/sdk.ts';
export type { OpenSlideConfig } from './config.ts';
export type { Locale, Plural } from './locale/types.ts';
export type {
  PptxDeck,
  PptxElement,
  PptxLineElement,
  PptxRectElement,
  PptxSlide,
  PptxTextElement,
} from './pptx/exporter.ts';
export { createPptxBuffer, pxToEmu } from './pptx/exporter.ts';
export type {
  DefineDeckInput,
  GroupInput,
  LineInput,
  ListInput,
  RectInput,
  SlideAstDeck,
  SlideAstElement,
  SlideAstGroup,
  SlideAstLine,
  SlideAstList,
  SlideAstRect,
  SlideAstSlide,
  SlideAstTable,
  SlideAstText,
  SlideInput,
  TableInput,
  TextInput,
} from './slide-ast/index.tsx';
export {
  defineDeck,
  group,
  isSlideAstDeck,
  line,
  list,
  rect,
  renderPptx,
  renderReact,
  slide,
  table,
  text,
} from './slide-ast/index.tsx';
