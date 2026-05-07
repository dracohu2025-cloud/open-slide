# @open-slide/core

Runtime and CLI for [open-slide](https://github.com/1weiho/open-slide) — a React-based slide framework where you write slides and the framework handles the Vite/React stack, layout, navigation, hot reload, and fullscreen play mode.

## Install

```bash
pnpm add @open-slide/core
```

Most users get this installed automatically by running `npx @open-slide/cli init`. Use this package directly only if you're wiring up an existing workspace by hand.

## What's inside

- **Runtime** — home page, slide viewer, thumbnail rail, keyboard navigation, and fullscreen presenter mode. Every slide renders into a fixed **1920×1080** canvas; the framework scales it.
- **Vite plugin** — discovers `slides/<id>/index.{tsx,jsx,ts,js}`, exposes them via virtual modules, and reloads when slides are added or removed.
- **CLI** — `open-slide dev | build | preview` so workspaces never need to touch Vite, React, or tsconfig directly.

## CLI

Once installed, the `open-slide` bin is available in the workspace:

| Command | Description |
| --- | --- |
| `open-slide dev` | Start the dev server. Flags: `-p, --port <port>`, `--host [host]`, `--open`. |
| `open-slide build` | Build a static site. Flags: `--out-dir <dir>` (defaults to `dist`). |
| `open-slide preview` | Preview the production build. Flags: `-p, --port <port>`, `--host [host]`, `--open`. |
| `open-slide export:pptx <slide-id>` | Export a schema-based editable PPTX file. Flags: `-o, --output <file>`. |

## Config

Create `open-slide.config.ts` in the workspace root (all fields optional):

```ts
import type { OpenSlideConfig } from '@open-slide/core';

const openSlideConfig: OpenSlideConfig = {
  slidesDir: 'slides',
  port: 5173,
};

export default openSlideConfig;
```

## Authoring slides

Slides live under `slides/<kebab-case-id>/index.tsx` and default-export an array of `Page` components:

```tsx
import type { Page } from '@open-slide/core';

const Cover: Page = () => (
  <div className="flex h-full w-full items-center justify-center">
    <h1 className="text-[120px] font-bold">Hello, open-slide</h1>
  </div>
);

const pages: Page[] = [Cover];
export default pages;
export const meta = { title: 'Hello' };
```

## Editable PPTX export

`open-slide export:pptx <slide-id>` creates a real `.pptx` with editable PowerPoint shapes and text. It does **not** screenshot the React DOM.

Preferred path: define a shared Slide AST once, then render it both ways:

```tsx
import { defineDeck, group, list, rect, renderReact, slide, table, text } from '@open-slide/core';

export const ast = defineDeck({
  title: 'Hello',
  slides: [
    slide({
      background: '#0F172A',
      children: [
        group({
          x: 96,
          y: 96,
          w: 900,
          h: 260,
          children: [
            rect({ x: 0, y: 0, w: 900, h: 260, fill: '#22C55E', radius: 24 }),
            text({ x: 32, y: 32, w: 820, h: 70, text: 'Hello editable PPTX', fontSize: 44, bold: true, color: '#FFFFFF' }),
            list({ x: 40, y: 122, w: 760, h: 90, items: ['React pages', 'Editable PowerPoint shapes'], fontSize: 28, color: '#FFFFFF' }),
          ],
        }),
        table({
          x: 96,
          y: 420,
          w: 760,
          h: 180,
          rows: [
            ['Renderer', 'Output'],
            ['React', 'Page[]'],
            ['PPTX', 'Shapes'],
          ],
          headerFill: '#E0F2FE',
        }),
      ],
    }),
  ],
});

export default renderReact(ast);
```

Then run:

```bash
open-slide export:pptx hello -o hello.pptx
```

The CLI reads `ast` automatically and renders it to editable PPTX. For compatibility, it still accepts the lower-level `pptx: PptxDeck` export.

Coordinates use the same fixed 1920×1080 canvas as Open Slide and map to PowerPoint widescreen EMUs.

## Exports

```ts
import {
  CANVAS_WIDTH,   // 1920
  CANVAS_HEIGHT,  // 1080
  type Page,
  type SlideMeta,
  type SlideModule,
  type OpenSlideConfig,
} from '@open-slide/core';
```

The Vite plugin is exposed under a subpath for advanced setups:

```ts
import { createViteConfig } from '@open-slide/core/vite';
```

## License

MIT
