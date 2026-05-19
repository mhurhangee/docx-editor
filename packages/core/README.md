# @eigenpal/docx-editor-core

Framework-agnostic core for the [docx-editor](https://docx-editor.dev). Parses DOCX, builds the document model, runs ProseMirror, and renders Word-fidelity pages. Powers the React and Vue adapters and anything else you build on top.

## Quick Start

Most users want the [React](https://www.npmjs.com/package/@eigenpal/docx-editor-react) or [Vue](https://www.npmjs.com/package/@eigenpal/docx-editor-vue) adapter. Reach for core directly when building a custom adapter, running headless on the server, or driving DOCX parsing/serialization without a UI.

```bash
npm install @eigenpal/docx-editor-core
```

```ts
import { readFile } from 'node:fs/promises';
import { parseDocx } from '@eigenpal/docx-editor-core/docx';

const buffer = await readFile('contract.docx');
const document = await parseDocx(buffer);
console.log(document.paragraphs.length);
```

Each subpath tree-shakes independently. Pick the smallest entry point that gives you what you need.

## Subpath map

| Building...                       | Import from                                                      | What you get                                                                       |
| --------------------------------- | ---------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| A new framework adapter           | `./docx`, `./prosemirror/conversion`, `./prosemirror/extensions` | `parseDocx`, `toProseDoc` / `fromProseDoc`, `createStarterKit`, `ExtensionManager` |
| Custom layout / rendering         | `./layout-engine`, `./layout-bridge`, `./layout-painter`         | `layoutDocument`, `mouseToPosition`, `renderPage`, `LayoutPainter`                 |
| Editor commands and plugins       | `./prosemirror/commands`, `./prosemirror/plugins`                | Formatting, tables, suggestion mode, selection tracker                             |
| Saving back to `.docx`            | `./docx`                                                         | `repackDocx`, `attemptSelectiveSave`                                               |
| Headless agents (no UI)           | `./agent`                                                        | `DocumentAgent`, `executeCommand`, `AgentCommand` types                            |
| An MCP server                     | `./mcp`                                                          | Model Context Protocol server scaffolding                                          |
| Just unit/color/clipboard helpers | `./utils`                                                        | `twipsToPixels`, `resolveColor`, font loading, clipboard, selection helpers        |
| Just a type                       | `./types/document`, `./types/content`, `./types/agentApi`        | `Document`, `Paragraph`, `Comment`, `AgentCommand`, ...                            |
| Default editor stylesheet         | `./prosemirror/editor.css`                                       | Import once at the top of your app                                                 |

## Stability

`./layout-engine`, `./layout-painter`, `./layout-bridge`, and `./plugin-api` are **`@experimental`** — used by the first-party adapters but the API may change in minor releases until a third-party adapter validates it. Pin a version range. Everything else follows SemVer.

## Peer dependencies

ProseMirror packages are declared as `peerDependencies` so consumer bundles don't ship duplicates:

```bash
npm i prosemirror-commands prosemirror-dropcursor prosemirror-history \
      prosemirror-keymap prosemirror-model prosemirror-state \
      prosemirror-tables prosemirror-transform prosemirror-view
```

## Architecture

Dual-rendering: a hidden ProseMirror instance owns editing state (selection, undo/redo, commands) while `layout-painter` produces the visible pages. Full breakdown: [CLAUDE.md](https://github.com/eigenpal/docx-editor/blob/main/CLAUDE.md#editor-architecture--dual-rendering-system).
