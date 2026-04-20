# Repository Guidelines

## Project Structure & Module Organization
- `src/`: React + TypeScript renderer (components, contexts, utilities, editor extensions).
- `electron/`: Electron main/preload process and desktop integrations (filesystem, MCP, packaging hooks).
- `public/`: static assets served by Vite.
- `docs/`: development notes and troubleshooting docs.
- `website/`: project website/demo assets.
- Build outputs: `dist/` (web bundle) and `release/` (installers). Treat both as generated artifacts.

## Build, Test, and Development Commands
- `npm install`: install root dependencies.
- `npm run dev:electron`: recommended local workflow; starts Vite on port `3000` and launches Electron.
- `npm run dev`: renderer-only development in browser.
- `npm run build`: create production bundle in `dist/`.
- `npm run build:electron`: build and package desktop app to `release/`.
- `npm run preview`: preview the built renderer locally.
- Optional (vendored MCP server): `cd brave-search-mcp-server && npm run format:check && npm run build`.

## Coding Style & Naming Conventions
- Use TypeScript in `src/` (strict mode is enabled) and CommonJS (`.cjs`) in `electron/`.
- Match existing formatting: 2-space indentation, single quotes, and no semicolons.
- Naming: `PascalCase` for React component files (for example, `WordEditor.tsx`), `camelCase` for utilities (for example, `docxParser.ts`), and `UPPER_SNAKE_CASE` for constants.
- Keep UI/state logic in `src/components` and `src/context`; keep OS/file/network integration logic in `electron/`.

## Testing Guidelines
- No automated root test suite is configured yet.
- Minimum validation for each change: run `npm run build` and smoke-test with `npm run dev:electron`.
- Document manual verification steps in each PR (input file, action, expected result).
- If adding tests, colocate as `*.test.ts` or `*.test.tsx` near the feature and include run instructions in the PR.

## Commit & Pull Request Guidelines
- Follow Conventional Commit style used in history: `feat:`, `fix:`, `chore:`.
- Keep commit subjects imperative and scoped (example: `fix: handle missing BRAVE_API_KEY in search init`).
- PRs should include a concise summary, linked issue (if available), verification steps, and screenshots/GIFs for UI changes.
- Never commit secrets. Copy `.env.example` to `.env` locally and keep credentials out of version control.
