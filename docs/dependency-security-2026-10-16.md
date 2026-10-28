# Dependency advisory review — 2026-10-16

Scope: the lockfile at main `fc8db657d520ccbbf538a91ae06d024ae847334a`, independently of the diagram work in PR #3. No application code, test expectations, input semantics, or renderer changed. This is an advisory/reachability review, not a penetration test or proof of absence of vulnerabilities.

## Result and bounded update

- Before: a fresh read-only `npm audit --json` reported **10 affected packages** (2 critical, 5 high, 3 moderate); `npm audit --omit=dev --json` reported **2 high** production packages. These include parent/metavulnerability entries, not 10 distinct reachable attacks. The earlier CI installation log reported 8; the two observations are retained without assuming why they differed.
- After: fresh `npm ci`, then both audit commands reported **0 known vulnerabilities** on 2026-10-16. Advisory databases can change; this is a point-in-time result.
- Vitest: `2.1.9` → **4.1.11**, the minimum release addressing the reviewed mocker advisory as well as the older UI advisory. Version 3.2.6 alone would leave the mocker advisory. No test/config API changes were required: all **651 existing tests in 27 files** pass unchanged, with typecheck and build.
- Vite: `5.4.21` → **6.4.3**, constrained under the exact Vitest version. This avoids resolving a newer major with a higher Node minimum. Tinypool and vite-node are removed by Vitest 4; nested esbuild is `0.25.12`. PostCSS resolves to `8.5.30` with patched `source-map-js@1.2.2`.
- PptxGenJS stays **3.12.0**. Its manifest is pinned exactly so that the narrowly scoped child override cannot silently apply to a future renderer. Only its `image-size` dependency changes from `1.2.1` to **2.0.4** (fix floor: 2.0.3). Even PptxGenJS 4.0.1 still declares `image-size ^1.2.1`; a renderer major upgrade would not solve this alert.
- No `npm audit fix --force`, new service, API key, runtime network call, or audit suppression was used. The lockfile also contains the necessary test-tool transitive changes; unrelated direct dependencies remain unchanged.

## Installed paths, prerequisites, and observed reachability

| Lockfile/dependency path before | Class | Official advisory / fix floor | What is established in this repository |
| --- | --- | --- | --- |
| `pptxgenjs@3.12.0 → image-size@1.2.1` (`node_modules/image-size`) | Production, high | [JXL/HEIF DoS](https://github.com/advisories/GHSA-5p2g-fcmc-qvqq), [ICNS DoS](https://github.com/advisories/GHSA-w3rx-r6r6-pgpr); 2.0.3 | Malformed image parsing can loop indefinitely. The vulnerable package is installed, but no import or call occurs in any of PptxGenJS 3.12's four distributed JS files. `src/render/surface.ts`, `src/templates/master.ts`, and `src/templates/loader.ts` use preloaded image data URIs and explicit geometry, not image-size. No reachable parser call was found. Extension filtering alone is not treated as a mitigation. |
| `vitest@2.1.9` | Development, critical | [UI/API arbitrary file access/execution](https://github.com/advisories/GHSA-5xrq-8626-4rwp); 3.2.6 / 4.1.0 | Repository runs Node-only tests, with no UI/API exposure or browser configuration. `api` and `ui` are false by default; UI/browser packages are absent. The advisory's exposed API/UI or Windows UI/browser prerequisites are not configured here. |
| `vitest → @vitest/mocker@2.1.9` | Development, moderate | [Public mocker/interceptor HMR access](https://github.com/advisories/GHSA-82fw-gwwq-j7x9); 4.1.11 | No public mocker/interceptor plugin or reachable HMR server is configured. The affected dependency exists; the advisory's unauthenticated path was not established. The advisory lists no patched 2.x/3.x release as of this review. |
| `vitest → tinypool@1.1.1` | Development, critical | [Worker option gadget](https://github.com/advisories/GHSA-5gmw-xhrv-c9v3); 2.1.1. [Run option gadget](https://github.com/advisories/GHSA-85c8-ppgw-ccpr); 2.1.2 | Vitest actually uses Tinypool. These gadgets require an upstream prototype-pollution primitive. No attacker-controlled pollution source reaching the test parent process was established. The worker/run path is relevant; installed severity is not evidence of demonstrated remote execution. Vitest 4 removes the package. |
| `vitest → vite@5.4.21` (also via `vite-node`) | Development, high | [Source-map traversal](https://github.com/advisories/GHSA-4w7w-66w2-5vf9), [Windows file bypass](https://github.com/advisories/GHSA-fx2h-pf6j-xcff), [Windows editor/NTLM path](https://github.com/advisories/GHSA-v6wh-96g9-6wx3); all addressed by 6.4.3 | Require a serving development server, with additional platform/filesystem or NTLM conditions for the Windows issues. No exposed server configuration is present; CI runs on Ubuntu. This is not a claim about developers who independently enable a server. |
| `vite → esbuild@0.21.5` (`node_modules/vite/node_modules/esbuild`) | Development, moderate | [Serve cross-origin data access](https://github.com/advisories/GHSA-67mh-4wv8-2f99); 0.25.0 | No use of esbuild's serve feature found. The separate esbuild used by tsx was already 0.28.2. |
| `vite → postcss@8.5.26 → source-map-js@1.2.1` | Development, high | [Indexed source-map DoS](https://github.com/advisories/GHSA-68fv-2mgg-jv7q); 1.2.2 | Requires processing an attacker-controlled indexed source map. No presentation-input-to-source-map path found. Compatible patched dependency is available. |

The production audit's second entry is the PptxGenJS parent/metavulnerability; the parser findings are not two independent reachable PptxGenJS call paths. No exploit payloads were executed and no actual compromise was observed or inferred.

## Compatibility boundaries

The [Vitest 4 migration guide](https://v4.vitest.dev/guide/migration) describes breaking changes to pools, mocks, coverage, and advanced APIs. This repository uses basic Node test APIs, a manually created fetch spy, and only `include`, `environment`, and `testTimeout` configuration. Its existing suite passed without source, expectation, or configuration changes. Development uses the Node versions supported by Vitest 4.1.11 (`^20`, `^22`, or `>=24`); CI exercises 20 and 24. The application's existing runtime engine declaration is unchanged; that declaration is not a promise that the development toolchain runs on odd Node 21/23.

[image-size 2.0](https://github.com/image-size/image-size/releases/tag/v2.0.0) changes CommonJS usage to named `imageSize` and removes synchronous file-path sizing. The override is justified only for the inspected **PptxGenJS 3.12.0**, which does not call the library. Re-evaluate the override and image paths before any PptxGenJS upgrade, rather than widening its scope. Existing image/template/import tests remain part of the unchanged test suite.

## Reproduction and visual regression

```sh
npm ci
npm ls vitest vite pptxgenjs image-size source-map-js tinypool vite-node --all
npm audit --json
npm audit --omit=dev --json
npm run check
npm run example:relationships
npm run example:research
npm run example:explanatory
npm run example:decomposition
npm run preview -- examples/relationships/generated/six-relationships.pptx .preview/relationships
npm run preview -- examples/research/generated/nec-disaster-vlm-2026-10-14.pptx .preview/research
```

Preview requires LibreOffice, Poppler, and the portable theme's Noto Sans CJK JP fonts (see [research workflow](../examples/research/README.md)); generation and tests need no external AI calls. Compare the same inputs, fonts, and render software before/after. PDF binary hashes and complete PPTX hashes are not stable comparisons because metadata contains generation times. Compare page pixels and unzipped slide content instead. Native PowerPoint rendering on user operating systems is not asserted.

Local visual regression: all 10 pages (six relationships + four NEC research slides, 1600 × 901 pixels each) are pixel-identical before/after; both PPTX archives have identical entry contents except generation timestamps in `docProps/core.xml`. All four CI example commands regenerate successfully (19 pages total), with no strict layout issues. The unchanged suite includes embedded-image/template/import tests. See the separate dependency PR for CI results. The update is independent of PR #3 and does not incorporate its new diagram or argument-review behavior.
