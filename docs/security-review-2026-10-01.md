# Security and Zotero 10 compatibility review

Reviewed on 2026-10-01. This is a source review with targeted regression tests and dependency
auditing, not a penetration test or a guarantee that all issues were found.

## Findings

| Severity                      | Finding                                                          | Evidence and impact                                                                                                                                                                                                                                                                                                                                                                                   | Status / recommendation                                                                                                                                                                                    |
| ----------------------------- | ---------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| High                          | Markdown link scheme bypass                                      | `src/ui/markdown.ts` checked a scheme regex before URL parsing. `java\tscript:alert(1)` passes the relative-link branch, but the URL parser normalizes it to `javascript:`. Model output or malicious document instructions can supply such a link. Clicking it could execute script; execution in Zotero's chrome context was not tested. Relative links also resolve against a privileged document. | Fixed: parse first and accept only absolute HTTP, HTTPS, and mailto URLs. Regression tests cover tabs, data/chrome schemes, and relative URLs.                                                             |
| Medium                        | Plain-text provider credentials                                  | `src/preferences/provider-profile.ts:99-123` reads and writes raw OpenAI, Anthropic, and Gemini keys through Zotero prefs. The separate secret-reference abstraction is not used for these settings. Profile backups and local processes with profile access can recover keys.                                                                                                                        | Open: migrate to an OS credential store and clear legacy prefs after a successful migration. Correct claims that no raw keys are stored.                                                                   |
| Medium                        | Unauthenticated proxy shutdown, including opaque browser origins | `scripts/llm-proxy/server.mjs` explicitly exempts `POST /api/shutdown` from bearer auth, while `enforceTransportPolicy` allows `Origin: null`. A local process can stop the service. An opaque-origin sandboxed iframe can potentially send a simple POST to loopback without needing to read its response, subject to browser private-network restrictions.                                          | Open: authenticate shutdown, redesign orphan ownership, and reject opaque origins. The exemption currently supports orphan takeover, so removing it alone breaks that workflow.                            |
| Medium                        | Manual proxy start is unauthenticated                            | With `LLM_PROXY_AUTH_TOKEN` unset, bearer enforcement returns true for all routes. Local processes can submit prompts using the user's subscription. The null-origin allowance increases possible browser exposure. Automatically spawned plugin instances do supply a fresh token.                                                                                                                   | Open: fail closed or require an explicit development opt-in for no-auth mode.                                                                                                                              |
| Low                           | Fork still trusted upstream updates                              | The manifest and generated update links targeted `vishnutskumar` rather than this fork. A later upstream release could replace locally installed fork code. This is a trust/ownership mismatch, not evidence of malicious upstream behavior.                                                                                                                                                          | Fixed: update URLs now target `GitPancaked/zotero-ai-explain`. The extension ID is retained for in-place upgrades. Fork automatic updates require publishing releases there.                               |
| Development tooling: critical | Vulnerable locked development dependencies                       | `npm audit` reports 10 affected packages: 2 critical, 6 high, 1 moderate, 1 low. Includes Vitest UI server arbitrary file read/execution and transitive Vite, undici, and brace-expansion findings. Counts include dependent packages, not 10 distinct exploit paths. `npm audit --omit=dev` reports zero.                                                                                            | Open: update development dependencies in a separate tested change. These packages are not shipped as runtime dependencies in the XPI. The configured `vitest run` does not start the vulnerable UI server. |

## Additional boundaries

### Follow-up: secure credentials in v0.4.2

The plain-text credential finding above is addressed in v0.4.2: direct API keys now use Windows
Credential Manager or Linux Secret Service, including Arch. Legacy keys migrate at startup and are
removed from preferences only after verified storage. A preexisting secure credential takes
precedence over a stale legacy preference. Failed migrations retain legacy preferences without using
them; failed new saves never fall back to plaintext. Old backups may still contain credentials.

Windows native write/read/delete was tested with a unique synthetic credential. Linux behavior has
adversarial mocked tests, but a real Arch keyring smoke test is still required. Keys are shared
across Zotero profiles of the same OS user. A locked or unavailable store can delay the first
startup lookup up to 30 seconds. Saving multiple keys attempts rollback on failure; if rollback or
preferences cleanup fails, settings display an explicit recovery message. OS stores protect keys at
rest, not against other applications running as the same user.

See [the credential smoke-test checklist](manual-verification/secure-credentials.md).

Follow-up validation: build/typecheck, repository-wide ESLint, and changed-file format checks
passed. Full suite: 96 files passed, 3 skipped; 1,218 tests passed, 6 skipped. The v0.4.2 archive
was checked for both OS integrations, secure-profile wiring, required bootstrap globals, ZIP
integrity, and its SHA-256 update hash.

- The proxy binds to IPv4 loopback, checks Host, authenticates normal plugin requests with per-spawn
  tokens, compares tokens with `timingSafeEqual`, and caps request bodies.
- Raw model HTML is rendered using text nodes. The settings `innerHTML` assignment clears a
  container using a constant empty string; it does not interpolate input.
- CLI calls use argument arrays rather than a shell. Codex runs with a read-only sandbox and
  isolated HOME/CODEX_HOME; Claude uses an isolated working directory, MCP restrictions, and an
  empty allowed-tools argument. These are not proof against prompt injection or all CLI tool use.
  Read-only sandboxing alone is not a confidentiality boundary. Model text should be treated as
  untrusted.
- Codex copies its OAuth auth file into a temporary directory. POSIX mkdtemp normally restricts
  directory access, but abrupt termination can leave credential copies. Verify Windows ACLs and
  cleanup behavior before relying on this as secure storage.
- Chat and embedding providers receive selected text or retrieved library chunks by design. The
  index stores extracted document text on disk. Cloud-provider choices therefore affect document
  privacy; plaintext indexes need protected profile access.

## Compatibility and packaging

The original manifest declared Zotero 8.0 through 9.99.99, excluding Zotero 10. The
[official migration guide](https://www.zotero.org/support/dev/zotero_10_for_developers) specifies
`10.0.*`. This branch updates that limit and bumps the plugin to 0.4.1.

Source searches found no uses of the removed singular collection-selection getters,
`itemsView.collectionTreeRow`, `fulltextWord`, or CookieSandbox. The full-text cache fallback was
checked against
[Zotero 10 source](https://github.com/zotero/zotero/blob/10.0/chrome/content/zotero/xpcom/fulltext.js);
PDF extraction uses PDFWorker rather than the removed full-text word tables. No Zotero profile was
modified or plugin installed during this review.

Windows packaging now falls back to the built-in tar ZIP writer when 7-Zip is absent. Only
manifest.json, bootstrap.js, content, and the dependency-free proxy are packaged. Install the
generated XPI using Zotero's Plugins manager, **Install Plugin From File**. Smoke-test reader
explanations, library indexing, citation jumps, settings persistence, and proxy startup/shutdown in
Zotero 10 before publishing this as a tested release.

## Validation results

- Build, TypeScript type checking, and ESLint passed using Node 22.22.3.
- Full Vitest suite: 90 files passed, 3 skipped; 1,175 tests passed, 6 skipped.
- XPI ZIP integrity, root entries, bundled manifest version/range, proxy inclusion, updated markdown
  code, fork update link, and SHA-256 update hash were verified.
- `npm run verify` stops at repository-wide formatting: 229 existing files report issues in this
  Windows checkout. An unchanged source sample has LF in Git and CRLF locally, explaining
  checkout-related warnings. Changed source files were formatted separately; no broad formatting
  rewrite was included.
- Real Zotero end-to-end tests were not run. The compatibility declaration follows source/API review
  and automated contracts, and still needs the smoke test above.
