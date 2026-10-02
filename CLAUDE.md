# CLAUDE.md

This repo's agent instructions live in [`AGENTS.md`](./AGENTS.md). Read it before making any change. The essentials:

1. **Reuse before you add (DRY).** Search for an existing component, style, string, icon or helper and reuse it. If one is close but not shared, extract it and move every copy onto it in the same change. AGENTS.md lists where the shared pieces live.
2. **Clean code.** Small, well-named units; refactor what you touch; match the surrounding style.
3. **Test-driven development.** Red → green → refactor, testing user-visible behavior. Keep tests current with the code; never skip or delete one to get a change through.
4. **Done means** `npm run format:check`, `npm test`, `npm run test:ct` and `npm run bundle` (the SHIP build CI runs) all pass.
