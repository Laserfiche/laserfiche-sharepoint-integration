<!--Copyright (c) Laserfiche.
Licensed under the MIT License. See LICENSE.md in the project root for license information.-->

# Contributing to laserfiche-sharepoint-integration

We encourage and appreciate feedback and contributions from the community!

- [Code of Conduct](#coc)
- [Questions](#question)
- [Issues and Feature Requests](#issue)
- [Making a Change](#change)
- [Coding Guidelines](#rules)
- [Submitting a Pull Request](#submit-pr)

## <a name="coc"></a> Code of Conduct

Please read and follow our [Code of Conduct](./code_of_conduct.md).

## <a name="question"></a> Questions

Post general questions on [Laserfiche Answers](https://answers.laserfiche.com/), and mention _laserfiche-sharepoint-integration_.

## <a name="issue"></a> Issues and Feature Requests

Search [GitHub Issues](https://github.com/Laserfiche/laserfiche-sharepoint-integration/issues) for an existing bug report or feature request first. If yours isn't there, open a new one.

## <a name="change"></a> Making a Change

1. Branch off **`1.x`**, the active release branch. `main` is no longer used.
2. Set up the project as described in the [README](./README.md#getting-started). Use Node.js 22 and install with `npm ci`.
3. Make your change test-first (see below), then check that it builds and passes locally:

   ```bash
   npm test          # unit tests
   npm run test:ct   # component tests (type-checked first)
   npm run bundle    # release build (fails on any lint warning)
   ```

[`AGENTS.md`](./AGENTS.md) covers the project's conventions in more depth: pinned versions, SharePoint's script-loading restrictions, and the test setup.

## <a name="rules"></a> Coding Guidelines

- **Don't repeat yourself.** Before adding a component, style, string, icon or helper, look for an existing one and reuse it. If similar code exists in more than one place, extract it to a shared place and use it everywhere. [`AGENTS.md`](./AGENTS.md#1-reuse-before-you-add-dry) lists where the shared pieces live.
- **Write tests first.** Add a test that fails, make it pass, then refactor. Every bug fix starts with a test that reproduces the bug. Test what users see, not implementation details.
- **Keep tests current.** When behavior changes, update its tests in the same change. Don't skip or delete a failing test to get a change through.
- **Keep it clean.** Use small functions and components with clear, informative names. Explain _why_ in comments, not _what_. Leave the code better than you found it.
- **Follow the formatters.** Prettier and ESLint (the SPFx config) define the code style. Run `npm run format` before you commit.

## <a name="submit-pr"></a> Submitting a Pull Request

- Search the existing [Pull Requests](https://github.com/Laserfiche/laserfiche-sharepoint-integration/pulls) to make sure yours isn't a duplicate.
- Target **`1.x`**, and give the PR a short, imperative title that names the affected component.
- Link the issue the PR addresses.
- Include tests for the change. If part of it can only be verified in SharePoint, describe the manual steps you ran.
- Update the documentation (`jekyll_files/`, the README or `AGENTS.md`) if the change affects it.
- CI runs a formatting check, the release build, unit tests and component tests. All must pass before a PR can be merged.

After you submit, the project's core members will review the code.
