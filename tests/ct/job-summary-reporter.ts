// Playwright reporter that appends a GitHub Actions job summary for the
// `component-tests` CI job, in the same shape as the "Vitest Test Report"
// Vitest's built-in `github-actions` reporter writes for the `build` job.
// Playwright ships nothing equivalent (its `github` reporter only emits
// annotations). A no-op outside GitHub Actions, where GITHUB_STEP_SUMMARY
// is unset.
import { appendFileSync } from 'fs';
import { relative } from 'path';
import type {
  FullConfig,
  FullResult,
  Reporter,
  Suite,
  TestCase,
} from 'playwright/types/testReporter';

const TITLE = 'Playwright Component Test Report';
const SEPARATOR = ' · ';

function noun(count: number, singular: string, plural: string): string {
  return count === 1 ? singular : plural;
}

function repoPath(file: string): string {
  return relative(process.cwd(), file).replace(/\\/g, '/');
}

// Links to the file at the exact commit CI ran, when the run context allows.
function testLink(test: TestCase): string {
  const { file, line } = test.location;
  const label = `\`${test.titlePath().filter(Boolean).slice(1).join(' › ')}\``;
  const { GITHUB_SERVER_URL, GITHUB_REPOSITORY, GITHUB_SHA } = process.env;
  if (!GITHUB_REPOSITORY || !GITHUB_SHA) return label;
  const server = GITHUB_SERVER_URL ?? 'https://github.com';
  return `[${label}](${server}/${GITHUB_REPOSITORY}/blob/${GITHUB_SHA}/${repoPath(file)}#L${line})`;
}

export default class JobSummaryReporter implements Reporter {
  private rootSuite: Suite | undefined;

  onBegin(_config: FullConfig, suite: Suite): void {
    this.rootSuite = suite;
  }

  onEnd(result: FullResult): void {
    const outputPath = process.env.GITHUB_STEP_SUMMARY;
    if (!outputPath || !this.rootSuite) return;

    const tests = this.rootSuite.allTests();
    const failed = tests.filter((t) => t.outcome() === 'unexpected');
    const flaky = tests.filter((t) => t.outcome() === 'flaky');
    const skipped = tests.filter((t) => t.outcome() === 'skipped');
    const expectedFail = tests.filter(
      (t) => t.outcome() === 'expected' && t.expectedStatus === 'failed'
    );
    // Vitest counts a test that passed on retry as a pass and lists it again
    // under "Flaky Tests"; do the same.
    const passed = tests.filter(
      (t) =>
        (t.outcome() === 'expected' && t.expectedStatus !== 'failed') || t.outcome() === 'flaky'
    );

    const files = new Map<string, boolean>(); // file -> any test failed
    for (const t of tests) {
      if (t.outcome() === 'skipped') continue;
      const file = t.location.file;
      files.set(file, (files.get(file) ?? false) || t.outcome() === 'unexpected');
    }
    const failedFiles = [...files.values()].filter(Boolean).length;
    const passedFiles = files.size - failedFiles;

    const fileInfo: string[] = [];
    if (failedFiles > 0)
      fileInfo.push(`❌ **${failedFiles} ${noun(failedFiles, 'failure', 'failures')}**`);
    if (passedFiles > 0)
      fileInfo.push(`✅ **${passedFiles} ${noun(passedFiles, 'pass', 'passes')}**`);
    fileInfo.push(`${files.size} total`);

    const testInfo: string[] = [];
    if (failed.length > 0)
      testInfo.push(`❌ **${failed.length} ${noun(failed.length, 'failure', 'failures')}**`);
    if (passed.length > 0)
      testInfo.push(`✅ **${passed.length} ${noun(passed.length, 'pass', 'passes')}**`);
    if (expectedFail.length > 0) {
      testInfo.push(
        `🔵 **${expectedFail.length} expected ${noun(expectedFail.length, 'failure', 'failures')}**`
      );
    }
    testInfo.push(`${failed.length + passed.length + expectedFail.length} total`);

    let summary = `## ${TITLE}\n\n### Summary\n\n`;
    summary += `- **Test Files**: ${fileInfo.join(SEPARATOR)}\n`;
    summary += `- **Test Results**: ${testInfo.join(SEPARATOR)}\n`;
    if (skipped.length > 0) {
      summary += `- **Other**: ${skipped.length} ${noun(skipped.length, 'skip', 'skips')}${SEPARATOR}${skipped.length} total\n`;
    }
    // A build error (e.g. Vite failing to bundle the components) or a global
    // timeout can fail the run with no individual test failing.
    if (result.status !== 'passed' && failed.length === 0) {
      summary += `- **Run status**: ❌ **${result.status}** (see the job log)\n`;
    }

    if (failed.length > 0) {
      summary += '\n### Failed Tests\n\n';
      for (const t of failed) summary += `- ${testLink(t)}\n`;
    }
    if (flaky.length > 0) {
      summary +=
        '\n### Flaky Tests\n\nThese tests passed only after one or more retries, indicating potential instability.\n\n';
      for (const t of flaky) {
        const retries = t.results.length - 1;
        summary += `- ${testLink(t)} (passed on retry ${retries} out of ${t.retries})\n`;
      }
    }

    try {
      appendFileSync(outputPath, summary);
    } catch (error) {
      console.warn(`Could not write the job summary to ${outputPath}:`, error);
    }
  }

  printsToStdio(): boolean {
    return false;
  }
}
