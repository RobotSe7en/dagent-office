import type { Page } from '@playwright/test';

/** SDK probes own their host DOM; the workspace has separate integration tests. */
export async function openSdkHost(page: Page): Promise<void> {
  await page.route('**/sdk-test-host.html', route => route.fulfill({
    contentType: 'text/html',
    body: '<!doctype html><html><head><meta charset="utf-8"><title>SDK test host</title></head><body></body></html>',
  }));
  await page.goto('/sdk-test-host.html');
}
