import { test as base, expect } from '@playwright/test';

export const test = base.extend<{ isolatedCatalog: void }>({
  isolatedCatalog: [async ({ request }, use) => {
    const response = await request.post('/__test__/reset', {
      headers: { 'X-Test-Token': process.env.VGHTPE_UI_TEST_TOKEN || '' },
    });
    expect(response.ok()).toBeTruthy();
    await use();
  }, { auto: true }],
});
export { expect };
