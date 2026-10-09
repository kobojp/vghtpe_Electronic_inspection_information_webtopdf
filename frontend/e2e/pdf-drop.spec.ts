import { test, expect } from './fixtures';
import { mkdir } from 'node:fs/promises';
import { resolve } from 'node:path';

test('PDF drop: highlight, native path delivery, filtering, duplicate selection and extraction', async ({ page, request }) => {
  const failures: string[] = [];
  page.on('pageerror', error => failures.push(error.message));
  await page.goto('/');
  await expect(page.locator('[data-app-ready]')).toBeVisible();
  const response = await request.get('/__test__/drop-files', { headers: { 'X-Test-Token': process.env.VGHTPE_UI_TEST_TOKEN || '' } });
  expect(response.ok()).toBeTruthy();
  const payload = await response.json();
  expect(payload.paths).toHaveLength(2);
  expect(payload.rejected).toEqual(['note.txt', 'broken.pdf']);
  const deliver = () => page.evaluate(result => window.dispatchEvent(new CustomEvent('vghtpe:pdf-drop', { detail: result })), payload);

  // A late desktop callback must not add files while this page is hidden.
  await deliver();
  await page.getByRole('button', { name: /PDF 擷取.*搜尋需要的頁面/ }).click();
  const zone = page.locator('#pdf-drop-zone');
  await expect(page.getByText('將 PDF 拖曳到這裡', { exact: true })).toBeVisible();
  const transfer = await page.evaluateHandle(() => {
    const data = new DataTransfer();
    data.items.add(new File(['%PDF'], 'sample.pdf', { type: 'application/pdf' }));
    return data;
  });
  await zone.dispatchEvent('dragenter', { dataTransfer: transfer });
  await expect(zone).toHaveClass(/drag-active/);
  await expect(page.getByText('放開滑鼠即可加入 PDF')).toBeVisible();
  await mkdir(resolve('../.cache/screenshots'), { recursive: true });
  await page.screenshot({ path: resolve('../.cache/screenshots/pdf-drop.png') });
  await zone.dispatchEvent('dragleave', { dataTransfer: transfer });
  await expect(zone).not.toHaveClass(/drag-active/);
  // Ordinary browser drops cannot expose OS paths; retain the select button fallback.
  await zone.dispatchEvent('drop', { dataTransfer: transfer });
  await expect(page.getByRole('status')).toContainText('請在桌面程式使用拖曳');
  await expect(page.getByRole('button', { name: '選擇 PDF', exact: true })).toBeEnabled();

  // Desktop DOMEventHandler supplies validated absolute paths through this event.
  await zone.evaluate(element => { (element as HTMLElement).dataset.nativeDropReady = 'true'; });
  await zone.dispatchEvent('dragenter', { dataTransfer: transfer });
  await zone.dispatchEvent('drop', { dataTransfer: transfer });
  await deliver();
  await expect(page.locator('.selected-files .file-row')).toHaveCount(2);
  await expect(page.getByRole('status')).toContainText('note.txt、broken.pdf');
  await expect(zone).not.toHaveClass(/drag-active/);
  await deliver();
  await page.getByRole('button', { name: '選擇 PDF', exact: true }).click();
  await expect(page.locator('.selected-files .file-row')).toHaveCount(2);

  // File drops anywhere must not navigate away from the application.
  const prevented = await page.evaluate(() => {
    const data = new DataTransfer();
    data.items.add(new File(['%PDF'], 'outside.pdf', { type: 'application/pdf' }));
    const event = new DragEvent('drop', { dataTransfer: data, bubbles: true, cancelable: true });
    document.body.dispatchEvent(event);
    return event.defaultPrevented;
  });
  expect(prevented).toBeTruthy();
  await page.getByLabel('搜尋關鍵字').fill('TARGET');
  await page.getByRole('button', { name: '開始擷取' }).click();
  await expect(page.getByText('已完成', { exact: true })).toBeVisible({ timeout: 10_000 });
  await expect(page.locator('.task-counts')).toContainText('成功 2');

  // Starting another task prevents late drops and edits to the input list.
  await page.getByRole('button', { name: /報表下載.*瀏覽與下載報表/ }).click();
  await page.getByLabel('名稱搜尋').fill('二門診滅火器');
  await page.getByRole('checkbox', { name: '選取目前清單的全部報表' }).check();
  await page.getByLabel('起始月份').fill('2026-07');
  await page.getByLabel('結束月份').fill('2026-12');
  await page.getByRole('button', { name: '開始下載' }).click();
  await page.getByRole('button', { name: /PDF 擷取.*搜尋需要的頁面/ }).click();
  await expect(zone).toHaveAttribute('aria-disabled', 'true');
  await page.evaluate(() => window.dispatchEvent(new CustomEvent('vghtpe:pdf-drop', { detail: { paths: ['C:/should-not-be-added.pdf'], rejected: [] } })));
  await expect(page.locator('.selected-files .file-row')).toHaveCount(2);
  await page.getByRole('button', { name: '取消任務' }).click();
  await expect(page.getByText('已取消', { exact: true }).first()).toBeVisible({ timeout: 10_000 });
  await page.getByRole('button', { name: '清除全部', exact: true }).click();
  await expect(page.locator('.selected-files .file-row')).toHaveCount(0);
  await expect(page.getByText('將 PDF 拖曳到這裡', { exact: true })).toBeVisible();
  expect(failures).toEqual([]);
});
