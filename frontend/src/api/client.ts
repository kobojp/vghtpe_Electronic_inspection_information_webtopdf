import type { Bootstrap, Catalog, ReportInput, DownloadInput, ExtractInput, Settings, Task } from '../types';

let sessionToken = '';

async function request<T>(url: string, init: RequestInit = {}): Promise<T> {
  const headers = new Headers(init.headers);
  if (sessionToken) headers.set('X-App-Token', sessionToken);
  if (init.body) headers.set('Content-Type', 'application/json');
  const response = await fetch('/api' + url, { ...init, headers });
  const data = await response.json();
  if (!response.ok) {
    const detail = Array.isArray(data.detail)
      ? data.detail.map((item: { msg: string }) => item.msg).join('；')
      : data.detail;
    throw new Error(detail || '操作失敗，請稍後再試');
  }
  return data as T;
}

export const api = {
  async bootstrap(signal?: AbortSignal) {
    const data = await request<Bootstrap>('/bootstrap', { signal });
    sessionToken = data.token;
    return data;
  },
  catalog: (signal?: AbortSignal) => request<Catalog>('/catalog', { signal }),
  addReport: (input: ReportInput, revision: string) => request<Catalog>('/catalog/reports', { method: 'POST', body: JSON.stringify({ ...input, revision }) }),
  editReport: (id: string, input: ReportInput, revision: string) => request<Catalog>('/catalog/reports/' + encodeURIComponent(id), { method: 'PUT', body: JSON.stringify({ ...input, revision }) }),
  deleteReport: (id: string, revision: string) => request<Catalog>('/catalog/reports/' + encodeURIComponent(id), { method: 'DELETE', body: JSON.stringify({ revision }) }),
  task: (signal?: AbortSignal) => request<Task | null>('/tasks/current', { signal }),
  download: (input: DownloadInput) => request<Task>('/tasks/download', { method: 'POST', body: JSON.stringify(input) }),
  extract: (input: ExtractInput) => request<Task>('/tasks/extract', { method: 'POST', body: JSON.stringify(input) }),
  cancel: (id: string) => request<Task>('/tasks/' + id + '/cancel', { method: 'POST' }),
  saveSettings: (settings: Settings) => request<Settings>('/settings', { method: 'PUT', body: JSON.stringify(settings) }),
  selectPdfs: () => request<{ paths: string[] }>('/desktop/select-pdfs', { method: 'POST' }),
  selectFolder: () => request<{ path: string | null }>('/desktop/select-folder', { method: 'POST' }),
  openFolder: (kind: 'download' | 'extract' | 'task') => request('/desktop/open-folder', { method: 'POST', body: JSON.stringify({ kind }) }),
};
