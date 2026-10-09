import { useCallback, useEffect, useState } from 'react';
import { api } from './api/client';
import { Icon } from './components/Icon';
import { TaskPanel } from './components/TaskPanel';
import { useTask } from './hooks/useTask';
import { Downloads } from './pages/Downloads';
import { PdfExtract } from './pages/PdfExtract';
import { ReportManager } from './pages/ReportManager';
import { Settings } from './pages/Settings';
import type { Bootstrap, Catalog, ReportInput, DownloadInput, ExtractInput, Settings as SettingsData } from './types';
import { isTaskActive } from './types';

type Page = 'downloads' | 'pdf' | 'catalog' | 'settings';
const navigation = [
  { id: 'downloads', title: '報表下載', description: '瀏覽與下載報表', icon: 'download' },
  { id: 'pdf', title: 'PDF 擷取', description: '搜尋需要的頁面', icon: 'files' },
  { id: 'catalog', title: '報表管理', description: '新增與維護報表', icon: 'catalog' },
  { id: 'settings', title: '設定', description: '資料夾與偏好', icon: 'settings' },
] as const;

export default function App() {
  const [page, setPage] = useState<Page>('downloads');
  const [data, setData] = useState<Bootstrap | null>(null);
  const [catalog, setCatalog] = useState<Catalog | null>(null);
  const [error, setError] = useState('');
  const [pending, setPending] = useState(false);
  const { task, setTask, connectionError } = useTask(data !== null);
  const busy = pending || isTaskActive(task) || Boolean(connectionError);

  useEffect(() => {
    const controller = new AbortController();
    api.bootstrap(controller.signal).then(value => { if (!controller.signal.aborted) setData(value); }).catch(reason => {
      if (!controller.signal.aborted) setError(reason instanceof Error ? reason.message : '無法啟動本機服務');
    });
    return () => controller.abort();
  }, []);

  const updateCatalog = useCallback((value: Catalog) => {
    setCatalog(value);
    setData(previous => previous ? {
      ...previous, reports: value.reports, report_types: value.report_types, catalog_revision: value.revision,
    } : previous);
  }, []);
  const ready = data !== null;
  useEffect(() => {
    if (page !== 'catalog' || !ready || catalog) return;
    const controller = new AbortController();
    api.catalog(controller.signal).then(value => {
      if (!controller.signal.aborted) updateCatalog(value);
    }).catch(reason => {
      if (!controller.signal.aborted) setError(reason instanceof Error ? reason.message : '無法讀取報表資料');
    });
    return () => controller.abort();
  }, [page, ready, catalog, updateCatalog]);

  async function perform<T>(operation: () => Promise<T>): Promise<T | undefined> {
    setPending(true);
    setError('');
    try { return await operation(); }
    catch (reason) { setError(reason instanceof Error ? reason.message : '操作失敗'); return undefined; }
    finally { setPending(false); }
  }
  async function download(input: DownloadInput) { const result = await perform(() => api.download(input)); if (result) setTask(result); }
  async function extract(input: ExtractInput) { const result = await perform(() => api.extract(input)); if (result) setTask(result); }
  async function cancel() { if (task) { const result = await perform(() => api.cancel(task.id)); if (result) setTask(result); } }
  async function openFolder(kind: 'download' | 'extract' | 'task') { await perform(() => api.openFolder(kind)); }
  async function selectPdfs() { return (await perform(api.selectPdfs))?.paths || []; }
  async function selectFolder() { return (await perform(api.selectFolder))?.path || null; }
  async function reloadCatalog() {
    const value = await perform(() => api.catalog());
    if (value) updateCatalog(value);
    return value;
  }
  async function saveReport(input: ReportInput, revision: string, id?: string) {
    const value = await perform(() => id ? api.editReport(id, input, revision) : api.addReport(input, revision));
    if (value) updateCatalog(value);
    return value;
  }
  async function deleteReport(id: string, revision: string) {
    const value = await perform(() => api.deleteReport(id, revision));
    if (value) updateCatalog(value);
    return value;
  }
  async function saveSettings(settings: SettingsData) {
    const saved = await perform(() => api.saveSettings(settings));
    if (!saved) return false;
    setData(previous => previous ? { ...previous, settings: saved } : previous);
    return true;
  }
  return <div className="app-shell" data-app-ready={data ? 'true' : undefined}>
    <aside className="sidebar"><div className="brand"><img src="/vghtpe.png" alt="臺北榮總" /><div><strong>臺北榮總</strong><span>水電消防報表</span></div></div><div className="sidebar-caption">工作區</div><nav aria-label="主要導覽">{navigation.map(item => <button key={item.id} className={'nav-item ' + (page === item.id ? 'nav-active' : '')} aria-current={page === item.id ? 'page' : undefined} onClick={() => setPage(item.id)}><Icon name={item.icon} /><div><strong>{item.title}</strong><small>{item.description}</small></div></button>)}</nav><div className="sidebar-footer"><span className={'connection-dot ' + (data && !connectionError ? 'connected' : '')} /><div><strong>{data && !connectionError ? '本機服務已連線' : '等待本機服務'}</strong><small>報表與 PDF 儲存在這台電腦</small></div></div></aside>
    <div className="workspace"><main className="main-content">{(error || connectionError) && <div className="error-banner" role="alert"><span>{connectionError || error}</span>{!connectionError && <button aria-label="關閉錯誤訊息" onClick={() => setError('')}>×</button>}</div>}
      {data ? <><div hidden={page !== 'downloads'}><Downloads data={data} busy={busy} onStart={download} onOpenFolder={() => openFolder('download')} /></div><div hidden={page !== 'pdf'}><PdfExtract active={page === 'pdf'} busy={busy} onSelect={selectPdfs} onStart={extract} onOpenFolder={() => openFolder('extract')} /></div><div hidden={page !== 'catalog'}><ReportManager catalog={catalog} busy={pending || Boolean(connectionError)} onReload={reloadCatalog} onSave={saveReport} onDelete={deleteReport} /></div><div hidden={page !== 'settings'}><Settings settings={data.settings} busy={busy} onSelectFolder={selectFolder} onSave={saveSettings} /></div></> : <div className="loading-view"><div className="loading-mark"><Icon name="download" size={32} /></div><h1>{error ? '無法載入工作區' : '正在準備工作區'}</h1><p>{error ? '請關閉並重新開啟程式，或查看錯誤日誌。' : '正在讀取報表目錄與你的設定…'}</p></div>}
    </main><TaskPanel task={task} pending={pending} onCancel={cancel} onOpenFolder={() => openFolder('task')} /></div>
  </div>;
}
