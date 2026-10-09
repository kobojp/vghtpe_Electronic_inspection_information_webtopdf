import { useEffect, useRef, useState } from 'react';
import type { Catalog, CatalogReport, ReportInput } from '../types';
import { Icon } from '../components/Icon';

interface Props {
  catalog: Catalog | null;
  busy: boolean;
  onReload: () => Promise<Catalog | undefined>;
  onSave: (input: ReportInput, revision: string, id?: string) => Promise<Catalog | undefined>;
  onDelete: (id: string, revision: string) => Promise<Catalog | undefined>;
}
interface Editor { id?: string; revision: string; values: ReportInput }
interface Removal { report: CatalogReport; revision: string }
const emptyReport = (): ReportInput => ({ name: '', report_type: '消防', api_1: '', api_2: '' });

export function ReportManager({ catalog, busy, onReload, onSave, onDelete }: Props) {
  const [type, setType] = useState('');
  const [search, setSearch] = useState('');
  const [editor, setEditor] = useState<Editor | null>(null);
  const [removal, setRemoval] = useState<Removal | null>(null);
  const [message, setMessage] = useState('');
  const confirmRef = useRef<HTMLDialogElement>(null);
  useEffect(() => { if (removal && confirmRef.current && !confirmRef.current.open) confirmRef.current.showModal(); }, [removal]);
  const query = search.trim().toLocaleLowerCase();
  const visible = catalog?.reports.filter(report => (!type || report.report_type === type) && report.name.toLocaleLowerCase().includes(query)) || [];

  function edit(report?: CatalogReport) {
    if (!catalog) return;
    setMessage('');
    setRemoval(null);
    setEditor({ id: report?.id, revision: catalog.revision, values: report
      ? { name: report.name, report_type: report.report_type, api_1: report.api_1, api_2: report.api_2 }
      : emptyReport() });
  }
  function update(key: keyof ReportInput, value: string) {
    setEditor(previous => previous ? { ...previous, values: { ...previous.values, [key]: value } } : null);
  }
  async function save() {
    if (!editor) return;
    const saved = await onSave(editor.values, editor.revision, editor.id);
    if (saved) {
      setMessage(editor.id ? '報表已更新，下載清單已同步。' : '報表已新增，下載清單已同步。');
      setEditor(null);
    }
  }
  async function remove() {
    if (!removal) return;
    const saved = await onDelete(removal.report.id, removal.revision);
    if (saved) { setRemoval(null); setMessage('報表已刪除，已下載的 PDF 保留。'); }
  }
  async function reload() {
    if (await onReload()) { setEditor(null); setRemoval(null); setMessage('已重新載入報表資料。'); }
  }

  return <>
    <header className="page-heading">
      <div><p className="eyebrow">REPORT CATALOG</p><h1>報表管理</h1><p>維護報表清單，保存後立即同步至下載頁面。</p></div>
      <div className="catalog-heading-actions"><button className="button secondary" disabled={busy} onClick={() => void reload()}>重新載入資料</button><button className="button primary" disabled={busy || !catalog} onClick={() => edit()}><Icon name="plus" />新增報表</button></div>
    </header>
    {message && <div className="catalog-message" role="status"><Icon name="check" size={16} />{message}</div>}
    {!catalog ? <section className="card"><p>正在讀取報表資料；若讀取失敗，可使用「重新載入資料」重試。</p></section> :
    <div className={'catalog-layout ' + (editor ? 'has-editor' : '')}>
      <section className="card catalog-list" aria-labelledby="catalog-list-heading">
        <div className="card-heading"><div><h2 id="catalog-list-heading">報表清單</h2><p>依類別查詢，選擇報表後修改或刪除。</p></div><span className="subtle-pill">共 {catalog.reports.length} 份報表</span></div>
        <div className="filters"><label className="field"><span>管理報表類別</span><select value={type} onChange={event => setType(event.target.value)}><option value="">全部類別</option>{catalog.report_types.map(value => <option key={value}>{value}</option>)}</select></label><label className="field"><span>搜尋管理報表</span><div className="input-icon"><Icon name="search" /><input value={search} onChange={event => setSearch(event.target.value)} placeholder="輸入報表名稱" /></div></label></div>
        <div className="selection-toolbar"><span>顯示 <strong>{visible.length}</strong> 份</span><span>刪除清單資料不會移除 PDF</span></div>
        <div className="report-table-scroll catalog-table-scroll"><table className="report-table"><thead><tr><th>報表名稱</th><th className="type-column">類別</th><th className="catalog-actions-column">操作</th></tr></thead><tbody>{visible.map(report => <tr key={report.id}><td>{report.name}</td><td><span className={'category-tag ' + report.category}>{report.report_type}</span></td><td><div className="catalog-row-actions"><button className="text-button" aria-label={'修改 ' + report.name} disabled={busy} onClick={() => edit(report)}>修改</button><button className="text-button danger-text" aria-label={'刪除 ' + report.name} disabled={busy} onClick={() => { setEditor(null); setMessage(''); setRemoval({ report, revision: catalog.revision }); }}>刪除</button></div></td></tr>)}</tbody></table>{!visible.length && <div className="empty-list">沒有符合的報表，請調整搜尋條件或新增報表。</div>}</div>
        <div className="catalog-storage"><strong>資料檔案</strong><span>{catalog.data_path}</span><small>每次保存前，會將上一版備份至 data.json.bak。程式更新時請保留自訂的 data.json。</small></div>
      </section>
      {editor && <section className="card catalog-editor" aria-labelledby="catalog-editor-heading">
        <div className="card-heading"><div><h2 id="catalog-editor-heading">{editor.id ? '修改報表' : '新增報表'}</h2><p>填寫名稱、類型與原報表網址中的參數。</p></div></div>
        <form onSubmit={event => { event.preventDefault(); void save(); }}>
          <fieldset disabled={busy}><label className="field"><span>報表名稱</span><input required maxLength={180} value={editor.values.name} onChange={event => update('name', event.target.value)} placeholder="例如：長青樓滅火器月檢查" /></label>
          <label className="field"><span>報表類型</span><select value={editor.values.report_type} onChange={event => update('report_type', event.target.value)}>{catalog.report_types.map(value => <option key={value}>{value}</option>)}</select></label>
          <label className="field"><span>網址參數一（api_1）</span><input required maxLength={100} pattern="/[0-9]+/" value={editor.values.api_1} onChange={event => update('api_1', event.target.value)} placeholder="/16/" /><small>月份前的參數，例如 /16/。</small></label>
          <label className="field"><span>網址參數二（api_2）</span><input required maxLength={100} pattern="/?[0-9]+/[0-9]+" value={editor.values.api_2} onChange={event => update('api_2', event.target.value)} placeholder="/16/32" /><small>月份後的參數，例如 /16/32。</small></label></fieldset>
          <p className="note">月份由下載頁指定。保存前會檢查名稱與參數格式；同一大類不可使用相同名稱。</p>
          <div className="catalog-form-actions"><button className="button secondary" type="button" disabled={busy} onClick={() => setEditor(null)}>取消編輯</button><button className="button primary" type="submit" disabled={busy}>儲存報表</button></div>
        </form>
      </section>}
    </div>}
    {removal && <dialog ref={confirmRef} className="catalog-confirm" aria-labelledby="delete-report-heading" onCancel={event => { event.preventDefault(); if (!busy) setRemoval(null); }}>
      <h2 id="delete-report-heading">確認刪除報表</h2><p>要從清單移除「{removal.report.name}」嗎？</p><small>已下載的 PDF 保留。保存前會備份目前資料。</small>
      <div className="catalog-form-actions"><button className="button secondary" disabled={busy} onClick={() => setRemoval(null)}>取消刪除</button><button className="button cancel" disabled={busy} onClick={() => void remove()}>確認刪除</button></div>
    </dialog>}
  </>;
}
