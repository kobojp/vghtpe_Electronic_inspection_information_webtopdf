import { useEffect, useState } from 'react';
import type { Bootstrap, DownloadInput } from '../types';
import { Icon } from '../components/Icon';
import { monthCount } from '../utils';

interface Props { data: Bootstrap; busy: boolean; onStart: (input: DownloadInput) => Promise<void>; onOpenFolder: () => Promise<void> }

export function Downloads({ data, busy, onStart, onOpenFolder }: Props) {
  const [type, setType] = useState('');
  const [search, setSearch] = useState('');
  const [selected, setSelected] = useState<Set<string>>(() => new Set());
  const [start, setStart] = useState(data.last_month);
  const [end, setEnd] = useState(data.last_month);
  useEffect(() => { setSelected(new Set()); }, [data.catalog_revision]);
  const query = search.trim().toLocaleLowerCase();
  const visible = data.reports.filter(report => (!type || report.report_type === type) && report.name.toLocaleLowerCase().includes(query));
  const months = monthCount(start, end);
  const allSelected = visible.length > 0 && visible.every(report => selected.has(report.id));

  function toggle(id: string) {
    setSelected(previous => { const next = new Set(previous); if (next.has(id)) next.delete(id); else next.add(id); return next; });
  }
  function selectVisible() {
    setSelected(previous => {
      const next = new Set(previous);
      for (const report of visible) { if (allSelected) next.delete(report.id); else next.add(report.id); }
      return next;
    });
  }
  return <>
    <header className="page-heading">
      <div><p className="eyebrow">REPORT WORKSPACE</p><h1>報表下載</h1><p>選擇需要的報表與月份，讓例行整理更輕鬆。</p></div>
      <button className="button secondary" onClick={() => void onOpenFolder()}><Icon name="folder" />開啟下載資料夾</button>
    </header>
    <div className="summary-strip">
      {[['消防', '月檢查與消防設備'], ['電力', '每日、每週與每月'], ['排水', '巡查與設備保養']].map(([category, description]) => <div className="summary-item" key={category}><span className={'category-dot ' + category} /><div><strong>{category}<span>{data.reports.filter(report => report.category === category).length} 份</span></strong><small>{description}</small></div></div>)}
    </div>
    <div className="download-layout">
    <section className="card report-card" aria-labelledby="report-list-heading">
      <div className="card-heading"><div><h2 id="report-list-heading">選擇報表</h2><p>可依類別篩選，或搜尋大樓、設備名稱。</p></div><span className="subtle-pill">共 {data.reports.length} 份報表</span></div>
      <div className="filters">
        <label className="field"><span>報表類別</span><select value={type} onChange={event => setType(event.target.value)} disabled={busy}><option value="">全部類別</option>{data.report_types.map(value => <option key={value}>{value}</option>)}</select></label>
        <label className="field search-field"><span>名稱搜尋</span><div className="input-icon"><Icon name="search" /><input value={search} onChange={event => setSearch(event.target.value)} placeholder="例如：長青樓、滅火器、消防泵浦" disabled={busy} /></div></label>
      </div>
      <div className="selection-toolbar"><span>顯示 <strong>{visible.length}</strong> 份 · 已選 <strong>{selected.size}</strong> 份</span><div><button className="text-button" disabled={busy || !visible.length} onClick={selectVisible}>{allSelected ? '取消選取目前清單' : '選取目前清單'}</button><button className="text-button muted" disabled={busy || !selected.size} onClick={() => setSelected(new Set())}>清除選取</button></div></div>
      <div className="report-table-scroll">
        <table className="report-table"><thead><tr><th className="checkbox-column"><input type="checkbox" aria-label="選取目前清單的全部報表" checked={allSelected} disabled={busy || !visible.length} onChange={selectVisible} /></th><th>報表名稱</th><th className="type-column">類別</th></tr></thead><tbody>{visible.map(report => <tr key={report.id} className={selected.has(report.id) ? 'selected-row' : ''}><td><input id={'report-' + report.id} type="checkbox" aria-label={'選取 ' + report.name} checked={selected.has(report.id)} disabled={busy} onChange={() => toggle(report.id)} /></td><td><label className="report-name" htmlFor={'report-' + report.id}>{report.name}</label></td><td><span className={'category-tag ' + report.category}>{report.report_type}</span></td></tr>)}</tbody></table>
        {!visible.length && <div className="empty-list">沒有符合的報表，請調整搜尋條件。</div>}
      </div>
    </section>
    <section className="card download-options" aria-labelledby="date-heading">
      <div><h2 id="date-heading">下載月份</h2><p>支援單月與跨年度區間。</p></div>
      <div className="date-fields"><label className="field"><span>起始月份</span><input type="month" min="2022-01" value={start} onChange={event => setStart(event.target.value)} disabled={busy} /></label><span className="date-arrow"><Icon name="arrow" /></span><label className="field"><span>結束月份</span><input type="month" min="2022-01" value={end} onChange={event => setEnd(event.target.value)} disabled={busy} /></label><button className="button secondary small" disabled={busy} onClick={() => { setStart(data.last_month); setEnd(data.last_month); }}>上個月</button></div>
      <div className="download-footer"><div><strong>{selected.size * months}<span> 份下載項目</span></strong><small>{months > 0 ? selected.size + ' 份報表 × ' + months + ' 個月份；有效的既有 PDF 會自動跳過。' : '請選擇有效月份，起始月份不可晚於結束月份。'}</small></div><button className="button primary" disabled={busy || !selected.size || !months} onClick={() => void onStart({ report_ids: [...selected], catalog_revision: data.catalog_revision, start_month: start, end_month: end })}><Icon name="download" />開始下載</button></div>
    </section>
    </div>
  </>;
}
