import { useState } from 'react';
import type { Settings as SettingsData } from '../types';
import { Icon } from '../components/Icon';

interface Props { settings: SettingsData; busy: boolean; onSelectFolder: () => Promise<string | null>; onSave: (settings: SettingsData) => Promise<boolean> }

export function Settings({ settings, busy, onSelectFolder, onSave }: Props) {
  const [draft, setDraft] = useState(settings);
  const [saved, setSaved] = useState(false);
  function update(values: Partial<SettingsData>) { setDraft(previous => ({ ...previous, ...values })); setSaved(false); }
  async function choose(key: 'output_folder' | 'extract_folder') {
    const folder = await onSelectFolder();
    if (folder) update({ [key]: folder });
  }
  async function save() { setSaved(await onSave(draft)); }
  return <>
    <header className="page-heading"><div><p className="eyebrow">PREFERENCES</p><h1>設定</h1><p>選擇報表保存位置，調整工作完成後的行為。</p></div></header>
    <section className="card settings-card" aria-labelledby="output-heading"><div className="card-heading"><div><h2 id="output-heading">輸出資料夾</h2><p>設定儲存後，下次開啟程式會沿用。</p></div></div>
      {([['output_folder', '報表下載資料夾', '報表會依消防、電力、排水與月份建立子資料夾。'], ['extract_folder', 'PDF 擷取資料夾', '擷取結果會依處理日期建立子資料夾。']] as const).map(([key, title, description]) => <div className="folder-setting" key={key}><div className="field"><label htmlFor={key}>{title}</label><div className="folder-input"><input id={key} value={draft[key]} onChange={event => update({ [key]: event.target.value })} disabled={busy} /><button className="button secondary" disabled={busy} onClick={() => void choose(key)} aria-label={'選擇' + title}><Icon name="folder" />選擇</button></div></div><p>{description}</p></div>)}
      <label className="check-option"><input type="checkbox" checked={draft.open_folder_after_completion} onChange={event => update({ open_folder_after_completion: event.target.checked })} disabled={busy} /><div><strong>完成後開啟資料夾</strong><small>下載或擷取完成後，自動開啟該次的輸出資料夾。</small></div></label>
      <div className="download-footer"><span className="save-message" role="status">{saved ? <><Icon name="check" />設定已儲存</> : '設定會保存在這台電腦。'}</span><button className="button primary" disabled={busy || !draft.output_folder.trim() || !draft.extract_folder.trim()} onClick={() => void save()}><Icon name="check" />儲存設定</button></div>
    </section>
  </>;
}
