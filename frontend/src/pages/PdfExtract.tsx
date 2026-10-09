import { useEffect, useRef, useState } from 'react';
import type { DragEvent } from 'react';
import type { ExtractInput } from '../types';
import { Icon } from '../components/Icon';
import { fileName, mergePdfPaths } from '../utils';

interface Props { active: boolean; busy: boolean; onSelect: () => Promise<string[]>; onStart: (input: ExtractInput) => Promise<void>; onOpenFolder: () => Promise<void> }
interface DropResult { paths: string[]; rejected: string[]; error?: string }

export function PdfExtract({ active, busy, onSelect, onStart, onOpenFolder }: Props) {
  const [paths, setPaths] = useState<string[]>([]);
  const [keyword, setKeyword] = useState('');
  const [keepContractor, setKeepContractor] = useState(true);
  const [dragging, setDragging] = useState(false);
  const [dropMessage, setDropMessage] = useState('');
  const dragDepth = useRef(0);
  const zoneRef = useRef<HTMLElement>(null);

  useEffect(() => {
    const receive = (event: Event) => {
      if (!active || busy) return;
      const result = (event as CustomEvent<DropResult>).detail;
      if (!result || !Array.isArray(result.paths) || !Array.isArray(result.rejected)) return;
      setPaths(previous => mergePdfPaths(previous, result.paths));
      const skipped = result.rejected.length ? '已略過非有效 PDF 或無法讀取的檔案：' + result.rejected.join('、') + '。' : '';
      setDropMessage(result.error || (result.paths.length ? '已加入拖入的 PDF，重複檔案會自動略過。' : '') + skipped);
      dragDepth.current = 0;
      setDragging(false);
    };
    window.addEventListener('vghtpe:pdf-drop', receive);
    // Prevent the browser from navigating to dropped files, including outside the zone.
    const preventFileNavigation = (event: globalThis.DragEvent) => {
      if (Array.from(event.dataTransfer?.types || []).includes('Files')) event.preventDefault();
    };
    window.addEventListener('dragover', preventFileNavigation);
    window.addEventListener('drop', preventFileNavigation);
    return () => {
      window.removeEventListener('vghtpe:pdf-drop', receive);
      window.removeEventListener('dragover', preventFileNavigation);
      window.removeEventListener('drop', preventFileNavigation);
    };
  }, [active, busy]);
  useEffect(() => {
    if (busy || !active) { dragDepth.current = 0; setDragging(false); }
  }, [active, busy]);

  async function chooseFiles() {
    const chosen = await onSelect();
    setPaths(previous => mergePdfPaths(previous, chosen));
    if (chosen.length) setDropMessage('');
  }
  function isFileDrag(event: DragEvent<HTMLElement>) {
    return Array.from(event.dataTransfer.types).includes('Files');
  }
  function enter(event: DragEvent<HTMLElement>) {
    if (!isFileDrag(event)) return;
    event.preventDefault();
    if (busy) return;
    dragDepth.current += 1;
    setDragging(true);
  }
  function leave(event: DragEvent<HTMLElement>) {
    if (!isFileDrag(event)) return;
    dragDepth.current = Math.max(0, dragDepth.current - 1);
    if (!dragDepth.current) setDragging(false);
  }
  function drop(event: DragEvent<HTMLElement>) {
    event.preventDefault();
    dragDepth.current = 0;
    setDragging(false);
    if (busy || !event.dataTransfer.files.length) return;
    if (zoneRef.current?.dataset.nativeDropReady !== 'true') {
      setDropMessage('目前無法取得拖入檔案的路徑，請在桌面程式使用拖曳，或點選「選擇 PDF」。');
    }
  }

  return <>
    <header className="page-heading"><div><p className="eyebrow">PDF WORKSPACE</p><h1>PDF 擷取</h1><p>找到關鍵字所在頁面，將需要的內容獨立保存。</p></div><button className="button secondary" onClick={() => void onOpenFolder()}><Icon name="folder" />開啟擷取資料夾</button></header>
    <section ref={zoneRef} id="pdf-drop-zone" className={'card pdf-drop-zone ' + (dragging ? 'drag-active' : '')} aria-labelledby="pdf-files-heading" aria-disabled={busy} onDragEnter={enter} onDragLeave={leave} onDragOver={event => { if (isFileDrag(event)) { event.preventDefault(); event.dataTransfer.dropEffect = busy ? 'none' : 'copy'; } }} onDrop={drop}>
      <div className="card-heading"><div><h2 id="pdf-files-heading">選擇 PDF 檔案</h2><p>支援單份或多份檔案，每份來源會分別產生擷取結果。</p></div><button className="button secondary" disabled={busy} onClick={() => void chooseFiles()}><Icon name="files" />選擇 PDF</button></div>
      {!paths.length ? <div className="file-empty"><span className="empty-icon"><Icon name="files" size={32} /></span><strong>{dragging ? '放開滑鼠即可加入 PDF' : '將 PDF 拖曳到這裡'}</strong><p>可一次拖入多份 PDF，或點選「選擇 PDF」加入檔案。</p></div> : <><div className="selected-files">{paths.map(path => <div className="file-row" key={path}><Icon name="files" /><div><strong>{fileName(path)}</strong><small title={path}>{path}</small></div><button className="text-button" disabled={busy} aria-label={'移除 ' + fileName(path)} onClick={() => setPaths(previous => previous.filter(item => item !== path))}>移除</button></div>)}</div><div className="selection-toolbar"><span>已加入 <strong>{paths.length}</strong> 份 PDF</span><button className="text-button muted" disabled={busy} onClick={() => { setPaths([]); setDropMessage(''); }}>清除全部</button></div><p className="drop-hint">{dragging ? '放開滑鼠即可加入其他 PDF' : '可繼續將其他 PDF 拖曳到此區域，重複檔案會自動略過。'}</p></>}
      {dropMessage && <p className="drop-message" role="status">{dropMessage}</p>}
    </section>
    <section className="card extraction-options" aria-labelledby="keyword-heading"><div className="card-heading"><div><h2 id="keyword-heading">設定擷取條件</h2><p>搜尋會處理 PDF 文字中的空白與換行。</p></div></div><label className="field"><span>搜尋關鍵字</span><input value={keyword} onChange={event => setKeyword(event.target.value)} maxLength={200} placeholder="例如：醫學科技大樓 2F" disabled={busy} /></label><label className="check-option"><input type="checkbox" checked={keepContractor} onChange={event => setKeepContractor(event.target.checked)} disabled={busy} /><div><strong>保留承商簽核頁</strong><small>找到關鍵字頁面時，一併保留「承商駐院工程師」簽核頁。</small></div></label><div className="note">結果依處理日期建立資料夾；同名檔案會加上月份或編號，保留各份來源的內容。掃描影像若沒有文字層，可能找不到關鍵字。</div><div className="download-footer"><span>{paths.length} 份 PDF 待處理</span><button className="button primary" disabled={busy || !paths.length || !keyword.trim()} onClick={() => void onStart({ paths, keyword: keyword.trim(), keep_contractor_page: keepContractor })}><Icon name="files" />開始擷取</button></div></section>
  </>;
}
