import type { Task, TaskStatus } from '../types';
import { isTaskActive } from '../types';
import { Icon } from './Icon';

const statusLabels: Record<TaskStatus, string> = { running: '執行中', cancelling: '正在取消', cancelled: '已取消', completed: '已完成', failed: '執行失敗' };
interface Props { task: Task | null; pending: boolean; onCancel: () => Promise<void>; onOpenFolder: () => Promise<void> }

export function TaskPanel({ task, pending, onCancel, onOpenFolder }: Props) {
  const active = isTaskActive(task);
  return <section className="task-panel" aria-label="任務進度">
    <div className="task-heading"><div className="task-title"><span className={'task-indicator ' + (active ? 'active' : '')} /><strong>{task ? task.title : '目前沒有執行中的任務'}</strong>{task && <span className={'status-label ' + task.status}>{statusLabels[task.status]}</span>}</div><div className="task-actions">{task && <button className="text-button" onClick={() => void onOpenFolder()}><Icon name="folder" size={16} />輸出資料夾</button>}{active && <button className="button cancel small" disabled={pending || task?.status === 'cancelling'} onClick={() => void onCancel()}><Icon name="stop" size={15} />{task?.status === 'cancelling' ? '正在取消…' : '取消任務'}</button>}</div></div>
    {task ? <><div className="progress-row"><span className="current-item" title={task.current_item}>{active ? task.current_item || '準備開始' : task.message}</span><span>{task.processed} / {task.total}<strong>{task.progress.toFixed(1)}%</strong></span></div><div className="progress-track" role="progressbar" aria-label="任務完成進度" aria-valuemin={0} aria-valuemax={100} aria-valuenow={task.progress}><div style={{ width: task.progress + '%' }} /></div><div className="task-bottom"><div className="task-counts"><span>成功 <b>{task.successful}</b></span><span>跳過 <b>{task.skipped}</b></span><span className={task.failed ? 'failure-count' : ''}>失敗 <b>{task.failed}</b></span><span>已取消 <b>{task.cancelled}</b></span></div><span className="task-hint">{active ? '切換頁面不會中斷任務' : '可以開始下一個任務'}</span></div><details className="task-logs"><summary>操作紀錄 <span>最近 {task.logs.length} 筆</span></summary><div className="log-list">{task.logs.map((entry, index) => <div key={index}><time>{entry.time}</time><span>{entry.message}</span></div>)}</div></details></> : <p className="idle-message">選擇報表開始下載，或到 PDF 擷取頁加入檔案。任務進度會顯示在這裡。</p>}
  </section>;
}
