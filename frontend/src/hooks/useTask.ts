import { useEffect, useState } from 'react';
import { api } from '../api/client';
import type { Task } from '../types';

export function useTask(enabled: boolean) {
  const [task, setTask] = useState<Task | null>(null);
  const [connectionError, setConnectionError] = useState('');
  useEffect(() => {
    if (!enabled) return;
    const controller = new AbortController();
    let timer: ReturnType<typeof setTimeout>;
    async function poll() {
      try {
        const current = await api.task(controller.signal);
        if (!controller.signal.aborted) {
          setTask(current);
          setConnectionError('');
        }
      } catch (error) {
        if (!controller.signal.aborted) setConnectionError(error instanceof Error ? error.message : '無法連線到本機服務');
      } finally {
        if (!controller.signal.aborted) timer = setTimeout(poll, 1000);
      }
    }
    void poll();
    return () => { controller.abort(); clearTimeout(timer); };
  }, [enabled]);
  return { task, setTask, connectionError };
}
