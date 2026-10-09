export interface Report {
  id: string;
  name: string;
  report_type: string;
  category: string;
}

export interface Settings {
  output_folder: string;
  extract_folder: string;
  open_folder_after_completion: boolean;
}

export interface Bootstrap {
  token: string;
  reports: Report[];
  report_types: string[];
  last_month: string;
  settings: Settings;
  native_dialogs: boolean;
  catalog_revision: string;
}

export type TaskStatus = 'running' | 'cancelling' | 'cancelled' | 'completed' | 'failed';
export interface Task {
  id: string;
  kind: 'download' | 'extract';
  title: string;
  status: TaskStatus;
  total: number;
  processed: number;
  successful: number;
  skipped: number;
  failed: number;
  cancelled: number;
  progress: number;
  current_item: string;
  message: string;
  output_folder: string;
  logs: { time: string; message: string }[];
}

export interface DownloadInput {
  report_ids: string[];
  catalog_revision: string;
  start_month: string;
  end_month: string;
}
export interface ExtractInput {
  paths: string[];
  keyword: string;
  keep_contractor_page: boolean;
}
export const isTaskActive = (task: Task | null) => task?.status === 'running' || task?.status === 'cancelling';

export interface ReportInput {
  name: string;
  report_type: string;
  api_1: string;
  api_2: string;
}
export interface CatalogReport extends Report, ReportInput {}
export interface Catalog {
  reports: CatalogReport[];
  report_types: string[];
  revision: string;
  data_path: string;
  backup_path: string;
}
