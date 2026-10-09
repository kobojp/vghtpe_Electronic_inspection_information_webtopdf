type IconName = 'download' | 'files' | 'settings' | 'folder' | 'search' | 'arrow' | 'check' | 'stop' | 'plus' | 'catalog';
const paths: Record<IconName, string> = {
  plus: 'M12 4v16M4 12h16',
  catalog: 'M4 3h16v18H4zM8 7h8M8 12h8M8 17h5',
  download: 'M12 3v12m-5-5 5 5 5-5M5 16v4h14v-4',
  files: 'M8 3h9l4 4v14H8zM17 3v5h4M4 7H3v14h1',
  settings: 'M12 8a4 4 0 1 0 0 8 4 4 0 0 0 0-8M12 3v2m0 14v2M3 12h2m14 0h2M5.6 5.6 7 7m10 10 1.4 1.4M5.6 18.4 7 17M17 7l1.4-1.4',
  folder: 'M3 7V5h6l2 2h10v13H3z',
  search: 'M10.5 3a7.5 7.5 0 1 0 0 15 7.5 7.5 0 0 0 0-15M16 16l5 5',
  arrow: 'M5 12h14m-5-5 5 5-5 5',
  check: 'm5 12 4 4L19 6',
  stop: 'M6 6h12v12H6z',
};
export function Icon({ name, size = 20 }: { name: IconName; size?: number }) {
  return <svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true"><path d={paths[name]} /></svg>;
}
