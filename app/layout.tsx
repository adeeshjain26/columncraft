import type { Metadata, Viewport } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'ColumnCraft — Excel Column Extractor',
  description: 'Extract only the Excel columns you need. Your files stay on your device.',
  applicationName: 'ColumnCraft',
  appleWebApp: { capable: true, statusBarStyle: 'default', title: 'ColumnCraft' },
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  themeColor: '#f7f8f5',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return <html lang="en"><body>{children}</body></html>;
}
