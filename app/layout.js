export const metadata = {
  title: 'Dashboard Monitoring Akses Toilet Karyawan',
};

export default function RootLayout({ children }) {
  return (
    <html lang="id">
      <body style={{ margin: 0, fontFamily: 'system-ui, sans-serif', background: '#0f1115', color: '#e6e6e6' }}>
        {children}
      </body>
    </html>
  );
}
