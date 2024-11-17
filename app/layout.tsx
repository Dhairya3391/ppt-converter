import './globals.css'

export const metadata = {
  title: 'PPT to PDF Converter',
  description: 'Convert your PowerPoint presentations to PDF',
}

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
