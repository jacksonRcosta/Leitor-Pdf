import { useState } from 'react'
import { Menu, X } from 'lucide-react'

export default function AppShell({ appName = 'App', nav = [], topbarActions = null, children }) {
  const [mobileOpen, setMobileOpen] = useState(false)

  const NavItems = () => (
    <nav className="flex flex-col gap-1 px-3">
      {nav.map((item) => {
        const Icon = item.icon
        return (
          <a
            key={item.label}
            href={item.href}
            className={[
              'flex items-center gap-3 rounded-lg px-3 py-2 text-sm font-medium transition-colors duration-150',
              item.active
                ? 'bg-primary/10 text-primary'
                : 'text-text-muted hover:bg-border/50 hover:text-text',
            ].join(' ')}
          >
            {Icon ? <Icon className="h-4 w-4 shrink-0" aria-hidden /> : null}
            <span>{item.label}</span>
          </a>
        )
      })}
    </nav>
  )

  return (
    <div className="min-h-screen bg-bg text-text">
      {/* Sidebar desktop */}
      <aside className="fixed inset-y-0 left-0 z-30 hidden w-64 flex-col border-r border-border bg-surface md:flex">
        <div className="flex h-14 items-center border-b border-border px-6">
          <img src="/logo.png" alt="Asa Branca Distribuidora" className="h-7 w-auto" />
        </div>
        <div className="flex-1 overflow-y-auto py-4">
          <NavItems />
        </div>
      </aside>

      {/* Drawer mobile */}
      {mobileOpen && (
        <div className="fixed inset-0 z-40 md:hidden">
          <div className="absolute inset-0 bg-text/40" onClick={() => setMobileOpen(false)} aria-hidden />
          <aside className="absolute inset-y-0 left-0 flex w-64 flex-col bg-surface">
            <div className="flex h-14 items-center justify-between border-b border-border px-6">
              <img src="/logo.png" alt="Asa Branca Distribuidora" className="h-7 w-auto" />
              <button
                onClick={() => setMobileOpen(false)}
                className="rounded-md p-1 text-text-muted hover:bg-border/50 hover:text-text focus-visible:ring-2 focus-visible:ring-primary"
                aria-label="Fechar menu"
              >
                <X className="h-5 w-5" />
              </button>
            </div>
            <div className="flex-1 overflow-y-auto py-4">
              <NavItems />
            </div>
          </aside>
        </div>
      )}

      {/* Coluna principal */}
      <div className="md:pl-64">
        <header className="sticky top-0 z-20 flex h-14 items-center gap-3 border-b border-border bg-surface/80 px-4 backdrop-blur md:px-8">
          <button
            onClick={() => setMobileOpen(true)}
            className="rounded-md p-1 text-text-muted hover:bg-border/50 hover:text-text focus:outline-none focus-visible:ring-2 focus-visible:ring-primary md:hidden"
            aria-label="Abrir menu"
          >
            <Menu className="h-5 w-5" />
          </button>
          <div className="ml-auto flex items-center gap-3">{topbarActions}</div>
        </header>
        <main className="mx-auto max-w-7xl px-4 py-6 md:px-8">{children}</main>
      </div>
    </div>
  )
}
