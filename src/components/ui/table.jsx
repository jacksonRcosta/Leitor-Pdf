import { cn } from '@/lib/utils'

export function Table({ className, ...props }) {
  return (
    <div className="w-full overflow-auto">
      <table className={cn('w-full caption-bottom text-sm', className)} {...props} />
    </div>
  )
}

export function TableHeader({ className, ...props }) {
  return <thead className={cn('bg-bg', className)} {...props} />
}

export function TableBody({ className, ...props }) {
  return <tbody className={cn('[&_tr:last-child]:border-0', className)} {...props} />
}

export function TableRow({ className, ...props }) {
  return <tr className={cn('border-b border-border transition-colors hover:bg-bg/60', className)} {...props} />
}

export function TableHead({ className, ...props }) {
  return (
    <th className={cn('h-10 px-3 text-left align-middle text-xs font-semibold uppercase tracking-wide text-text-muted', className)} {...props} />
  )
}

export function TableCell({ className, ...props }) {
  return <td className={cn('px-3 py-3 align-middle', className)} {...props} />
}
