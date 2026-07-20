import Link from 'next/link';
import { cn } from '@/lib/utils';
export function Button({href,children,variant='primary',className}:{href:string;children:React.ReactNode;variant?:'primary'|'ghost';className?:string}){return <Link className={cn('inline-flex items-center justify-center rounded-full px-6 py-3 text-sm font-semibold transition hover:scale-[1.02] focus:outline-none focus:ring-2 focus:ring-[#2BA8FF]',variant==='primary'?'bg-white text-black hover:bg-[#D9D9D9]':'border border-white/20 bg-white/5 text-white hover:bg-white/10',className)} href={href}>{children}</Link>}
