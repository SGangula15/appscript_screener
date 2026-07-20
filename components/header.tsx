import Link from 'next/link';
import { nav } from '@/constants/site';
const slug=(s:string)=>'/' + s.toLowerCase().replaceAll(' ','-');
export function Header(){return <header className="fixed top-0 z-50 w-full border-b border-white/10 bg-[#050505]/70 backdrop-blur-xl"><nav className="mx-auto flex max-w-7xl items-center justify-between px-5 py-4"><Link href="/" className="font-semibold tracking-tight">AutoSys<span className="text-[#2BA8FF]">.</span></Link><div className="hidden gap-6 text-xs text-white/70 lg:flex">{nav.map(n=><Link key={n} href={slug(n)} className="hover:text-white">{n}</Link>)}</div><Link href="/contact" className="rounded-full border border-white/15 px-4 py-2 text-xs">Start Project</Link></nav></header>}
