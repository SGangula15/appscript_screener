import type { Metadata } from 'next';
import { Inter } from 'next/font/google';
import './globals.css';
import { Header } from '@/components/header';
import { Footer } from '@/components/footer';
const inter = Inter({ subsets: ['latin'], display: 'swap' });
export const metadata: Metadata = { title: { default: 'AutoSys Tech Solutions', template: '%s | AutoSys Tech Solutions' }, description: 'End-to-end automotive engineering, embedded systems, AI, functional safety, cybersecurity, validation and staffing solutions.', openGraph: { title: 'AutoSys Tech Solutions', description: 'Engineering Intelligence. Driving Innovation.', type: 'website' } };
export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) { return <html lang="en" className={inter.className}><body><Header />{children}<Footer /></body></html>; }
