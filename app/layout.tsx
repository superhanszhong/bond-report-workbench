import type { Metadata } from "next";
import { headers } from "next/headers";
import "./globals.css";

export async function generateMetadata(): Promise<Metadata> {
  const requestHeaders = await headers();
  const host = requestHeaders.get("x-forwarded-host") || requestHeaders.get("host") || "localhost:3000";
  const protocol = requestHeaders.get("x-forwarded-proto") || (host.includes("localhost") ? "http" : "https");
  const base = new URL(`${protocol}://${host}`);
  const preview = new URL("/og.png", base).toString();
  return {
    metadataBase: base,
    title: "利率债发行工作台",
    description: "地方债日表、利差图、四周滚动分析与周报生成的一体化工作台。",
    icons: { icon: "/favicon.svg", shortcut: "/favicon.svg" },
    openGraph: { title: "利率债发行工作台", description: "地方债日表 · 一二级利差 · 四周滚动分析 · 周报生成", images: [{ url: preview, width: 1536, height: 1024 }] },
    twitter: { card: "summary_large_image", title: "利率债发行工作台", description: "地方债日表 · 一二级利差 · 四周滚动分析 · 周报生成", images: [preview] },
  };
}

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
