import type { Metadata } from "next";
import { Be_Vietnam_Pro } from "next/font/google";
import "./globals.css";
import { StoreProvider } from "@/context/StoreProvider";
import { SiteChrome } from "@/components/SiteChrome";
const beVietnam=Be_Vietnam_Pro({subsets:["latin","vietnamese"],weight:["400","500","600","700","800"],variable:"--font-be-vietnam"});
export const metadata:Metadata={title:"Cửa hàng dụng cụ thể thao",description:"Website mẫu thương mại điện tử cầu lông và pickleball."};
export default function RootLayout({children}:{children:React.ReactNode}){return <html lang="vi" className={beVietnam.variable} data-scroll-behavior="smooth"><body><StoreProvider><SiteChrome>{children}</SiteChrome></StoreProvider></body></html>}
