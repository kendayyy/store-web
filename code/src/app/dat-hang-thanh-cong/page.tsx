"use client";

import Link from "next/link";
import { Check, PackageCheck } from "lucide-react";
import { useSearchParams } from "next/navigation";
import { Suspense, useState } from "react";
import { formatPrice } from "@/data/catalog";
import type { OrderSummary } from "@/types";

function SuccessContent(){const params=useSearchParams();const[order]=useState<(OrderSummary&{name?:string})|null>(()=>{try{const saved=typeof window!=="undefined"?sessionStorage.getItem("last-order"):null;return saved?JSON.parse(saved):null}catch{return null}});const orderId=params.get("order")||order?.orderId||"TT-DEMO";return <main className="success-page container"><div className="success-mark"><PackageCheck size={42}/></div><p className="eyebrow">Đặt hàng thành công</p><h1>Cảm ơn{order?.name?`, ${order.name}`:" bạn"}.</h1><p>Đơn hàng đã được ghi nhận trong bản demo. Cửa hàng sẽ xác nhận lại thông tin trước khi giao.</p><div className="success-details"><div><span>Mã đơn hàng</span><b>{orderId}</b></div><div><span>Số sản phẩm</span><b>{order?.itemCount??"-"}</b></div><div><span>Tổng thanh toán</span><b>{order?formatPrice(order.total):"-"}</b></div></div><ul><li><Check size={17}/>Xác nhận đơn hàng qua điện thoại</li><li><Check size={17}/>Chuẩn bị và bàn giao cho đơn vị vận chuyển</li><li><Check size={17}/>Cập nhật trạng thái tới khách hàng</li></ul><div className="success-actions"><Link className="btn btn-primary" href="/">Về trang chủ</Link><Link className="btn btn-outline" href="/san-pham">Tiếp tục mua sắm</Link></div></main>}
export default function SuccessPage(){return <Suspense fallback={<main className="empty-state container">Đang tải đơn hàng...</main>}><SuccessContent/></Suspense>}
