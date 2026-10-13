"use client";

import Link from "next/link";
import { useParams } from "next/navigation";
import { Check, ChevronRight, Heart, Minus, Plus, ShieldCheck, Star, Truck } from "lucide-react";
import { useMemo, useState } from "react";
import ProductCard, { ProductVisual } from "@/components/ProductCard";
import { formatPrice, getProduct, products } from "@/data/catalog";
import { useStore } from "@/context/StoreProvider";

export default function ProductDetailPage() {
  const params = useParams<{ slug: string }>();
  const product = getProduct(params.slug);
  const { dispatch, wishlist } = useStore();
  const [variant, setVariant] = useState(product?.variants[0]?.id ?? "");
  const [quantity, setQuantity] = useState(1);
  const related = useMemo(
    () => products.filter((item) => item.sport === product?.sport && item.id !== product?.id).slice(0, 4),
    [product],
  );

  if (!product) {
    return <main className="empty-state container"><p className="eyebrow">Không tìm thấy sản phẩm</p><h1>Sản phẩm này chưa có trong bản demo.</h1><Link className="btn btn-primary" href="/san-pham">Về danh mục</Link></main>;
  }

  const selected = product.variants.find((item) => item.id === variant) ?? product.variants[0];
  const liked = wishlist.includes(product.id);
  const addToCart = () => {
    for (let i = 0; i < quantity; i += 1) dispatch({ type: "add", productId: product.id, variantId: selected.id });
  };

  return <main>
    <div className="container breadcrumbs"><Link href="/">Trang chủ</Link><ChevronRight size={14}/><Link href="/san-pham">Sản phẩm</Link><ChevronRight size={14}/><span>{product.name}</span></div>
    <section className="container product-detail">
      <div className="detail-gallery"><ProductVisual product={product} className="detail-visual"/><div className="detail-thumbs" aria-label="Ảnh sản phẩm"><button className="thumb is-active" aria-label="Ảnh chính"><ProductVisual product={product}/></button><button className="thumb" aria-label="Mặt bên sản phẩm"><ProductVisual product={product}/></button><button className="thumb" aria-label="Chi tiết sản phẩm"><ProductVisual product={product}/></button></div></div>
      <div className="detail-copy">
        <span className="eyebrow">{product.brand} · {product.sport}</span>
        <h1>{product.name}</h1>
        <div className="rating detail-rating"><Star size={17} fill="currentColor"/><b>{product.rating}</b><span>{product.reviews} đánh giá</span><span className="stock-dot">Còn {selected.stock} sản phẩm</span></div>
        <p className="detail-description">{product.description}</p>
        <div className="detail-price"><strong>{formatPrice(product.price)}</strong>{product.oldPrice && <del>{formatPrice(product.oldPrice)}</del>}</div>
        <div className="choice-block"><b>Phiên bản</b><div className="variant-picker">{product.variants.map((item)=><button key={item.id} className={variant===item.id?"is-active":""} onClick={()=>setVariant(item.id)}>{item.label}<small>{item.stock > 0 ? `Còn ${item.stock}` : "Hết hàng"}</small></button>)}</div></div>
        <div className="purchase-row"><div className="quantity-control"><button aria-label="Giảm số lượng" onClick={()=>setQuantity(Math.max(1,quantity-1))}><Minus size={17}/></button><span>{quantity}</span><button aria-label="Tăng số lượng" onClick={()=>setQuantity(Math.min(selected.stock,quantity+1))}><Plus size={17}/></button></div><button className="btn btn-primary detail-add" onClick={addToCart}>Thêm vào giỏ · {formatPrice(product.price*quantity)}</button><button className={`icon-btn detail-wish ${liked?"is-active":""}`} aria-label="Yêu thích sản phẩm" onClick={()=>dispatch({type:"wishlist",productId:product.id})}><Heart size={20} fill={liked?"currentColor":"none"}/></button></div>
        <div className="service-list"><span><Check size={17}/>Hàng chính hãng có chứng từ</span><span><Truck size={17}/>Miễn phí giao hàng từ 1 triệu</span><span><ShieldCheck size={17}/>Đổi sản phẩm trong 7 ngày</span></div>
      </div>
    </section>
    <section className="detail-band"><div className="container detail-specs"><div><p className="eyebrow">Thông tin sản phẩm</p><h2>Thông số quan trọng</h2><p>Mọi thông tin trong bản mẫu là dữ liệu minh họa và có thể thay bằng catalog thực tế của cửa hàng.</p></div><dl>{Object.entries(product.specs).map(([label,value])=><div key={label}><dt>{label}</dt><dd>{value}</dd></div>)}</dl></div></section>
    <section className="section container"><div className="section-heading"><div><p className="eyebrow">Cùng môn chơi</p><h2>Sản phẩm liên quan</h2></div><Link href={`/san-pham?sport=${encodeURIComponent(product.sport)}`}>Xem tất cả <ChevronRight size={16}/></Link></div><div className="product-grid">{related.map((item)=><ProductCard key={item.id} product={item}/>)}</div></section>
    <div className="mobile-product-bar"><div><small>{product.name}</small><b>{formatPrice(product.price)}</b></div><button className="btn btn-primary" onClick={addToCart}>Thêm vào giỏ</button></div>
  </main>;
}
