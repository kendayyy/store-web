"use client";
import Link from "next/link";
import { Heart, Plus, Star } from "lucide-react";
import { formatPrice } from "@/data/catalog";
import { useStore } from "@/context/StoreProvider";
import type { Product } from "@/types";

export function ProductVisual({product,className=""}:{product:Product;className?:string}){
  return <div className={`product-visual tone-${product.tone} ${className}`} role="img" aria-label={`Ảnh minh họa ${product.name}`}><span>{product.category}</span><b>{product.brand}</b></div>;
}

export default function ProductCard({product}:{product:Product}){
 const {wishlist,dispatch}=useStore();const liked=wishlist.includes(product.id);
 return <article className="product-card">
  <div className="product-media"><Link href={`/san-pham/${product.slug}`} aria-label={`Xem ${product.name}`}><ProductVisual product={product}/></Link>{product.badge&&<span className="badge">{product.badge}</span>}<button className={`icon-btn wishlist ${liked?"is-active":""}`} aria-label={liked?"Bỏ yêu thích":"Thêm vào yêu thích"} aria-pressed={liked} onClick={()=>dispatch({type:"wishlist",productId:product.id})}><Heart size={18} fill={liked?"currentColor":"none"}/></button></div>
  <div className="product-info"><span className="eyebrow">{product.brand} · {product.sport}</span><Link className="product-name" href={`/san-pham/${product.slug}`}>{product.name}</Link><div className="rating"><Star size={14} fill="currentColor"/><b>{product.rating}</b><span>({product.reviews})</span></div><div className="price-row"><strong>{formatPrice(product.price)}</strong>{product.oldPrice&&<del>{formatPrice(product.oldPrice)}</del>}</div><button className="card-add" onClick={()=>dispatch({type:"add",productId:product.id,variantId:product.variants[0].id})}><Plus size={17}/>Thêm vào giỏ</button></div>
 </article>;
}
