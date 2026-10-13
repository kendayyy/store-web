export type Sport = "Cầu lông" | "Pickleball";
export type PriceTier = "Phổ thông" | "Trung cấp" | "Cao cấp";
export interface ProductVariant { id: string; label: string; stock: number }
export interface Product { id:string; slug:string; name:string; brand:string; sport:Sport; category:string; tier:PriceTier; price:number; oldPrice?:number; rating:number; reviews:number; badge?:string; description:string; specs:Record<string,string>; variants:ProductVariant[]; tone:string; featured?:boolean; bestseller?:boolean }
export interface Category { id:string; name:string; sport?:Sport; description:string }
export interface Brand { id:string; name:string }
export interface Promotion { id:string; title:string; subtitle:string }
export interface Article { id:string; category:string; title:string; excerpt:string; date:string }
export interface CartItem { productId:string; variantId:string; quantity:number }
export interface CheckoutForm { fullName:string; phone:string; email:string; address:string; city:string; note:string; shipping:"standard"|"express"; payment:"cod"|"bank" }
export interface OrderSummary { orderId:string; total:number; itemCount:number }
export interface ThemeConfig { brandName:string; accent:string; logoText:string }
export type ChatParticipant = "customer"|"staff"|"system";
export interface ProductRecommendation { productId:string; reason:string }
export interface ChatMessage { id:string; sender:ChatParticipant; text:string; createdAt:number; recommendations?:ProductRecommendation[] }
export interface ChatSession { messages:ChatMessage[]; isOpen:boolean }
export interface LeadRequest { name:string; phone:string; note:string }
