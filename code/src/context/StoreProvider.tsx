"use client";
import { createContext, useContext, useEffect, useMemo, useReducer } from "react";
import type { CartItem, ChatMessage, ThemeConfig } from "@/types";
type State={cart:CartItem[];wishlist:string[];theme:ThemeConfig;messages:ChatMessage[];hydrated:boolean};
type Action={type:"add";productId:string;variantId:string}|{type:"quantity";productId:string;quantity:number}|{type:"remove";productId:string}|{type:"clearCart"}|{type:"wishlist";productId:string}|{type:"theme";value:Partial<ThemeConfig>}|{type:"message";value:ChatMessage}|{type:"clearChat"}|{type:"hydrate";value:State}|{type:"ready"};
const initial:State={cart:[],wishlist:[],theme:{brandName:"[TÊN THƯƠNG HIỆU]",accent:"#F4511E",logoText:"TS"},messages:[{id:"welcome",sender:"staff",text:"Xin chào, tôi là Minh từ đội tư vấn. Bạn đang tìm đồ cầu lông hay pickleball?",createdAt:1}],hydrated:false};
function reducer(state:State,action:Action):State{
 if(action.type==="hydrate")return{...action.value,hydrated:true};
 if(action.type==="ready")return{...state,hydrated:true};
 if(action.type==="add"){const found=state.cart.find(i=>i.productId===action.productId&&i.variantId===action.variantId);return{...state,cart:found?state.cart.map(i=>i===found?{...i,quantity:i.quantity+1}:i):[...state.cart,{productId:action.productId,variantId:action.variantId,quantity:1}]};}
 if(action.type==="quantity")return{...state,cart:state.cart.map(i=>i.productId===action.productId?{...i,quantity:Math.max(1,action.quantity)}:i)};
 if(action.type==="remove")return{...state,cart:state.cart.filter(i=>i.productId!==action.productId)};
 if(action.type==="clearCart")return{...state,cart:[]};
 if(action.type==="wishlist")return{...state,wishlist:state.wishlist.includes(action.productId)?state.wishlist.filter(id=>id!==action.productId):[...state.wishlist,action.productId]};
 if(action.type==="theme")return{...state,theme:{...state.theme,...action.value}};
 if(action.type==="message")return{...state,messages:[...state.messages,action.value]};
 if(action.type==="clearChat")return{...state,messages:initial.messages};return state;
}
type Store=State&{dispatch:React.Dispatch<Action>;cartCount:number};const StoreContext=createContext<Store|null>(null);
export function StoreProvider({children}:{children:React.ReactNode}){const[state,dispatch]=useReducer(reducer,initial);useEffect(()=>{try{const saved=localStorage.getItem("sports-store-state");if(saved)dispatch({type:"hydrate",value:{...initial,...JSON.parse(saved)}});else dispatch({type:"ready"});}catch{dispatch({type:"ready"})}},[]);useEffect(()=>{if(!state.hydrated)return;localStorage.setItem("sports-store-state",JSON.stringify(state));document.documentElement.style.setProperty("--accent",state.theme.accent);},[state]);const value=useMemo(()=>({...state,dispatch,cartCount:state.cart.reduce((s,i)=>s+i.quantity,0)}),[state]);return <StoreContext.Provider value={value}>{children}</StoreContext.Provider>}
export function useStore(){const value=useContext(StoreContext);if(!value)throw new Error("useStore must be used in StoreProvider");return value}
