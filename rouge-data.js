/* ROUGE product/data layer — Google Sheets connected */
const ROUGE_API_URL = 'https://script.google.com/macros/s/AKfycbwLnGjb7qiKsGNmkfkxhQGZ7xFC4pus8Ur2NHi9pMM-vdo8MMReg1KfwQq78ranPuUe/exec';

// Kept as a fallback so the design still works if the API is temporarily unavailable.
const ROUGE_FALLBACK_PRODUCTS = [
{id:1,externalId:'R001',name:'White Silk Dress',category:'Dresses',price:2850,tag:'NEW IN',color:'Ivory',sizes:['XS','S','M','L','XL'],stock:10,image:'assets/products/1.svg',images:['assets/products/1.svg'],description:'A sculpted silk silhouette with an effortless drape, designed for evening light and unforgettable entrances.',material:'92% silk, 8% elastane',sku:'ROG-001'},
{id:2,externalId:'R002',name:'Rouge Evening Dress',category:'Evening',price:3200,tag:'SIGNATURE',color:'Rouge',sizes:['XS','S','M','L'],stock:8,image:'assets/products/2.svg',images:['assets/products/2.svg'],description:'A deep rouge evening silhouette cut to move with the body and catch the light.',material:'Satin viscose blend',sku:'ROG-002'},
{id:3,externalId:'R003',name:'Noir Tailored Blazer',category:'Clothing',price:2700,tag:'ESSENTIAL',color:'Noir',sizes:['XS','S','M','L','XL'],stock:12,image:'assets/products/3.svg',images:['assets/products/3.svg'],description:'Sharp tailoring softened with a feminine line. A modern wardrobe signature.',material:'Wool blend',sku:'ROG-003'},
{id:4,externalId:'R004',name:'Ivory Draped Top',category:'Clothing',price:1950,tag:'NEW IN',color:'Ivory',sizes:['XS','S','M','L'],stock:21,image:'assets/products/4.svg',images:['assets/products/4.svg'],description:'Fluid ivory fabric and a soft architectural drape for an elevated everyday silhouette.',material:'Silk-touch crepe',sku:'ROG-004'},
{id:5,externalId:'R005',name:'Rouge Satin Skirt',category:'Dresses',price:2350,tag:'EDITORIAL',color:'Rouge',sizes:['XS','S','M','L'],stock:13,image:'assets/products/5.svg',images:['assets/products/5.svg'],description:'Liquid satin in signature rouge, finished with a clean, elongated line.',material:'Recycled satin',sku:'ROG-005'},
{id:6,externalId:'R006',name:'Velvet Evening Set',category:'Evening',price:3900,tag:'LIMITED',color:'Bordeaux',sizes:['S','M','L'],stock:6,image:'assets/products/6.svg',images:['assets/products/6.svg'],description:'A rich velvet two-piece designed for after-dark occasions.',material:'Stretch velvet',sku:'ROG-006'},
{id:7,externalId:'R007',name:'Silk Rose Scarf',category:'Accessories',price:950,tag:'ACCESSORY',color:'Rose',sizes:['ONE SIZE'],stock:10,image:'assets/products/7.svg',images:['assets/products/7.svg'],description:'A delicate silk accent in the language of the ROUGE house.',material:'100% silk',sku:'ROG-007'},
{id:8,externalId:'R008',name:'Signature Rouge Bag',category:'Accessories',price:3100,tag:'SIGNATURE',color:'Rouge',sizes:['ONE SIZE'],stock:6,image:'assets/products/8.svg',images:['assets/products/8.svg'],description:'A structured rouge bag with a refined evening profile.',material:'Italian leather',sku:'ROG-008'}
];

let ROUGE_PRODUCTS = ROUGE_FALLBACK_PRODUCTS.map(x => ({...x}));
let ROUGE_PRODUCTS_READY = false;
let ROUGE_PRODUCTS_ERROR = null;

/**
 * General-purpose call to the ROUGE Apps Script API using a <script>-tag GET
 * request (JSONP), which avoids the CORS problems that plague cross-origin
 * fetch() POST requests to Apps Script Web Apps. Used for anything that reads
 * or writes data — e.g. rougeApiCall('placeOrder', {email, items, ...}).
 */
function rougeApiCall(action, params = {}) {
  return new Promise((resolve) => {
    const cb = '__rougeApi_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    let settled = false;
    const finish = (data) => {
      if (settled) return;
      settled = true;
      try { delete window[cb]; } catch (e) { window[cb] = undefined; }
      script.remove();
      resolve(data);
    };
    const timeout = setTimeout(() => finish({ success:false, error:'Request timed out — check your internet connection' }), 25000);
    window[cb] = (data) => { clearTimeout(timeout); finish(data); };
    script.onerror = () => { clearTimeout(timeout); finish({ success:false, error:'Could not reach the ROUGE API' }); };
    const qp = new URLSearchParams({ action, callback: cb, _: Date.now(), ...params });
    script.src = `${ROUGE_API_URL}?${qp.toString()}`;
    document.head.appendChild(script);
  });
}

const money = n => `EGP ${Number(n || 0).toLocaleString('en-EG')}`;
const cartKey='rougeCartV4', wishKey='rougeWishlistV4', ordersKey='rougeOrdersV4';

function onRougeReady(callback){
  if(ROUGE_PRODUCTS_READY) callback();
  else window.addEventListener('rougeProductsReady', callback, {once:true});
}

function getDriveId(value){
  const url = String(value || '').trim();
  if(!url) return '';
  const patterns = [
    /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/,
    /drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/,
    /drive\.google\.com\/uc\?(?:[^#]*?&)?id=([a-zA-Z0-9_-]+)/,
    /drive\.google\.com\/thumbnail\?id=([a-zA-Z0-9_-]+)/
  ];
  for(const re of patterns){ const m=url.match(re); if(m) return m[1]; }
  if(/^[a-zA-Z0-9_-]{20,}$/.test(url)) return url;
  return '';
}

function imageCandidates(value){
  const url = String(value || '').trim();
  if(!url) return [];
  const id = getDriveId(url);
  if(id){
    return [
      `https://drive.google.com/thumbnail?id=${id}&sz=w1600`,
      `https://drive.google.com/uc?export=view&id=${id}`,
      url
    ];
  }
  return [url];
}

function getDriveId(value){
  const url = String(value || '').trim();
  if(!url) return '';
  const patterns=[/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/,/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/,/drive\.google\.com\/uc\?(?:[^#]*?&)?id=([a-zA-Z0-9_-]+)/,/drive\.google\.com\/thumbnail\?id=([a-zA-Z0-9_-]+)/];
  for(const re of patterns){const m=url.match(re);if(m)return m[1];}
  if(/^[a-zA-Z0-9_-]{20,}$/.test(url)) return url;
  return '';
}
function imageCandidates(value){
  const url=String(value||'').trim(); if(!url)return [];
  const id=getDriveId(url);
  if(id)return [`https://drive.google.com/thumbnail?id=${id}&sz=w1600`,`https://drive.google.com/uc?export=view&id=${id}`,url];
  return [url];
}
function normalizeImageUrl(value){return imageCandidates(value)[0]||'';}
function imageFallbackAttrs(value){const c=imageCandidates(value).slice(1);return c.length?` data-image-fallbacks="${encodeURIComponent(JSON.stringify(c))}"`:'';}
function handleImageFallback(img){try{const list=JSON.parse(decodeURIComponent(img.dataset.imageFallbacks||'[]'));const next=list.shift();if(next){img.dataset.imageFallbacks=encodeURIComponent(JSON.stringify(list));img.src=next;}else img.classList.add('image-unavailable');}catch(e){img.classList.add('image-unavailable');}}
function escapeHtml(value){return String(value??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function productImageMarkup(p, extraClass=''){const src=p.image||`assets/products/${p.id}.svg`;return `<img class="${extraClass}" src="${src}" alt="${escapeHtml(p.name)}"${imageFallbackAttrs(p.originalImage||src)} onerror="handleImageFallback(this)">`;}
function pageProductImageMarkup(p,pageKey,extraClass=''){const o=p.pageOverrides?.[pageKey]||{};const src=o.image||p.image||`assets/products/${p.id}.svg`;const fallback=o.image||p.originalImage||src;return `<img class="${extraClass}" src="${src}" alt="${escapeHtml(o.title||p.name)}"${imageFallbackAttrs(fallback)} onerror="handleImageFallback(this)">`;}
function pageProductTitle(p,pageKey){return String((p.pageOverrides?.[pageKey]?.title||p.name||'')).trim();}
function pageProductSort(p,pageKey){return Number(p.pageOverrides?.[pageKey]?.sortOrder||0);}

function normalizeProduct(raw, index){
  const rawId = String(raw.id || raw.ID || '').trim();
  const numericId = Number((rawId.match(/\d+/) || [index + 1])[0]) || index + 1;
  const originalImages = Array.isArray(raw.images) ? raw.images.map(x=>String(x||'').trim()).filter(Boolean) : [];
  const images = originalImages.map(normalizeImageUrl).filter(Boolean);
  const image = images[0] || `assets/products/${numericId}.svg`;
  const sizes = Array.isArray(raw.sizes) && raw.sizes.length ? raw.sizes : ['ONE SIZE'];
  const totalStock = Number(raw.stock || 0);
  const stockMap = sizes.reduce((out, size) => { out[size] = totalStock; return out; }, {});
  return {
    ...raw,
    id: numericId,
    externalId: rawId || `R${String(numericId).padStart(3,'0')}`,
    name: String(raw.name || '').trim(),
    category: String(raw.category || '').trim(),
    price: Number(raw.price || 0),
    description: String(raw.description || '').trim(),
    color: String(raw.color || '').trim(),
    sizes,
    stock: totalStock,
    stockBySize: stockMap,
    originalImages: originalImages,
    images: images.length ? images : [image],
    image,
    tag: String(raw.tag || (raw.featured ? 'NEW IN' : 'ROUGE EDIT')).trim(),
    material: String(raw.material || 'Please see product details.').trim(),
    sku: String(raw.sku || rawId || `ROG-${String(numericId).padStart(3,'0')}`).trim(),
    featured: Boolean(raw.featured),
    status: String(raw.status || 'Active').trim()
  };
}

function loadRougeProducts(){
  return new Promise((resolve) => {
    const callbackName = '__rougeProducts_' + Date.now() + '_' + Math.random().toString(36).slice(2);
    const script = document.createElement('script');
    const cleanup = () => {
      try { delete window[callbackName]; } catch(e) { window[callbackName] = undefined; }
      script.remove();
    };

    const timeout = setTimeout(() => {
      cleanup();
      ROUGE_PRODUCTS_ERROR = new Error('ROUGE API timeout.');
      // Keep the demo fallback visible if the API cannot be reached.
      ROUGE_PRODUCTS_READY = true;
      window.dispatchEvent(new Event('rougeProductsReady'));
      resolve(ROUGE_PRODUCTS);
    }, 12000);

    window[callbackName] = (data) => {
      clearTimeout(timeout);
      cleanup();
      try {
        if (!data || !data.success) throw new Error(data?.error || 'Product API failed');
        const incoming = Array.isArray(data.products) ? data.products.map((p,i)=>normalizeProduct(p,i)).filter(p=>p.name) : [];
        if (incoming.length) {
          ROUGE_PRODUCTS = incoming;
          ROUGE_PRODUCTS_ERROR = null;
        } else {
          ROUGE_PRODUCTS_ERROR = new Error('The API returned no active products.');
          // Do NOT erase the fallback catalog when API returns no products.
        }
      } catch(error) {
        ROUGE_PRODUCTS_ERROR = error;
      } finally {
        ROUGE_PRODUCTS_READY = true;
        window.dispatchEvent(new Event('rougeProductsReady'));
        resolve(ROUGE_PRODUCTS);
      }
    };

    script.onerror = () => {
      clearTimeout(timeout);
      cleanup();
      ROUGE_PRODUCTS_ERROR = new Error('Could not load the ROUGE product API.');
      ROUGE_PRODUCTS_READY = true;
      window.dispatchEvent(new Event('rougeProductsReady'));
      resolve(ROUGE_PRODUCTS);
    };

    script.src = `${ROUGE_API_URL}?action=products&callback=${encodeURIComponent(callbackName)}&_=${Date.now()}`;
    document.head.appendChild(script);
  });
}
loadRougeProducts();

function getCart(){try{return JSON.parse(localStorage.getItem(cartKey)||'[]')}catch{return[]}}
function setCart(v){localStorage.setItem(cartKey,JSON.stringify(v));updateBagCount();renderBag?.()}
function getWishlist(){try{return JSON.parse(localStorage.getItem(wishKey)||'[]')}catch{return[]}}
function setWishlist(v){localStorage.setItem(wishKey,JSON.stringify(v));updateWishCount()}
function updateBagCount(){document.querySelectorAll('[data-bag-count],#bagCount').forEach(e=>e.textContent=getCart().reduce((s,x)=>s+x.qty,0))}
function updateWishCount(){document.querySelectorAll('[data-wish-count]').forEach(e=>e.textContent=getWishlist().length)}
function findProduct(id){return ROUGE_PRODUCTS.find(x=>String(x.id)===String(id)||String(x.externalId).toLowerCase()===String(id).toLowerCase())}
function addProduct(id,size=null){const p=findProduct(id);if(!p)return;size=size||p.sizes.find(s=>(p.stockBySize?.[s] ?? p.stock)>0)||p.sizes[0];const available=p.stockBySize?.[size] ?? p.stock ?? 99;const c=getCart();const found=c.find(x=>String(x.id)===String(p.id)&&x.size===size);if(found)found.qty=Math.min(found.qty+1,available||99);else c.push({id:p.id,externalId:p.externalId,size,qty:1});setCart(c);openBag?.()}
function changeQty(id,size,delta){const c=getCart();const x=c.find(i=>String(i.id)===String(id)&&i.size===size);if(!x)return;const p=findProduct(id);const max=p?.stockBySize?.[size] ?? p?.stock ?? 99;x.qty=Math.min(x.qty+delta,max||99);if(x.qty<=0)c.splice(c.indexOf(x),1);setCart(c)}
function removeCartItem(id,size){setCart(getCart().filter(x=>!(String(x.id)===String(id)&&x.size===size)))}
function toggleWish(id){const numeric=Number(findProduct(id)?.id ?? id);const w=getWishlist();const i=w.indexOf(numeric);i>=0?w.splice(i,1):w.push(numeric);setWishlist(w);document.querySelectorAll(`[data-wish="${numeric}"]`).forEach(b=>{b.classList.toggle('is-wished',w.includes(numeric));b.classList.remove('wish-pulse');void b.offsetWidth;b.classList.add('wish-pulse');setTimeout(()=>b.classList.remove('wish-pulse'),700)})}
function isWished(id){return getWishlist().includes(Number(findProduct(id)?.id ?? id))}
function openBag(){document.querySelector('#bagDrawer,#cartDrawer')?.classList.add('is-open','open');document.querySelector('#bagBackdrop,#backdrop')?.classList.add('is-open','show');renderBag?.()}
function closeBag(){document.querySelector('#bagDrawer,#cartDrawer')?.classList.remove('is-open','open');document.querySelector('#bagBackdrop,#backdrop')?.classList.remove('is-open','show')}
function renderBag(){const box=document.querySelector('#bagItems,#cartItems'),totalEl=document.querySelector('#bagTotal,#cartTotal');if(!box||!totalEl)return;const c=getCart();if(!c.length)box.innerHTML='<p class="bag-empty empty">Your bag is empty.</p>';else box.innerHTML=c.map(x=>{const p=findProduct(x.id);if(!p)return '';return `<div class="bag-row cart-row"><img src="${p.image}" alt="${p.name}"><div><strong>${p.name}</strong><small>SIZE ${x.size}</small><small>${money(p.price)}</small><div class="qty"><button onclick="changeQty(${p.id},'${x.size}',-1)">−</button><b>${x.qty}</b><button onclick="changeQty(${p.id},'${x.size}',1)">+</button></div></div><button class="remove" onclick="removeCartItem(${p.id},'${x.size}')">×</button></div>`}).join('');totalEl.textContent=money(c.reduce((s,x)=>s+(findProduct(x.id)?.price||0)*x.qty,0))}
function initShared(){
  if(window.__rougeSharedInitialized) {
    updateBagCount();
    updateWishCount();
    renderBag?.();
    document.querySelectorAll('[data-wish]').forEach(b=>b.classList.toggle('is-wished',isWished(b.dataset.wish)));
    return;
  }
  window.__rougeSharedInitialized = true;

  updateBagCount();
  updateWishCount();
  renderBag?.();

  // Delegated clicks: works for products rendered after the page loads too.
  document.addEventListener('click', (e) => {
    const wish = e.target.closest('[data-wish]');
    if(wish){
      e.preventDefault();
      e.stopPropagation();
      toggleWish(wish.dataset.wish);
      return;
    }

    const add = e.target.closest('[data-add]');
    if(add){
      e.preventDefault();
      e.stopPropagation();
      if(add.classList.contains('adding')) return;
      add.classList.add('adding');
      const originalLabel = add.getAttribute('aria-label');
      add.setAttribute('aria-label','Added to bag');
      addProduct(add.dataset.add, add.dataset.size || null);
      setTimeout(()=>{
        add.classList.remove('adding');
        add.setAttribute('aria-label', originalLabel || 'Add to bag');
      }, 850);
      return;
    }
  }, true);

  document.querySelector('#bagButton,#bagBtn')?.addEventListener('click',openBag);
  document.querySelector('#bagClose,#closeBag')?.addEventListener('click',closeBag);
  document.querySelector('#bagBackdrop,#backdrop')?.addEventListener('click',closeBag);
  document.querySelector('#searchButton,#searchBtn')?.addEventListener('click',()=>document.querySelector('#searchPanel,#searchModal')?.classList.add('is-open','open'));
  document.querySelector('#searchClose,#closeSearch')?.addEventListener('click',()=>document.querySelector('#searchPanel,#searchModal')?.classList.remove('is-open','open'));
  document.querySelector('#menuButton,#menuBtn')?.addEventListener('click',()=>document.querySelector('#mobileNav')?.classList.toggle('is-open'));
  window.addEventListener('scroll',()=>document.querySelector('#siteNav,#nav')?.classList.toggle('scrolled',scrollY>40));
  window.addEventListener('mousemove',e=>{const g=document.querySelector('.cursor-glow');if(g)g.style.transform=`translate(${e.clientX}px,${e.clientY}px)`});
  const mobileNav=document.querySelector('#mobileNav');
  if(mobileNav) mobileNav.classList.remove('is-open');
  document.addEventListener('keydown',e=>{if(e.key==='Escape'){mobileNav?.classList.remove('is-open');document.querySelector('#bagDrawer,#cartDrawer')?.classList.remove('is-open','open');document.querySelector('#bagBackdrop,#backdrop')?.classList.remove('is-open','show');document.querySelector('#searchPanel,#searchModal')?.classList.remove('is-open','open')}});
  setTimeout(()=>document.querySelector('#pageLoader,#loader')?.classList.add('done'),650);
}
document.addEventListener('DOMContentLoaded',()=>{if(ROUGE_PRODUCTS_READY)initShared();else onRougeReady(initShared)});
