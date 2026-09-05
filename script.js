let products=[];let cart=getCart();const grid=document.getElementById('productGrid');const money2=money;
function renderProducts(){grid.innerHTML=products.map(p=>`<article class="product-card"><a href="product.html?id=${p.id}"><div class="product-image">${productImageMarkup(p)}<button class="quick-add" data-id="${p.id}" aria-label="Add ${p.name}">+</button></div><div class="product-info"><div><div class="product-name">${p.name}</div><div class="product-meta">${p.category} · ${p.tag}</div></div><div class="product-price">${money2(p.price)}</div></div></a></article>`).join('');document.querySelectorAll('.quick-add').forEach(b=>b.addEventListener('click',e=>{
e.preventDefault();e.stopPropagation();
if(b.classList.contains('adding'))return;
b.classList.add('adding');
addToCart(Number(b.dataset.id));
setTimeout(()=>b.classList.remove('adding'),850);
}))}
function addToCart(id){addProduct(id)}
function renderCart(){renderBag()}
const drawer=document.getElementById('cartDrawer'),backdrop=document.getElementById('backdrop');function openCart(){openBag()}function closeCart(){closeBag()}
document.getElementById('bagBtn').onclick=openCart;document.getElementById('closeBag').onclick=closeCart;backdrop.onclick=closeCart;
const search=document.getElementById('searchModal');document.getElementById('searchBtn').onclick=()=>{search.classList.add('open');setTimeout(()=>document.getElementById('searchInput').focus(),300)};document.getElementById('closeSearch').onclick=()=>search.classList.remove('open');
document.querySelectorAll('[data-shop-now]').forEach(b=>b.addEventListener('click',()=>{window.location.href=b.dataset.shopNow;}));
window.addEventListener('scroll',()=>document.getElementById('nav').classList.toggle('scrolled',scrollY>40));window.addEventListener('mousemove',e=>{const g=document.querySelector('.cursor-glow');if(g)g.style.transform=`translate(${e.clientX}px,${e.clientY}px)`});
const newsletter=document.getElementById('newsletterForm');
if(newsletter) newsletter.addEventListener('submit',async e=>{
 e.preventDefault(); const input=newsletter.querySelector('input'); const btn=newsletter.querySelector('button'); const email=input.value.trim(); const original=btn.textContent; btn.disabled=true; btn.textContent='JOINING…';
 try{const result=await new Promise(resolve=>{const cb='__rougeNews_'+Date.now();const sc=document.createElement('script');let done=false;const finish=d=>{if(done)return;done=true;try{delete window[cb]}catch(e){}sc.remove();resolve(d)};const t=setTimeout(()=>finish({success:false,error:'Request timed out.'}),12000);window[cb]=d=>{clearTimeout(t);finish(d)};sc.onerror=()=>{clearTimeout(t);finish({success:false,error:'Could not connect.'})};sc.src=`${ROUGE_API_URL}?${new URLSearchParams({action:'subscribeNewsletter',email,source:location.pathname,callback:cb,_:Date.now()}).toString()}`;document.head.appendChild(sc)});if(!result.success)throw new Error(result.error||'Could not subscribe.');btn.textContent=result.alreadySubscribed?'ALREADY ON THE LIST':'WELCOME TO ROUGE';input.value='';}catch(err){btn.textContent='TRY AGAIN';console.warn(err)}finally{setTimeout(()=>{btn.disabled=false;btn.textContent=original},2600)}});
onRougeReady(()=>{products=ROUGE_PRODUCTS;renderProducts();renderCart();setTimeout(()=>document.getElementById('loader')?.classList.add('done'),700);});