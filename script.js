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
const newsletter=document.getElementById('newsletterForm');if(newsletter)newsletter.addEventListener('submit',e=>{e.preventDefault();e.target.querySelector('button').textContent='WELCOME TO ROUGE';e.target.reset()});
onRougeReady(()=>{products=ROUGE_PRODUCTS;renderProducts();renderCart();setTimeout(()=>document.getElementById('loader')?.classList.add('done'),700);});