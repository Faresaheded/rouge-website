-- ROUGE checkout RPC: validates products, stock, sizes and creates the order atomically.
create sequence if not exists public.rouge_order_seq start 10004;

create or replace function public.place_order(
  p_email text,
  p_phone text,
  p_full_name text,
  p_country text,
  p_city text,
  p_postal_code text,
  p_address text,
  p_payment text,
  p_items jsonb,
  p_order_token text
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  raw jsonb;
  product_row public.products%rowtype;
  qty integer;
  size_value text;
  line_items jsonb := '[]'::jsonb;
  subtotal numeric(12,2) := 0;
  shipping numeric(12,2) := 0;
  total numeric(12,2);
  order_no text;
  existing_order public.orders%rowtype;
  allowed_sizes text[];
begin
  if coalesce(trim(p_email),'') = '' or coalesce(trim(p_full_name),'') = '' or coalesce(trim(p_address),'') = '' or coalesce(trim(p_city),'') = '' or coalesce(trim(p_phone),'') = '' or coalesce(trim(p_country),'') = '' then
    raise exception 'Please complete all required delivery details.';
  end if;
  if lower(trim(p_payment)) <> 'cash on delivery' then
    raise exception 'This checkout currently supports Cash on Delivery only.';
  end if;
  if coalesce(trim(p_order_token),'') = '' then
    raise exception 'Missing order token. Please refresh checkout and try again.';
  end if;
  if jsonb_typeof(p_items) <> 'array' or jsonb_array_length(p_items) = 0 then
    raise exception 'Your bag is empty.';
  end if;

  select * into existing_order from public.orders where order_token = trim(p_order_token) limit 1;
  if found then
    return jsonb_build_object('success',true,'orderNumber',existing_order.order_number,'date',existing_order.order_date);
  end if;

  for raw in select value from jsonb_array_elements(p_items)
  loop
    qty := floor(coalesce((raw->>'qty')::numeric,0));
    size_value := trim(coalesce(raw->>'size',''));
    select * into product_row from public.products where id = trim(coalesce(raw->>'id','')) and lower(status)='active' for update;
    if not found then raise exception 'One of the selected products is no longer available.'; end if;
    if qty < 1 or qty > 50 then raise exception 'Invalid quantity for %.', product_row.name; end if;
    if product_row.stock < qty then raise exception '% is no longer available in the requested quantity.', product_row.name; end if;
    allowed_sizes := regexp_split_to_array(regexp_replace(coalesce(product_row.sizes,''),'\s*,\s*',',','g'), ',');
    if coalesce(array_length(allowed_sizes,1),0) > 0 and size_value <> '' and not (size_value = any(allowed_sizes)) then
      raise exception '% does not have size %.', product_row.name, size_value;
    end if;
    line_items := line_items || jsonb_build_array(jsonb_build_object('id',product_row.id,'name',product_row.name,'size',coalesce(nullif(size_value,''),allowed_sizes[1],'ONE SIZE'),'qty',qty,'price',product_row.price));
    subtotal := subtotal + product_row.price * qty;
    update public.products set stock = stock - qty, updated_at = now() where id = product_row.id;
  end loop;

  total := subtotal + shipping;
  order_no := 'ROG-' || nextval('public.rouge_order_seq');
  insert into public.orders(order_number,order_date,email,phone,full_name,country,city,postal_code,address,payment,items,subtotal,shipping,total,status,order_token,payment_method)
  values(order_no,now(),trim(p_email),trim(p_phone),trim(p_full_name),trim(p_country),trim(p_city),trim(coalesce(p_postal_code,'')),trim(p_address),trim(p_payment),line_items,subtotal,shipping,total,'Pending',trim(p_order_token),trim(p_payment));

  return jsonb_build_object('success',true,'orderNumber',order_no,'date',now(),'total',total);
exception when unique_violation then
  select * into existing_order from public.orders where order_token = trim(p_order_token) limit 1;
  if found then return jsonb_build_object('success',true,'orderNumber',existing_order.order_number,'date',existing_order.order_date); end if;
  raise;
end;
$$;

grant execute on function public.place_order(text,text,text,text,text,text,text,text,jsonb,text) to anon, authenticated;
