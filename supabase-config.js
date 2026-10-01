/* ROUGE Supabase connection */
const SUPABASE_URL = 'https://jxafszvvccbtqwlzrymy.supabase.co';
const SUPABASE_ANON_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Imp4YWZzenZ2Y2NidHF3bHpyeW15Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTA4MDU1OTgsImV4cCI6MjEwNjM4MTU5OH0.NY67VfGDLfCbntXmSOMgd-N_sHYKemb1TIgN4Pcmhyo';
const SUPABASE_REST_URL = `${SUPABASE_URL}/rest/v1`;

function sbHeaders(extra = {}) {
  return { apikey: SUPABASE_ANON_KEY, 'Content-Type': 'application/json', ...extra };
}

async function sbFetch(path, options = {}) {
  const response = await fetch(`${SUPABASE_REST_URL}${path}`, {
    ...options,
    headers: sbHeaders(options.headers || {})
  });
  const text = await response.text();
  let data = null;
  try { data = text ? JSON.parse(text) : null; } catch { data = text; }
  if (!response.ok) {
    const message = data?.message || data?.error_description || data?.hint || data?.details || 'Supabase request failed';
    throw new Error(message);
  }
  return data;
}

function sbQuery(table, query = '') { return sbFetch(`/${table}${query}`); }
function sbInsert(table, rows, options = {}) {
  return sbFetch(`/${table}`, { method:'POST', headers:{ Prefer: options.select ? 'return=representation' : 'return=minimal' }, body:JSON.stringify(rows) });
}
function sbUpsert(table, rows, onConflict, options = {}) {
  const q = onConflict ? `?on_conflict=${encodeURIComponent(onConflict)}` : '';
  return sbFetch(`/${table}${q}`, { method:'POST', headers:{ Prefer: options.select ? 'return=representation,resolution=merge-duplicates' : 'resolution=merge-duplicates,return=minimal' }, body:JSON.stringify(rows) });
}
function sbPatch(table, query, values, options = {}) {
  return sbFetch(`/${table}${query}`, { method:'PATCH', headers:{ Prefer: options.select ? 'return=representation' : 'return=minimal' }, body:JSON.stringify(values) });
}
function sbDelete(table, query) { return sbFetch(`/${table}${query}`, { method:'DELETE', headers:{ Prefer:'return=minimal' } }); }
function sbRpc(fn, args = {}) { return sbFetch(`/rpc/${fn}`, { method:'POST', body:JSON.stringify(args) }); }

function sbFetchWithToken(path, token, options = {}) {
  return sbFetch(path, { ...options, headers: { ...(options.headers || {}), Authorization: `Bearer ${token}` } });
}
