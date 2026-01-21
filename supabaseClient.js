/*
  Supabase client + helpers.
  - Публичный просмотр расписания работает под ролью anon.
  - Редактирование доступно только после входа (Auth) и при наличии роли в profiles.

  Требуется config.js рядом с этим файлом (см. config.example.js).
*/

// supabase-js v2 подключается в HTML через CDN и экспортирует глобальный объект `supabase`.

function getSbConfig() {
  const cfg = window.SUPABASE_CONFIG;
  if (!cfg?.url || !cfg?.anonKey) return null;
  return cfg;
}

let sb = null;

function getSupabase() {
  if (sb) return sb;
  const cfg = getSbConfig();
  if (!cfg) return null;
  // eslint-disable-next-line no-undef
  sb = supabase.createClient(cfg.url, cfg.anonKey);
  return sb;
}

async function sbGetSession() {
  const client = getSupabase();
  if (!client) return { session: null, user: null };
  const { data, error } = await client.auth.getSession();
  if (error) return { session: null, user: null };
  return { session: data.session, user: data.session?.user ?? null };
}

async function sbSignIn(email, password) {
  const client = getSupabase();
  if (!client) throw new Error("Supabase not configured");
  return client.auth.signInWithPassword({ email, password });
}

async function sbSignOut() {
  const client = getSupabase();
  if (!client) return;
  await client.auth.signOut();
}

async function sbGetMyProfile() {
  const client = getSupabase();
  if (!client) return null;
  const { session } = await sbGetSession();
  if (!session?.user) return null;

  const { data, error } = await client
    .from("profiles")
    .select("id, role, display_name")
    .eq("id", session.user.id)
    .maybeSingle();

  if (error) return null;
  return data;
}

async function sbLoadSchoolData() {
  const client = getSupabase();
  if (!client) throw new Error("Supabase not configured");
  const cfg = getSbConfig();

  const { data, error } = await client
    .from("school_data")
    .select("id, data")
    .eq("id", cfg.schoolDataId || 1)
    .maybeSingle();

  if (error) throw error;
  return data?.data ?? null;
}

async function sbSaveSchoolData(nextData) {
  const client = getSupabase();
  if (!client) throw new Error("Supabase not configured");
  const cfg = getSbConfig();

  // update — должен пройти через RLS (только для admin/deputy)
  const { error } = await client
    .from("school_data")
    .update({ data: nextData, updated_at: new Date().toISOString() })
    .eq("id", cfg.schoolDataId || 1);

  if (error) throw error;
}

async function sbUpdateMyPassword(newPassword) {
  const client = getSupabase();
  if (!client) throw new Error("Supabase not configured");
  return client.auth.updateUser({ password: newPassword });
}

// Экспортируем в window (чтобы app.js мог использовать без сборки)
window.sbApi = {
  getSupabase,
  getSbConfig,
  sbGetSession,
  sbSignIn,
  sbSignOut,
  sbGetMyProfile,
  sbLoadSchoolData,
  sbSaveSchoolData,
  sbUpdateMyPassword,
};
