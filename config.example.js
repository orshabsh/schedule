// Скопируйте этот файл как config.js и заполните значения из Supabase
// Project Settings → API → Project URL и anon/public key.
//
// ВАЖНО: anon key можно хранить в фронтенде. Доступ к данным ограничивается
// Row Level Security (RLS) политиками в Postgres.
// См. Supabase Docs: https://supabase.com/docs/guides/database/secure-data

window.SUPABASE_CONFIG = {
  url: "https://YOUR_PROJECT_ID.supabase.co",
  anonKey: "YOUR_SUPABASE_ANON_KEY",
  // schoolDataId — первичный ключ строки в таблице school_data
  schoolDataId: 1,
};
