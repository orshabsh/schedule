// Скопируйте этот файл как config.js и заполните значения из Supabase
// Project Settings → API → Project URL и anon/public key.
//
// ВАЖНО: anon key можно хранить в фронтенде. Доступ к данным ограничивается
// Row Level Security (RLS) политиками в Postgres.
// См. Supabase Docs: https://supabase.com/docs/guides/database/secure-data

window.SUPABASE_CONFIG = {
  url: "https://kdpadhpqnboqgftikolb.supabase.co",
  anonKey: "sb_publishable__igllkiRCrYi51dskBnv3A_fBTnzkai",
  // schoolDataId — первичный ключ строки в таблице school_data
  schoolDataId: 1,
};
