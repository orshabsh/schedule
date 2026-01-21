// Скопируйте этот файл как config.js и заполните значения из Supabase
// Project Settings → API → Project URL и anon/public key.
//
// ВАЖНО: anon key можно хранить в фронтенде. Доступ к данным ограничивается
// Row Level Security (RLS) политиками в Postgres.
// См. Supabase Docs: https://supabase.com/docs/guides/database/secure-data

window.SUPABASE_CONFIG = {
  url: "https://kdpadhpqnboqgftikolb.supabase.co",
  anonKey: "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImtkcGFkaHBxbmJvcWdmdGlrb2xiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NjkwMDUxOTQsImV4cCI6MjA4NDU4MTE5NH0.lu8Lzip28WKLsLt2Zn-GpufPpJiGeijpQ4iBIRMQv7k",
  // schoolDataId — первичный ключ строки в таблице school_data
  schoolDataId: 1,
};
