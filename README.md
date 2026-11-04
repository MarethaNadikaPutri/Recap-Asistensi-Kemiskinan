# Rekap Asistensi Kemiskinan PPPM

Aplikasi frontend berbasis HTML, CSS, dan JavaScript dengan Supabase sebagai backend.

## Struktur file
- `index.html` — halaman utama aplikasi.
- `app.js` — logika aplikasi.
- `style.css` — stylesheet.
- `config.js` — konfigurasi Supabase (URL dan publishable/anon key).

## Menjalankan secara lokal
Buka folder ini menggunakan VS Code, lalu jalankan `index.html` dengan Live Server atau server lokal.

## Deploy ke Netlify
Hubungkan repositori GitHub ini ke Netlify. Build command dikosongkan dan publish directory diisi `.` (root folder).

## Keamanan
`config.js` hanya boleh berisi Supabase publishable/anon key. Jangan pernah memasukkan `service_role` key atau kata sandi database ke file frontend atau repositori GitHub. Pastikan Row Level Security (RLS) dan policies Supabase aktif dan benar.
