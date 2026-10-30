電気の疑問図鑑 — ホーム画面アプリ（PWA）一式

このフォルダの中身をそのまま HTTPS のサイト（GitHub Pages、Netlify、Cloudflare Pages など）に置くと、
ホーム画面に追加できる単独アプリになります。一度開くと全データがキャッシュされ、機内モードでも起動します。

追加方法
- iPhone / iPad（Safari）：共有ボタン →「ホーム画面に追加」
- Android（Chrome）：メニュー →「ホーム画面に追加」または「アプリをインストール」

更新するときは index.html を差し替え、sw.js の CACHE 名（denki-zukan-v2.2）を新しい値に変えてください。
