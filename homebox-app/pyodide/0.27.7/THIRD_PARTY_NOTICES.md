# 同梱しているサードパーティ製ソフトウェア

ホームボックスの「この端末でExcelを作る」機能は、次のソフトウェアをこのフォルダーから読み込み、端末内で動かす。
入手元は Pyodide 0.27.7 の公式配布（cdn.jsdelivr.net/pyodide/v0.27.7/full/）。各ファイルの SHA-256 は `pyodide-lock.json` と一致することを確認済み。
配布物は改変していない。

| ソフトウェア | 版 | ライセンス | ファイル |
|---|---|---|---|
| Pyodide（CPython 3.12.7 の WebAssembly 版を含む） | 0.27.7 | MPL-2.0（CPython 部分は PSF License） | pyodide.js, pyodide.asm.js, pyodide.asm.wasm, python_stdlib.zip, pyodide-lock.json |
| lxml | 5.2.1 | BSD-3-Clause（libxml2/libxslt は MIT） | lxml-5.2.1-…wasm32.whl |
| Pillow | 10.2.0 | MIT-CMU（HPND） | pillow-10.2.0-…wasm32.whl |
| cryptography | 42.0.5 | Apache-2.0 または BSD-3-Clause | cryptography-42.0.5-…wasm32.whl |
| OpenSSL | 1.1.1w | OpenSSL License / SSLeay License | openssl-1.1.1w.zip |
| cffi | 1.17.1 | MIT | cffi-1.17.1-…wasm32.whl |
| pycparser | 2.22 | BSD-3-Clause | pycparser-2.22-py3-none-any.whl |
| six | 1.16.0 | MIT | six-1.16.0-py2.py3-none-any.whl |

各ライセンスの全文は、それぞれの wheel 内の `*.dist-info/` と、Pyodide のソース（github.com/pyodide/pyodide）にある。

## 注意

- OpenSSL 1.1.1 はサポートが終了した系列。ここでは、端末内で承認済み一覧の電子署名（Ed25519）を検証する用途にだけ使い、通信には使わない。Pyodide を新しい版に上げるときに、あわせて更新する。
- Pyodide の版を上げるときは、`pyodide/<新しい版>/` に置き、`engine-worker.js` と `sw.js` の版を同時に変える。古い版のキャッシュは自動では消えないので、`sw.js` の削除対象に加える。
