# ひな形検査報告（承認候補）

> この報告は機械的な検査結果です。**これだけでは承認になりません。** 管理者がExcelで中身を確認し、approval_sign.py で署名したものだけが承認済みになります。
> セルの値・コメント本文・文書情報の値・外部リンクの参照先は、この報告には表示しません。Excelで直接確認してください。

| 項目 | 内容 |
|---|---|
| ファイル名 | synthetic_blank_template.xlsx |
| ファイル全体のSHA-256 | `0065aff3982c625af951215e94d01719d8e3c102720bbf57a068d42c8623c17d` |
| サイズ | 15,158 バイト |
| 様式の種類 | genba-ijou-sokuho-2026-10-17 |
| 変換ツールの変更範囲の版 | report-2026-10-17/engine-2.5.0 |
| 検査日時（UTC） | 2026-11-07T06:15:26Z |
| 検査ツール | 2.5.0-layout-preview |

## シート構成

| 順番 | シート名 | 表示状態 | 部品 | 変換ツールの扱い | 注意 |
|---|---|---|---|---|---|
| 1 | 報告書 | visible | xl/worksheets/sheet1.xml | 報告書として書き込む |  |
| 2 | 写真 | visible | xl/worksheets/sheet2.xml | 写真帳として書き込む |  |
| 3 | 補助リスト | visible | xl/worksheets/sheet3.xml | 変更しない（承認後はバイト一致を検査） |  |

## 人が確認すべきもの

| 確認項目 | 件数 | Excelでの確認場所 |
|---|---|---|
| コメント | 0 | 校閲 → コメントの表示 |
| 外部への参照 | 1 | データ → リンクの編集（参照先パスを確認） |
| 名前の定義 | 0 | 数式 → 名前の管理 |
| チェックボックス等 | 1 | 開発 → デザインモードで各コントロール |
| VML図形 | 1 | チェックボックス・コメント枠の正体 |
| EMF画像 | 1 | 各シートの図・ロゴ |
| 写真枠以外の画像 | 1 | 各シートの図・画像に過去の写真が無いか |

外部参照のある部品（参照先は表示しません）:
- xl/externalLinks/_rels/externalLink1.xml.rels: 1件（ネットワーク/ファイル（externalLinkPath））

文書情報（値は表示しません。変換時に空にする項目は出力側だけで空にします）:
- docProps/core.xml:creator: 値あり
- docProps/core.xml:lastModifiedBy: 値あり
- docProps/app.xml:Company: 値あり

## 部品一覧

| 部品 | 用途 | 変更区分 | 要確認 | サイズ | SHA-256（先頭16桁） |
|---|---|---|---|---|---|
| [Content_Types].xml | 部品の種類一覧 | 指定範囲のみ変更可 |  | 887 | `c8e4e85ef4ea7207` |
| _rels/.rels | 参照関係 | 変更禁止 |  | 296 | `2207473c15e949b2` |
| docProps/app.xml | アプリ情報 | 指定範囲のみ変更可 |  | 232 | `8d2cca18338d4de4` |
| docProps/core.xml | 文書情報 | 指定範囲のみ変更可 |  | 324 | `88b85017eabd1f65` |
| xl/_rels/workbook.xml.rels | ブックの参照 | 指定範囲のみ変更可 |  | 1,007 | `93d3e6f1d305b050` |
| xl/calcChain.xml | 計算順序 | 削除予定 |  | 163 | `6abb61480bddb99d` |
| xl/ctrlProps/ctrlProp1.xml | フォームコントロール（チェックボックス等） | 変更禁止 | 要確認の内部部品 | 164 | `c83c7b4409a81f23` |
| xl/drawings/_rels/drawing1.xml.rels | 図形の参照 | 変更禁止 |  | 291 | `484425fa27604290` |
| xl/drawings/_rels/drawing2.xml.rels | 写真帳の図形の参照 | 指定範囲のみ変更可 |  | 564 | `c7b29c0b1f134341` |
| xl/drawings/drawing1.xml | 図形配置 | 変更禁止 |  | 410 | `ed9ff6d8a0480ce4` |
| xl/drawings/drawing2.xml | 写真帳の図形配置 | 指定範囲のみ変更可 |  | 776 | `e79b48315ad58c96` |
| xl/drawings/vmlDrawing1.vml | VML図形（チェックボックス・コメント枠など） | 変更禁止 | VML図形 | 91 | `0f37daf06c1b8614` |
| xl/externalLinks/_rels/externalLink1.xml.rels | 外部ブックへのリンク | 変更禁止 | 要確認の内部部品 | 361 | `deabcfb99ac45757` |
| xl/externalLinks/externalLink1.xml | 外部ブックへのリンク | 変更禁止 | 要確認の内部部品 | 256 | `afdcf0cc80374d73` |
| xl/media/image1.jpeg | 写真枠の画像 | 指定範囲のみ変更可 |  | 819 | `9bc3bd5db08afe2f` |
| xl/media/image2.jpeg | 写真枠の画像 | 指定範囲のみ変更可 |  | 819 | `fb4634a67e9fff8b` |
| xl/media/image3.jpeg | 写真枠の画像 | 指定範囲のみ変更可 |  | 819 | `96b82ca7088eb83c` |
| xl/media/image4.emf | EMF画像（ロゴ・図など） | 変更禁止 | 写真枠以外の画像・EMF | 87 | `20228c7eb2a51b93` |
| xl/sharedStrings.xml | 共有文字列 | 指定範囲のみ変更可 |  | 698 | `11310e1607344a92` |
| xl/workbook.xml | ブック設定 | 指定範囲のみ変更可 |  | 413 | `7acdd2adf62a3906` |
| xl/worksheets/_rels/sheet1.xml.rels | ワークシートの参照 | 変更禁止 |  | 591 | `1d3f2ffdabb32e79` |
| xl/worksheets/_rels/sheet2.xml.rels | ワークシートの参照 | 変更禁止 |  | 298 | `8d64d6eb2217d154` |
| xl/worksheets/sheet1.xml | 報告書シート | 指定範囲のみ変更可 |  | 830 | `4ee7a9b01109ac4f` |
| xl/worksheets/sheet2.xml | 写真帳シート | 指定範囲のみ変更可 |  | 388 | `280a510f45756550` |
| xl/worksheets/sheet3.xml | ワークシート（変換ツールは触らない） | 変更禁止 | 変換ツールが触らない追加シート | 430 | `b7f32cca9e90a583` |

## 変更区分の意味

- **変更禁止**: 変換後もバイト単位で同一であることを検査します。
- **指定範囲のみ変更可**: 変換ツールのコード（engine_policy.py）に書かれた範囲だけ変わることを検査します。
  - xl/worksheets/sheet1.xml（報告書シート）: 入力欄として指定されたセルの値のみ。数式の計算済みの値は消去
  - xl/worksheets/sheet2.xml（写真帳シート）: 写真番号・場所・説明・加工指示のセルの値のみ。数式の計算済みの値は消去
  - xl/sharedStrings.xml（共有文字列）: 参照されなくなった文字列の中身を空にする（番号は変えない）
  - xl/drawings/drawing2.xml（写真帳の図形配置）: 旧加工指示の図形と未使用の写真枠を外す
  - xl/drawings/_rels/drawing2.xml.rels（写真帳の図形の参照）: 未使用の写真枠への参照を外す
  - xl/workbook.xml（ブック設定）: 再計算の指定（calcPr）のみ
  - xl/_rels/workbook.xml.rels（ブックの参照）: 計算順序（calcChain）への参照を外す。外部リンク等は保持
  - [Content_Types].xml（部品の種類一覧）: 計算順序（calcChain）の登録を外す
  - docProps/core.xml（文書情報）: 作成者・更新者・題名などを空にする
  - docProps/app.xml（アプリ情報）: 会社名・管理者を空にする
  - xl/media/image1.jpeg（写真枠の画像）: 新しい写真に差し替え。使わない枠は削除
  - xl/media/image2.jpeg（写真枠の画像）: 新しい写真に差し替え。使わない枠は削除
  - xl/media/image3.jpeg（写真枠の画像）: 新しい写真に差し替え。使わない枠は削除
- **削除予定**: 変換時に削除します。xl/calcChain.xml（数式を再計算させるため削除）
- **生成予定**: この版では新しく作る部品はありません。それ以外の部品が増えたら停止します。
