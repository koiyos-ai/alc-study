# alc-study　酒税法 学習リポジトリ

税務大学校「本科二次試験（酒税法）」の過去問（H30・R1・R3〜R7。R2 は試験なし）と、その解答・勉強用資料をまとめたリポジトリ。

## ファイル一覧

### 講本

| ファイル | 内容 |
|---|---|
| `all.pdf` | 税務大学校「間接税法（基礎編）令和8年度版」全体（国税庁ホームページの税大講本） |
| `shuzeiho-R8.pdf` | 上記から第2編「酒税法」部分だけを抜き出したもの（46ページ） |

### 問題（`mondai/`）

| ファイル | 内容 |
|---|---|
| `mondai/H30.md` | H30 第一問〜第四問、第五問の途中まで |
| `mondai/R1.md` | R1 第一問〜第四問、第五問の途中まで |
| `mondai/R3.md` | R3 第一問〜第四問、第五問の途中まで |
| `mondai/R4.md` | R4 全問（第一問〜第五問、別表・税率表を含む） |
| `mondai/R5.md` | R5 全問（第一問〜第五問、別表・税率表を含む） |
| `mondai/R6.md` | R6 第一問〜第四問、第五問の途中まで |
| `mondai/R7.md` | R7 第一問〜第四問、第五問の途中まで |

いずれも試験問題の写真から書き起こしたもの。

### 解答（`answers/`）

| ファイル | 内容 |
|---|---|
| `answers/H30_official.md` | H30 解答例（第一問〜第四問）の書き起こし |
| `answers/R1_official.md` | R1 解答例（第一問〜第四問）の書き起こし |
| `answers/R3_official.md` | R3 解答例（第一問〜第四問）の書き起こし |
| `answers/R4_official.md` | R4 公式解答例（全問）の書き起こし |
| `answers/R4.md` | R4 第一問〜第四問の自作解答。公式解答例と照合して修正済み |
| `answers/R5.md` | R5 第一問〜第四問の自作解答。第二問〜第四問は R1・R3 解答例と照合して修正済み |
| `answers/R6.md` | R6 第一問〜第四問の自作解答。第二問 問1 は R1 解答例と照合して修正済み |
| `answers/R7.md` | R7 第一問〜第四問の自作解答 |

自作解答は、講本・条文（e-Gov 法令検索）・国税庁「酒税法及び酒類行政関係法令等解釈通達」だけを根拠に作成した。R5〜R7 は公式解答例がないため、正解かどうかは確認できていない。

### 勉強用資料（`study/`、スマホ向け Word）

| ファイル | 内容 |
|---|---|
| `study/R4_shuzeiho_study.docx` | R4 の問題と公式解答例・学習ポイント（全問） |
| `study/shuzeiho_dai1.docx` | H30・R1・R3〜R7 の大問1（種類・品目の判定、35問）を答えの品目別に整理した問題集。判定の基準まとめ付き |
| `study/R8_yosou.md`・`study/R8_yosou.docx` | R8（令和8年度）の出題予想と対策 |
| `study/shuzeiho_ronjutsu.docx` | H30・R1・R3〜R7 の論述問題（第二問〜第四問、21問）をテーマ別に整理した問題集 |

### ツール（`tools/`）

| ファイル | 内容 |
|---|---|
| `tools/build_r4_study.js` | `study/R4_shuzeiho_study.docx` を生成するスクリプト |
| `tools/build_ronjutsu.js` | `study/shuzeiho_ronjutsu.docx` を生成するスクリプト |

実行例：`node tools/build_ronjutsu.js study/shuzeiho_ronjutsu.docx`（npm の `docx` パッケージが必要）
| `tools/extract_dai1.py` | `mondai/`・`answers/` から大問1を取り出して JSON にするスクリプト |
| `tools/build_dai1.js` | 上記 JSON から `study/shuzeiho_dai1.docx` を生成するスクリプト |

大問1の資料の作り直し：`python3 tools/extract_dai1.py dai1.json && node tools/build_dai1.js dai1.json study/shuzeiho_dai1.docx`
| `tools/build_md_mobile.js` | Markdown をスマホ版 Word にするスクリプト（`node tools/build_md_mobile.js study/R8_yosou.md study/R8_yosou.docx`） |
