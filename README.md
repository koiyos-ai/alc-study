# alc-study　酒税法 学習リポジトリ

税務大学校「本科二次試験（酒税法）」の過去問（R4〜R7）と、その解答・勉強用資料をまとめたリポジトリ。

## ファイル一覧

### 講本

| ファイル | 内容 |
|---|---|
| `all.pdf` | 税務大学校「間接税法（基礎編）令和8年度版」全体（国税庁ホームページの税大講本） |
| `shuzeiho-R8.pdf` | 上記から第2編「酒税法」部分だけを抜き出したもの（46ページ） |

### 問題（`mondai/`）

| ファイル | 内容 |
|---|---|
| `mondai/R4.md` | R4 全問（第一問〜第五問、別表・税率表を含む） |
| `mondai/R5.md` | R5 全問（第一問〜第五問、別表・税率表を含む） |
| `mondai/R6.md` | R6 第一問〜第四問、第五問の途中まで |
| `mondai/R7.md` | R7 第一問〜第四問、第五問の途中まで |

いずれも試験問題の写真から書き起こしたもの。

### 解答（`answers/`）

| ファイル | 内容 |
|---|---|
| `answers/R4_official.md` | R4 公式解答例（全問）の書き起こし |
| `answers/R4.md` | R4 第一問〜第四問の自作解答。公式解答例と照合して修正済み |
| `answers/R5.md` | R5 第一問〜第四問の自作解答 |
| `answers/R6.md` | R6 第一問〜第四問の自作解答 |
| `answers/R7.md` | R7 第一問〜第四問の自作解答 |

自作解答は、講本・条文（e-Gov 法令検索）・国税庁「酒税法及び酒類行政関係法令等解釈通達」だけを根拠に作成した。R5〜R7 は公式解答例がないため、正解かどうかは確認できていない。

### 勉強用資料（`study/`、スマホ向け Word）

| ファイル | 内容 |
|---|---|
| `study/R4_shuzeiho_study.docx` | R4 の問題と公式解答例・学習ポイント（全問） |
| `study/shuzeiho_ronjutsu.docx` | R4〜R7 の論述問題（第二問〜第四問、12問）をテーマ別に整理した問題集 |

### ツール（`tools/`）

| ファイル | 内容 |
|---|---|
| `tools/build_r4_study.js` | `study/R4_shuzeiho_study.docx` を生成するスクリプト |
| `tools/build_ronjutsu.js` | `study/shuzeiho_ronjutsu.docx` を生成するスクリプト |

実行例：`node tools/build_ronjutsu.js study/shuzeiho_ronjutsu.docx`（npm の `docx` パッケージが必要）
