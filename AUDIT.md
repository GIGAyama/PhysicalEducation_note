# 体育ノート 点検記録（GIGA Standard v4）

2026-08-03 時点。数字はすべて**実測**。測っていないものは「未計測」と書く。

## 測り方（ここを先に読んでください）

本番（`script.google.com`）へは作業環境から到達できないため、
**GAS が返す画面と同じものを手元で組み立てて**測った。

- `index.html` + `css.html` + `js.html` を、GAS の `include()` と同じ順で貼り合わせる
- `google.script.run` は差し替えのダミーにし、戻り値の見本を与えて画面を進める
- `cdn.jsdelivr.net` はこの環境から出られないため、**npm から同じ版を取って**
  同じパスで配る。これをやらないと **Bootstrap が当たらない素の HTML** を測ることになり、
  数字が意味を失う（実際、控えを用意する前は「コントラスト6件」、あとは「2件」だった）

計測環境：Chromium 141（Playwright）、1280×900 / DPR 2。
道具は Digital_textbook の [`scripts/measure/`](https://github.com/GIGAyama/Digital_textbook/tree/main/scripts/measure) にある。

**測れていないもの**：サーバーの戻り値に強く依存する画面（ポートフォリオの中身、
教員用の集計表の実データ）。ダミーは見本しか返さないため、行数の多い表は再現していない。

---

## 1. まとめ

| 区分 | 前 | 後 |
|---|---|---|
| コントラスト基準未満（児童画面） | **6件** | **0件** |
| コントラスト基準未満（教員画面） | **2件** | **0件** |
| タップ44px未満 | 1件 | **0件** |
| `viewport-fit=cover` | 無し | 有り（`index.html` と `code.gs` の両方） |
| `env(safe-area-inset-*)` | 無し | 有り |
| `100dvh` | 無し | 有り |
| `clamp()` | 無し | 有り（見出し） |
| `prefers-reduced-motion` | 無し | 有り |
| `forced-colors` | 無し | 有り |
| CDN の SRI | 無し（版も未固定） | 4本すべてに付与・版を固定（改ざん検知を実証） |
| LICENSE / dependabot | 無し | 追加 |
| OAuth スコープ | `auth/drive`（全体） | **変更していない**（§5 に理由） |

---

## 2. いちばん効いた修正：ふりがなが青いボタンの上で読めなかった

`css.html` は `rt`（ふりがな）の色を `#666` に決め打ちしていた。
そのため**青いボタンの上に置いたふりがなが、濃い灰色のまま重なる**。

実測で **比 1.28**。ほぼ読めない。

しかも、ふりがなが必要なのは低学年の児童で、
**いちばん読めなくて困る人がいちばん読めない**という形になっていた。

直し方は、色を継がせるだけ。

```css
.btn rt, .badge rt, .nav-link rt, [class*="bg-primary"] rt { color: inherit; }
```

こうすると、どの面に置いても必ず読める。
`rt` の既定色そのものも `#666` → `#5f6368` に寄せた。

---

## 3. Bootstrap の既定色が基準に届いていない

このアプリ固有の配色ではなく、**Bootstrap 5.3 の既定色**が原因だった。
白地・14px での実測。

| クラス | 色 | 比 | 判定 |
|---|---|---:|---|
| `.text-primary` | `#0d6efd` | 4.27 | ❌ |
| `.text-danger` | `#dc3545` | 4.30 | ❌ |
| `.text-secondary` | `#6c757d` | 4.45 | ❌（あと 0.05） |
| `.btn-outline-info` | `#0dcaf0` | **1.96** | ❌ |
| `--primary-color`（このアプリ） | `#1a73e8` | 4.27 | ❌ |

`--primary-color` は白抜き文字を載せたときも 4.27 で、**表にも裏にも届いていなかった**。
Google Blue 700（`#1967d2`）へ1段濃くすると両方 5.0 になる。色の印象はほとんど変わらない。

Bootstrap 側は変数の上書きで直した。**この3つを使っている箇所すべてに一度で効く。**

```css
.text-primary { color: #0a58ca !important; }
.text-danger  { color: #b02a37 !important; }
.text-secondary { color: #5c636a !important; }
.btn-outline-info { --bs-btn-color: #087990; --bs-btn-border-color: #087990; }
```

> **他の Bootstrap を使うリポジトリでも、まったく同じ4件が出るはず。**
> 個別に直すより、この上書きを丸ごと持っていくほうが速い。

---

## 4. CDN に SRI を付け、版を固定した

4本すべて `cdn.jsdelivr.net` から、**SRI 無し**で読んでいた。
うち2本は**版も固定されていなかった**。

```html
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
```

版を固定していない URL には SRI を付けられない（中身が変わるため）。
そのうえ `chart.js` はメジャー版が上がると勝手に追随し、**ある日突然壊れうる**。

ファイルを明示して版を固定し、そのうえでハッシュを付けた。

| 資産 | 版 |
|---|---|
| bootstrap | 5.3.0（css / bundle.js） |
| bootstrap-icons | 1.11.3 |
| chart.js | 4.5.1（`dist/chart.umd.min.js`） |
| sweetalert2 | 11.26.25（`dist/sweetalert2.all.min.js`） |

ハッシュは**記憶で書かず、npm から取った同じ版の実バイトから計算した**。
間違ったハッシュを書くと、そのファイルは読み込まれず**アプリが起動しなくなる**。

### 効いていることを確かめた

「SRI を書いた」だけでは、効いているのか無視されているのか分からない。
ミラー上の `bootstrap.min.css` に**1バイト足して**測った。

| | 正常時 | 1バイト改ざん時 |
|---|---|---|
| Bootstrap のスタイルシート | 2枚 | **1枚**（弾かれた） |
| ボタンの余白 | 48px | **6px**（＝当たっていない） |
| コンソール | なし | `Failed to find a valid digest in the 'integrity' attribute` |

**改ざんされた資産は本当に読み込まれない。**

---

## 5. OAuth スコープ：変更していません（判断をお願いします）

`appsscript.json` に `https://www.googleapis.com/auth/drive`（**Drive 全体**）がある。
規格上は ❌ にあたる。

ただしコードを読むと、これは不注意ではなく**設計上の要求**だった。

```js
// code.gs
const files = DriveApp.searchFiles(q);   // 先生のドライブ全体からお手本を探す
```

`listMediaFiles()` は**先生のドライブ全体**から画像・スライド・動画を探して
「お手本」として選ばせる機能。`drive.file` に落とすと、
アプリが作ったファイルしか見えなくなり、**この機能が壊れる**。

### 候補

`https://www.googleapis.com/auth/drive.readonly` ＋
`https://www.googleapis.com/auth/drive.file` の組み合わせ。
全体を**読める**が、**書けるのはアプリが作ったものだけ**になる。
`DriveApp.createFolder()`（お手本フォルダの自動作成）は `drive.file` で足りる。

### なぜ変更しなかったか

**Apps Script の `DriveApp` はスコープの粒度が粗く、この組み合わせで
実際に通るかはデプロイして確かめないと分からない。**
外して間違えると、**全教員で認可が通らなくなり、体育の授業が止まる。**

作業環境から `script.google.com` へ到達できないため、確かめられない。
**確かめられないものを、教室が止まるリスクを取ってまで変えるべきではない。**

→ 次に本番へ触れる人が、テスト用のコピーで試してから入れ替えてください。

---

## 6. 表示（P1）

| 項目 | 対応 |
|---|---|
| `viewport-fit=cover` | `index.html` の `<meta>` と `code.gs` の `addMetaTag()` の**両方**に追加 |
| `env(safe-area-inset-*)` | `body` の左右下に適用 |
| `100dvh` | `.full-viewport` を用意し、`@supports` で古い端末に `100vh` を残した |
| `clamp()` | `h1`〜`h3` を画面幅から決まる式に |
| `prefers-reduced-motion` | 動きを止める（走る人のアニメーションが常時動いていた） |
| `forced-colors` | ハイコントラストで境目が消えないよう枠線を出す |
| `touch-action: manipulation` | 児童の連打で勝手に拡大するのを止める（二本指の拡大は残す） |
| フォント | Google Fonts が塞がれても崩れないよう、端末側の日本語フォントを後ろに並べた |

`viewport-fit` は **2か所直さないと効かない**。
GAS は画面を iframe で包むため、`index.html` の `<meta>` だけ直しても
外枠側が古いままで、安全領域が使えるようにならない。

`user-scalable` / `maximum-scale` は**指定していない**。
指定すると児童が拡大できなくなり、見えづらい子が使えなくなる。

---

## 7. 未対応・未計測

- **PWA 化していない**（manifest / Service Worker / offline.html なし）。
  C型は `script.google.com` 配信のため、GitHub Pages 側のシェル（C+型）を
  作らないと PWA にできない。これは構成の変更なので、別途判断が要る。
- サーバーの戻り値に強く依存する画面（ポートフォリオ、教員用の集計表の実データ）は未計測。
- 本番の URL・デプロイ状態は未確認（到達できないため）。

---

## 8. この PR をマージしてよいか

**`.gs` と `.html` は GAS 本体**であり、この作業環境から本番の挙動を確認できません。
表示まわりは手元で実測していますが、**サーバー側の動作確認は取れていません。**

そのため**自動ではマージしていません。** 内容を確認のうえ、
テスト用のデプロイで一度動かしてから取り込んでください。
