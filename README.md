# asloop-survey

asloopが実施しているAro/Aceに関するアンケート結果を集計・可視化するためのスクリプト

## 構成

| ディレクトリ | 内容 |
| --- | --- |
| `config/` | 設問、選択肢、矛盾解答検知などの設定 |
| `data/` | アンケート結果データ、およびそれを含むアフターコーディング情報 |
| `gas/` | Google Apps Scriptで実行できるスクリプト |
| `*.ipynb` | Jupyter Notebookで実行できるスクリプト |

※ 各ファイルの詳細については後述します。

## 手順

TBW

## ipynb (Jupyter Notebook)

集計・グラフ化のためのスクリプト群

### スクリプト一覧

NOTE: 列名 (CとかANとか) がハードコーディングされているので注意。例えば年齢やorientation、居住地などの列名指定を修正する必要あり。
NOTE: 各スクリプトの設定と出力は後述。

- `setup.ipynb`: 設定を生成するためのスクリプト。
- `after_coding.ipynb`: アフターコーディングを行うためのスクリプト。
- `plots.ipynb`: アフターコーティング結果から単純集計結果を出力するためのスクリプト。現状では、出力したい内容に応じて修正が必要。
- `tabulation.ipynb`: アフターコーティング結果から調査結果報告書を出力するためのスクリプト。 `config/report_config.csv` に応じて出力。

## config

出力のための設定ファイル群

- `config/header.csv`
  - `setup.ipynb` で `questions.csv` を生成するために必要な、回答一覧シートのヘッダ行。
  - 回答一覧のシートから1行目をコピーしてください。
- `config/questions.csv`
  - 質問文の一覧と、それらのExcelのカラム名へのマッピング。
  - 回答結果のシートのヘッダ行をheader.csvに保存して、setup.ipynbを実行することで自動作成できます。
- `config/choices.csv`
  - 選択肢の一覧など、回答フォームから収集できる情報。
  - `gas/choices.gs` で生成されたファイルをCSVとしてダウンロードして設置してください。
- `config/duplicated.csv`
  - 自由記述が全く同じものを重複として検出するための、列の指定。
  - 質問票であらかじめ用意された選択肢に該当しないもの (その他) も自由記述として検出対象としています。
  - `out/duplicates.csv` に検出結果が出力されます。
  - 手動作成してください。
- `config/exclusive_choices.csv`
  - 矛盾回答の検出のための設定。
  - 「Aに当てはまる」と「どれにも当てはまらない」が同時に選択されていた場合に、どちらを優先するかを指定します。
  - 手動作成してください。
  - `out/exclusives.csv` に検出結果が出力されます。
- `config/age_contradictions.csv`
  - 年齢の矛盾回答の検出のための設定。
  - `out/age_contradictions.csv` に検出結果が出力されます。
  - 手動作成してください。

## gas (Google Apps Script)

回答フォームからの設定ファイル生成や、調査結果報告書の作成を行うためのスクリプト

### 初期設定手順

1. `cd gas; npm install` しておきます。
2. 質問票のGoogleフォームを **コピーし** 、メニューの「スクリプトエディタ」を開きます。
3. URLが `projects/.../edit` となっているので、...の部分を `gas/.clasp.json` ファイルの `scriptId` 欄にコピーします。
4. `npx clasp login` で権限が求められるので許可します。
5. `npx clasp push` でスクリプトをアップロードします。
6. スクリプトエディタに戻ってリロードすると、ファイル一覧が増えています。以上で準備完了です。
7. 左側のファイル一覧からファイルを選択して、上側の実行、デバッグボタンの横の関数名を切り替えて実行します。

### スクリプト一覧

- `choices.gs`: `getChoices`
  - 質問表から、選択肢などの一覧である `choices.csv` を生成します。
  - 下側のログに `SpreadSheet created on ...` と出るので、URLを開いて `config/choices.csv` として保存してください。
- `printTabulation.gs`
  - 単純集計結果報告書のベースになるGoogle Documentファイルを作成します。
  - TABULATION_FILE_URLに、 `tabulation.json` ファイルのDriveのURLを貼り付けて実行してください。
  - 下側のログに、作成されたGoogle DocumentsファイルのURLが出力されます。
- `printFullTabulation.gs`: `printFullTabulationPage1`, `printFullTabulationPage2`
  - 調査結果報告書のベースになるGoogle Documentファイルを作成します。
  - FULL_TABULATION_ZIP_FILE_URLに `full_tabulation.zip` ファイルのDriveのURLを貼り付けます。
  - FULL_TABULATION_OUTPUT_DOCUMENT_URLに、手動で作った空のGoogle DocumentsファイルのURLを貼り付けます。
  - ページ数が多すぎてタイムアウトするため、Page1とPage2に分割しています。順番に実行してください。

## data

回答一覧のデータを含むデータを含むファイル群を設置する場所

**！！このディレクトリ以下のファイルは絶対にgit commitしないでください！！**

- `data/source.csv`: 回答一覧の生データ
- `data/after_codings`: アフターコーティングで修正する情報が入るディレクトリ
- `data/after_codings/age_contradictions.csv`: `out/age_contradictions.csv` を参考に、矛盾回答を修正します。
- `data/after_codings/drops.csv`: テスト解答を除外します
- `data/after_codings/others.csv`: 「その他」選択肢の自由記述を、目視で別の選択肢に統合します
- `data/after_codings/exclusives.csv`: `out/exclusives.csv` を参考に、矛盾した選択肢を手動で修正します。
- `data/after_codings/transforms.csv`: 矛盾した回答内容を手動で修正します。
- `data/after_codings/duplicated.csv`:  `out/duplicates.csv` を参考に、二重送信されたと思しきデータを手動で削除します。

## out

各スクリプトによる出力データ

**！！このディレクトリ以下のファイルは絶対にgit commitしないでください！！**

- `after_coding.ipynb` で出力
  - `out/raw.csv`: カラム名にExcelの列名をつけた回答一覧の生データ (drops.csvは適用)
  - `out/free_answers.csv`: 自由記述の回答の一覧
  - `out_duplicates.csv`: `config/duplicated.csv` を利用して、重複送信と思しき回答の一覧
  - `out_exclusives.csv`: `config/exclusive_choices.csv` を利用して、矛盾回答と思しき回答の一覧
  - `out/age_contradictions.csv`: `config/age_contradictions.csv` を利用して、年齢の矛盾回答がある回答の一覧
  - `out/after_coded.csv`: `data/after_codings` 以下ににあるファイルを適用した回答の一覧
  - `out/after_coded.feather`: バイナリ形式で複数回答を配列として保持しており、`plot.ipynb` と `tabulation.ipynb` で使用するファイル
  - `out/drop.csv`: `data/after_codings/drops.csv` の設定か「当てはまるか」に「いいえ」で回答したために除外された回答の一覧
- `plot.ipynb` で出力
  - `out/images`: 結果概要報告のための画像
- `tabulation.ipynb` で出力
  - `out/tabulation.json`: 単純集計結果報告書に使う集計結果の表の情報で、gasの `printTabulation.gs` で使用
  - `out/full_tabulation`: 調査結果報告書に使う集計結果の画像、表、テキストの情報
  - `out/full_tabulation.zip`: 上記をまとめたzipファイルで、gasの `printFullTabulation.gs` で使用

## TODO

- Google Apps ScriptのAPIをスクリプトエディタからではなく直接実行することで、ファイル保存などを自動にしたい。
  - 認証周りがうまくいかず実現できていない。
