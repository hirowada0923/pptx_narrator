# pptx_narrator - Gemini APIを利用したPowerPoint自動音声注釈ツール

## 概要

`pptx_narrator.py`は、PowerPointプレゼンテーション（`.pptx`）のスピーカーノートを読み取り、GoogleのGemini APIを利用してテキストを音声（WAVファイル）に変換し、各スライドに自動で埋め込むPythonスクリプトです。

特に、スピーカーノート内で話者を指定することで、2人の話者が対話する形式のナレーションを簡単に追加できる点が特徴です。

### 主な機能

-   **自動音声注釈**: スピーカーノートのテキストを自動で音声化し、スライドに埋め込みます。
-   **複数話者対応**: ノート内で`Speaker 1:`や`Speaker 2:`と記述するだけで、話者を切り替えることができます。
-   **カスタマイズ性**: 話者の声の種類や話す速度をコマンドラインから簡単に変更できます。
-   **安全なAPIキー管理**: APIキーを環境変数から読み込むため、安全に利用できます。

---

## 前提条件

-   Python 3.8以上
-   Google Gemini APIキー

---

## 準備

### 1. 仮想環境の作成と有効化

プロジェクトの依存関係を管理するため、仮想環境を作成することを推奨します。

```bash
# 仮想環境を作成 (フォルダ名: venv)
python3 -m venv venv

# 仮想環境を有効化 (macOS/Linux)
source venv/bin/activate

# Windowsの場合
# venv\Scripts\activate
```

### 2. 必要なライブラリのインストール

以下のコマンドを実行して、スクリプトに必要なPythonライブラリをインストールします。

```bash
pip install google-genai python-pptx mutagen
```

### 3. APIキーの設定

スクリプトは環境変数 `GEMINI_API_KEY` からGoogle Gemini APIキーを読み込みます。以下のコマンドを実行して、APIキーを設定してください。

```bash
# "YOUR_API_KEY"を実際のキーに置き換えてください
export GEMINI_API_KEY="YOUR_API_KEY"
```
この設定はターミナルセッションごとに必要です。恒久的に設定したい場合は、`.zshrc`や`.bash_profile`などのシェル設定ファイルに追記してください。

---

## 実行手順

以下の形式でコマンドを実行します。

```bash
python3 pptx_narrator.py [入力ファイル名] [出力ファイル名] [オプション]
```

### 引数

-   **`source_pptx`** (必須): 入力となるPowerPointファイル（例: `presentation.pptx`）。
-   **`output_pptx`** (必須): 音声が埋め込まれたPowerPointの出力ファイル名（例: `annotated.pptx`）。
-   **`--name1`** (任意): Speaker 1の声の名前。デフォルトは `Autonoe` です。
-   **`--name2`** (任意): Speaker 2の声の名前。デフォルトは `Algieba` です。
-   **`--speed`** (任意): 話す速度。`1.0`が標準です。数値を大きくすると速くなります。

---

## 実行例 pptx_narrator.pptx参照

### 1. スピーカーノートの準備

PowerPointファイルのスライドのスピーカーノートに、以下のようにナレーションのテキストを記述します。話者を切り替えたい行の先頭に `Speaker 1:` または `Speaker 2:` を追加します。

> Speaker 1: 今日は、PowerPointのスピーカーノートを音声化する便利なツール「pptx_narrator.py」について話していきます。
> Speaker 2: はい、これすごく面白いですよね。Google Gemini APIを使って、各スライドに自動でWAVファイルを挿入してくれるんですよ。
> Speaker 1: しかも、話者を指定することで、こういう対話形式のナレーションも簡単に作れるんです。
> Speaker 2: そうそう。声の種類や話すスピードもカスタマイズできるし、APIキーの安全管理にも対応してるから、安心して使えますよね。
> Speaker 1: プレゼン資料にナレーションを入れたいときや、動画コンテンツを作るときにも重宝しそうです。
> Speaker 2: うん、ナレーション作成のハードルがぐっと下がりますね。

### 2. コマンドの実行

ターミナルで以下のコマンドを実行します。この例では、入力ファイル`pptx_narrator.pptx`から音声付きの`output.pptx`を生成します。

```bash
python3 pptx_narrator.py pptx_narrator.pptx output.pptx --name1 "Autonoe" --name2 "Algieba"
```

### 3. 結果の確認

実行が完了すると、指定した出力ファイル名（この例では `output.pptx`）で音声が埋め込まれたPowerPointファイルが生成されます。

また、処理中に生成された各スライドの音声ファイル（例: `pptx_narrator-001.wav`, `pptx_narrator-002.wav`...）がスクリプトと同じディレクトリに作成されます。
