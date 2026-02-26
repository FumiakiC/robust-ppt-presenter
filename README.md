# ppt-orchestrator

A zero-dependency, web-based PowerPoint orchestration tool designed for event staff and stage managers. It enables seamless, zero-downtime transitions between multiple speakers' slide decks using just a smartphone or tablet, without requiring any dedicated software installation on the host PC.

イベント進行スタッフやオペレーターのための、インストール不要なPowerPoint進行管理（オーケストレーション）ツールです。専用ソフトウェアをインストールできない社内PC環境でも、スマホやタブレットのWebブラウザから、複数登壇者のスライドをシームレス（待ち時間ゼロ）で切り替えることができます。

## ✨ Key Features (主な機能)

* **Seamless Transitions (シームレスな場面転換):**
    With a single tap on the Web UI, close the current presentation and instantly start the next speaker's slide deck in full-screen mode.
    現在の発表が終わると同時に、スマホからワンタップで次の人のファイルをフルスクリーンで開くことができます。エクスプローラーやデスクトップ画面を観客に見せることはありません。
* **Smart Queue Management (自動キュー管理):**
    Automatically detects `.ppt`/`.pptx` files in the folder and generates a playlist. Finished presentations are automatically moved to a `finish` folder to prevent operational mistakes.
    フォルダ内のPowerPointファイルを自動認識して待機リストを作成します。終わったファイルは自動的に別フォルダに退避されるため、本番中の誤操作を防ぎます。
* **Zero-Dependency & Portable (完全ポータブル):**
    Runs natively on Windows PC using PowerShell and Batch scripts. No Node.js, Python, or third-party web servers are required.
    Windows標準の機能（PowerShell/コマンドプロンプト）だけで動作します。追加のソフトウェアやランタイムのインストールは一切不要です。
* **Robust Network Handling (堅牢な通信処理):**
    Designed for real-world stage use. The built-in HTTP listener safely ignores mobile device sleep or Wi-Fi disconnections, preventing the host PC script from crashing.
    モバイル端末のスリープやWi-Fiの瞬断によるエラーを安全に無視し、プレゼン中の予期せぬスクリプト終了を防ぐ堅牢な設計（V7.3）です。