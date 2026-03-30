# ks(K's Spreadsheet)

## How to build

- dotnet new console -n ks
- cd ks
- copy ~/Program.cs .
- dotnet run

### How to release build ks

- dotnet build -c Release
  - Copy following 3 files in bin\Release\netXX.X\ , and Execute ks.exe in Command Prompt
    - ks.exe
    - ks.dll
    - ks.runtimeconfig.json


## How to download, install, and run ks

- Install `.NET Runtime` in your Windows
  - [.NET Runtime 10.0.5](https://aka.ms/dotnet-core-applaunch?missing_runtime=true&arch=x64&rid=win-x64&os=win10&apphost_version=10.0.5)

- Download `ks.zip`
  - [ks.zip](https://raw.githubusercontent.com/dotnsf/ks/refs/heads/main/release/ks.zip)
- Unzip `ks.zip`, and put following 3 files in same **PATH-specified** folder.
  - `ks.exe`
  - `ks.dll`
  - `ks.runtimeconfig.json`

- Run `ks.exe` from your Command Prompt
  - `> ks`
  - `> ks (filename)`


## How to use ks

- Normal Mode（基本）
  - 移動: h j k l（矢印キーでもOK）
  - 先頭/末尾: gg（先頭行） / G（最終行）
  - 0（先頭列） / $（最終列）
  - 編集: i（セル編集＝Insert モードへ）
  - 矩形選択: v（Visual モードへ）
  - Yank(コピー): y（Normal: 現セル / Visual: 矩形範囲）
  - Paste: p（内部クリップボードを貼り付け）
  - コマンド: :（Command モードへ）
  - ヘルプ: ?

- Insert Mode（「i」から入力してセル編集）
  - ふつうに入力
  - カーソル移動: 矢印キー
  - 保存して抜ける: Ctrl+S
  - 破棄して抜ける: Esc

- Command Mode（「:」または「/」から入力）
  - :w / :w path\to\file.csv : Save(as CSV)
  - :e path\to\file.csv : Read CSV
  - :q : Quit
  - :set width 12 : Set current column width as 12
  - :set width B 20 : Set culumn B width as 20
  - :set auto : Set current column width automatically
  - :help


## Licensing

This code is licensed under MIT.


## Copyright

2026  [K.Kimura @ Juge.Me](https://github.com/dotnsf) all rights reserved.
