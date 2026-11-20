# Windows 11「読み取り専用で開く」詳細設計

エクスプローラーで Office ファイルを選択し、右クリック → その他のオプションを表示 → 送る → 読み取り専用で開く、の順に操作する。デスクトップ版 Excel・Word・PowerPoint に対象を表示し、元ファイルへの通常の上書き保存を防ぐ。必要な場合は Office 側で別名保存できる。すでに編集中の文書と未保存変更はそのまま維持する。

| 項目 | 仕様 |
| --- | --- |
| 実行ファイル | `OfficeReadOnly.exe` |
| ビルド | C11、w64devkit（GCC 16.2.0 を基準）、`-O2`、x64 Windows GUI |
| 実行環境 | Windows 11、対応するデスクトップ版 Office をインストール済み |
| 入力 | エクスプローラーの「送る」が渡すファイルパス。複数指定可。日本語は UTF-16 のまま扱う |
| 権限 | ログオン中のユーザーで実行・登録。管理者権限は不要 |
| 成功時 | Office のウィンドウだけを表示。閲覧を続けられる状態で起動ツールは終了 |

## 1. 入力と処理順

1. `wWinMain` から `CommandLineToArgvW(GetCommandLineW(), &argc)` を呼び、`argv[1]` 以降を到着順に処理する。引数なしの場合は「エクスプローラーで Office ファイルを選び、送る → 読み取り専用で開くを選んでください」と表示する。
2. `GetFullPathNameW` で絶対パスを得る。可変長バッファを使い、`GetFileAttributesW` で存在とファイル種別を確認する。ローカル、ドライブ文字付きの共有フォルダー、`\\server\share\...` を受け付ける。ディレクトリ、デバイスパス、URL は対象外。長いパスは Windows API と Office の両方で扱える長さを実機で確認する。OS 側で解決できても Office が開けなければ個別の失敗として扱う。
3. ファイル名末尾の拡張子を `CompareStringOrdinal` で大文字小文字を区別せずに判定する。エラーや対象外のファイルは記録し、次へ進む。複数ファイルの実行範囲は Windows とエクスプローラーが起動時に渡せる引数数・長さに従う。
4. 入力がすべて処理対象外なら、その場で結果を表示する。処理対象があればミューテックスを取得して STA を初期化し、有効な入力を順番に開く。起動中に Office がダイアログを表示した場合は、ユーザーが操作するまでそのファイルの処理が続く。

| 種別 | 拡張子 | COM ProgID | 開くメソッド |
| --- | --- | --- | --- |
| Excel | `.xlsx` `.xlsm` `.xlsb` `.xls` | `Excel.Application` | `Workbooks.Open` |
| Word | `.docx` `.docm` `.doc` | `Word.Application` | `Documents.Open` |
| PowerPoint | `.pptx` `.pptm` `.ppt` | `PowerPoint.Application` | `Presentations.Open` |

## 2. 既存文書の扱い

対象アプリは必要になった時点で取得し、同じ起動内で再利用する。最初に `CLSIDFromProgID`、`GetActiveObject`、`QueryInterface(IID_IDispatch)` で既存のアプリを探す。`MK_E_UNAVAILABLE` なら `CoCreateInstance` を `CLSCTX_LOCAL_SERVER` で呼ぶ。返されたアプリのコレクションを毎回確認する。`CoCreateInstance` で取得した場合も、既存プロセスを共有することがある。

| 開く直前の確認結果 | 処理 |
| --- | --- |
| 同じファイルが読み取り専用 | 既存の文書ウィンドウをアクティブにする。`Open` は呼ばない |
| 同じファイルが保護ビューにある | 確認できた `ProtectedViewWindow` をアクティブにする |
| 同じファイルが書き込み可能 | 「同じファイルを編集中」として結果に記録する。既存ウィンドウの表示状態や編集内容は維持する |
| 既存文書の状態、または同一性を判定できない | そのファイルの処理を中止し、確認できなかった箇所を表示する |
| 該当する文書が見つからない | 手順 3 の引数で開く |

- 取得したアプリの通常文書コレクションと `ProtectedViewWindows` を、`Count` と `Item` で走査する。通常文書は `FullName`、保護ビューは `SourcePath` と `SourceName` を使う。未保存の新規文書はディスク上のファイルとの照合から除外する。
- 照合では、双方が解決可能なら `CreateFileW` を属性読み取り・共有読み書き削除許可で開き、`GetFileInformationByHandleEx(FileIdInfo)` のファイル ID を同じボリューム内で比較する。ハンドルは照合中だけ保持する。ファイル ID を取得できない場合は、解決できた絶対パスを比較する。文字列が完全に一致すれば同一と扱う。大文字小文字だけの差、共有先の別名、Office が返すクラウド URL などで同一性が不明な候補は判定不能とする。
- `GetActiveObject` の確認範囲は取得したアプリインスタンスに限る。別インスタンスに同じファイルがある場合は、取得したアプリが開ける範囲で読み取り専用の結果を確認する。別インスタンスのプロセスや文書を閉じる操作は行わない。
- Word の `Revert=False` は、同じ文書が開いていると既存文書をアクティブにする。開く直前の確認に加え、戻り値の `ReadOnly` も必ず検証する。Excel で同じファイル名の別ブックによる競合が報告された場合も、Office の既存ブックを置き換えず、理由を表示する。 [Word Documents.Open](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Word.Documents.Open.md)

## 3. 開く引数と結果判定

引数番号は Office API の宣言順（1 始まり）。パスは `VT_BSTR`、Boolean は `VT_BOOL` の `VARIANT_TRUE` または `VARIANT_FALSE`。列挙値・整数は `VT_I4` を使う。表にない中間引数は `VT_ERROR` と `DISP_E_PARAMNOTFOUND` で省略する。

| メソッド | 渡す引数 | 成功判定 |
| --- | --- | --- |
| `Workbooks.Open` | 1 `Filename=絶対パス`、2 `UpdateLinks=0`、3 `ReadOnly=True`、7 `IgnoreReadOnlyRecommended=True`、11 `Notify=False`、13 `AddToMru=False` | `Workbook.ReadOnly == VARIANT_TRUE` |
| `Documents.Open` | 1 `FileName=絶対パス`、2 `ConfirmConversions=False`、3 `ReadOnly=True`、4 `AddToRecentFiles=False`、7 `Revert=False`。文書ウィンドウは既定で表示 | `Document.ReadOnly == VARIANT_TRUE` |
| `Presentations.Open` | 1 `FileName=絶対パス`、2 `ReadOnly=msoTrue`、3 `Untitled=msoFalse`、4 `WithWindow=msoTrue`。`msoTrue=-1`、`msoFalse=0` | `Presentation.ReadOnly == msoTrue` |

1. 既存文書の確認が済んだら、新規に開く Excel・Word では `Application.Visible=True` を指定し、Office が出す入力ダイアログを操作できるようにする。手順 4 のセキュリティ設定を適用して `Open` を呼ぶ。戻った文書の `FullName` と `ReadOnly` を確認する。`ReadOnly` の取得失敗、想定外の型、書き込み可能状態は失敗として報告する。
2. Office が保護ビューで表示した場合は、`ProtectedViewWindows` から対象ファイルを確認し、保護ビューのまま表示成功とする。`Open` のエラーだけから保護ビューで開いたとは判断しない。保護ビューの解除、編集の有効化、通常の `Open` の再実行を自動で行わない。
3. 対象ファイルの照合と読み取り専用判定が完了したら、ウィンドウをアクティブにする。Excel は `Application.Visible=True` と対象 `Workbook.Windows` の表示・アクティブ化、Word は `Application.Visible=True` と `Document.Activate`、PowerPoint は対象 `Presentation.Windows.Item(1).Activate` を使う。複数のウィンドウがある場合も対象文書に属するものを選ぶ。ウィンドウを表示できない場合は失敗として報告する。`SetForegroundWindow` を使った強制前面化は行わない。
4. `Open` 後に状態が確認できなかった場合は「読み取り専用を確認できませんでした。Office 側のウィンドウを確認してください」と表示する。Office がすでにウィンドウを表示している可能性がある。保存や変更の破棄を伴う操作は呼ばない。

[Excel Workbooks.Open](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Workbooks.Open.md) · [Word Documents.Open](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Word.Documents.Open.md) · [PowerPoint Presentations.Open](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/PowerPoint.Presentations.Open.md) · [PowerPoint Presentation.ReadOnly](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/PowerPoint.Presentation.ReadOnly.md)

## 4. マクロと Office のダイアログ

- 新しく開くファイルごとに `Application.AutomationSecurity` を読み、元の値を保存する。`msoAutomationSecurityForceDisable`（`3`）を設定し、読み戻して確認してから `Open` を呼ぶ。取得・設定・読み戻しのいずれかが失敗した場合は、そのまま開く処理を中止する。
- 設定を試みたら、設定途中の失敗を含むすべての終了経路で保存した元の値に戻す。`Open` が終了したら、結果を検証する前に復元し、読み戻す。復元に失敗した場合はそのアプリでの後続のファイル処理を中止し、設定を元に戻せなかったことと直前の処理結果を報告する。他の種類の Office ファイルは処理を続ける。
- `AutomationSecurity` はアプリ全体に作用する。ツールの同時起動は `CreateMutexW` で作成するセッション内の名前付きミューテックス `Local\OfficeReadOnly` で直列化する。STA 初期化前に `WaitForSingleObject` で最大 30 秒待機し、タイムアウト時は先の処理が終わってから再実行するよう表示する。待機から所有権を得た場合は COM 処理の終了時に `ReleaseMutex` と `CloseHandle` を行い、その後に結果を表示する。Office を手動操作するユーザーや別の Automation クライアントは、このミューテックスの対象外。
- Excel の `UpdateLinks=0` で、`Workbooks.Open` が扱う外部参照の更新を抑える。3 アプリの保護ビューやトラストセンターの判断は Office に任せる。`DisplayAlerts` の一括変更、Office のセキュリティ設定やポリシーの変更、ダイアログでの自動承認は行わない。パスワード入力が必要な場合は Office のダイアログで処理し、ツール側でパスワードを保持しない。
- Excel 4.0（XLM）マクロは `ForceDisable` の対象外で、Excel が警告を表示する場合がある。XLM を含めて組織として実行を禁止する場合は、Office のポリシーで無効にする。既存ウィンドウの再利用時は、その文書が最初に開かれたときのセキュリティ設定が適用された状態のまま表示する。 [Excel AutomationSecurity](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Application.AutomationSecurity.md) · [Word AutomationSecurity](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Word.Application.AutomationSecurity.md) · [PowerPoint AutomationSecurity](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/PowerPoint.Application.AutomationSecurity.md)

## 5. COM 呼び出しと終了処理

- UI と Automation は `CoInitializeEx(NULL, COINIT_APARTMENTTHREADED)` を呼んだ 1 スレッドで扱う。`S_OK` と `S_FALSE` のどちらでも、解放処理の最後に `CoUninitialize()` を 1 回呼ぶ。初期化に失敗した場合は起動エラーにする。
- `IDispatch::GetIDsOfNames` と `Invoke` を共通化する。引数を `DISPPARAMS.rgvarg` に逆順で入れ、メソッドは `DISPATCH_METHOD`、プロパティ取得は `DISPATCH_PROPERTYGET`、プロパティ設定は `DISPATCH_PROPERTYPUT` と `DISPID_PROPERTYPUT` を使用する。`Item` はコレクションの定義に従い、Excel の `Workbooks.Item` はプロパティ取得、Word の `Documents.Item` と PowerPoint の `Presentations.Item` はメソッドとして呼ぶ。例えば `Presentations.Open` の `rgvarg` は `[msoTrue, msoFalse, msoTrue, FileName]`。
- `VARIANT` は `VariantInit` し、所有する値を `VariantClear` する。戻り値が `VT_DISPATCH` で、その参照を別の変数で保持する場合は `AddRef` してから `VariantClear` する。`VT_UNKNOWN` は `QueryInterface` で `IDispatch` を取得する。取得した参照は解放し、`CommandLineToArgvW` の領域は `LocalFree` する。
- `Invoke` の失敗時に `HRESULT`、必要なら `EXCEPINFO` の説明を残す。遅延取得用の `pfnDeferredFillIn` があれば呼び、`EXCEPINFO` の BSTR を解放する。`puArgErr` が有効な場合は逆順の添字を宣言順の番号に直して記録する。
- `CoRegisterMessageFilter` で STA に `IMessageFilter` を登録する。`RetryRejectedCall` は `SERVERCALL_RETRYLATER` の場合に限り、最初の呼び出しから 3 秒以内に収まる範囲で 200 ms 待って再試行する。`SERVERCALL_REJECTED` や応答を失った呼び出しは自動で再実行しない。応答を待つ間は `MessagePending` で `PENDINGMSG_WAITDEFPROCESS` を返す。この上限は COM が一時拒否した場合に適用し、Office 自身の入力待ちには適用しない。終了時に元のフィルターを復元し、所有する参照を解放する。
- Excel は表示後に `UserControl` を確認する。`False` なら `True` に設定して確認する。引き渡しを確認できない場合は表示の失敗として報告する。Word と PowerPoint は実機で COM 参照解放後のウィンドウ保持を確認する。文書を閉じる操作や `Application.Quit()` は起動ツールから呼ばない。 [Excel Application.UserControl](https://raw.githubusercontent.com/MicrosoftDocs/VBA-Docs/main/api/Excel.Application.UserControl.md)

## 6. 結果表示

成功は「新規に開いた」「既存を表示した」「保護ビューで表示した」のいずれかとする。全件成功なら通知せず終了する。失敗があれば最後に成功件数と失敗件数を出し、失敗した全パスと理由を入力順に表示する。短い結果は `MessageBoxW`、収まりきらない結果はスクロールとコピーができる読み取り専用テキスト欄に表示する。COM エラーは日本語の要約に続けて、必要に応じて Office 名、呼び出し箇所、`HRESULT` を示す。

| 終了コード | 意味 |
| --- | --- |
| `0` | 指定されたすべてのファイルを表示できた |
| `1` | 引数なし、または起動時の共通処理に失敗した |
| `2` | 1 件以上のファイルを処理できなかった、またはセキュリティ設定の復元に失敗した |

## 7. ファイル構成とビルド

```text
MAKEFILE
README.md
doc/
  design.md
src/
  main.c         wWinMain、処理順、結果表示
  path.c/.h      UTF-16 パス正規化、拡張子判定、ファイルの照合
  dispatch.c/.h  IDispatch 呼び出し、VARIANT とエラー処理、IMessageFilter
  office.c/.h    アプリ取得、既存文書確認、Open、読み取り専用判定と表示
bin/
  OfficeReadOnly.exe
  obj/           中間オブジェクト
```

`MAKEFILE` と `README.md` をルートに置く。設計書は `doc`、ソースは `src`、ビルド生成物と必要に応じたバイナリは `bin` に置く。ソース分割は上記のファイル単位で始める。サブフォルダーが必要になった場合は VS Code の Material Icon Theme が認識する名前を選ぶ。

64bit の w64devkit シェルで `gcc -dumpmachine` が `x86_64-w64-mingw32` を示すことを確認し、`make -f MAKEFILE` でビルドする。コンパイル時に `-std=c11 -O2 -Wall -Wextra -Wpedantic -DUNICODE -D_UNICODE -D_WIN32_WINNT=0x0A00` を指定し、リンク時は `-municode -mwindows` と `-lole32 -loleaut32 -lshell32 -luser32 -luuid` を使う。ライブラリはオブジェクトの後ろに置く。`objdump -f` と `objdump -p` で x64 の PE と Windows GUI サブシステムを確認する。

## 8. Windows 11 の「送る」への登録

1. `bin\OfficeReadOnly.exe` を `%LOCALAPPDATA%\Programs\OfficeReadOnly\OfficeReadOnly.exe` に配置する。
2. `Win + R` で `shell:sendto` を開き、配置先の exe を指す `読み取り専用で開く.lnk` を作る。リンク先には exe の絶対パスを指定し、ショートカットの引数欄は空欄にする。
3. ファイルを選択し、右クリック → その他のオプションを表示 → 送る → 読み取り専用で開く、から起動する。更新時は配置先の exe を入れ替える。登録を外すときは、このショートカットを削除する。

## 9. 実機での受け入れ確認

Windows 11 とデスクトップ版 Office を使用する。結果とあわせて Windows・Office のビルド、Office のビット数を記録する。Office が閉じた状態と起動済みの状態の両方で確認する。

- [ ] Excel・Word・PowerPoint で対象ファイルが開き、読み取り専用状態を確認できる。旧形式も含めて対応拡張子を各 1 ファイル以上試す。
- [ ] 閲覧用ファイルを変更して `Ctrl+S` を押しても、元ファイルの内容と最終更新日時が変わらない。別名保存する場合は別ファイルが作られる。
- [ ] 編集中のファイルを選んでも、既存の未保存変更が保持される。同じインスタンスでは編集中と表示し、別インスタンスがある場合も編集側の文書が閉じられたり保存されたりしない。
- [ ] 読み取り専用で開いているファイルをもう一度選ぶと、検出できた既存のウィンドウが再利用される。同名の別ファイルを選んだ場合は、目的のファイルまたは競合エラーを確認できる。
- [ ] 日本語、空白、括弧、`&` を含むパスと大文字拡張子を扱える。全角・半角を含む長いパスは Office の対応範囲内で確認する。
- [ ] Excel・Word・PowerPoint の混在した選択に対象外のファイルを加えても、有効なファイルを順に処理し、失敗したファイルを一度の結果表示で確認できる。
- [ ] `.xlsm`、`.docm`、`.pptm` の自動実行マクロが動作しない。XLM は別に確認し、Excel の確認画面が出た場合はツールが承認しないことを確認する。保護ビューのファイルは保護ビューのまま表示される。
- [ ] 起動ツールが終了した後も 3 アプリの閲覧ウィンドウが残り、ユーザーが開閉や閲覧を続けられる。特に Office 未起動からの Excel を確認する。
- [ ] UNC と、同じ共有先を割り当てたドライブの両方から開く。共有先のファイルを編集中の場合やアクセス権がない場合も確認する。

異常系は、引数なし、存在しないパス、ディレクトリ、破損ファイル、パスワード入力のキャンセル、Office 未インストールまたは COM 起動失敗、`ReadOnly` の取得失敗、マクロ設定の取得・設定・復元失敗を試す。COM の一時拒否では指定時間内の復帰と上限到達の両方を確認する。正常時と `Open` 失敗時の両方で `AutomationSecurity` が元の値に戻ることも確認する。実機で再現しづらい COM エラーは `IDispatch` のテスト用実装で戻り値を差し替え、エラー表示と後続処理を確かめる。