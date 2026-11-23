# Office Read-Only

Open Excel, Word, and PowerPoint files in read-only mode from the Windows 11 File Explorer **Send to** menu. You can select several files at once, including files with paths on network shares. Office still allows **Save As**.

| App | Supported extensions |
| --- | --- |
| Excel | .xlsx, .xlsm, .xlsb, .xls |
| Word | .docx, .docm, .doc |
| PowerPoint | .pptx, .pptm, .ppt |

## Build

Run the build from the shell.

```sh
make
```

The build writes `bin/OfficeReadOnly.exe`. To remove the generated files, run `make -f clean`.

## Add to the Send to menu

1. Copy `bin\OfficeReadOnly.exe` to `%LOCALAPPDATA%\Programs\OfficeReadOnly\OfficeReadOnly.exe`.
2. Press **Win + R**, enter `shell:sendto`, and press **Enter**.
3. Create a shortcut named `読み取り専用で開く` pointing to the copied executable. Leave the shortcut arguments empty.
4. In File Explorer, select the files, right-click, and choose **Show more options → Send to → 読み取り専用で開く**.

The shortcut is installed for your user account and requires no administrator rights. To remove it, delete the shortcut from `shell:sendto`.

## What to expect

- When Office already has the selected file open read-only, the launcher reuses that window if it can identify it. If it finds the file open for editing, it skips the file and keeps the existing edits.
- When a file opens in Protected View, it stays in Protected View. Excel's Automation security setting has an exception for Excel 4.0 (XLM) macros, which can trigger an Office prompt.
- The launcher reports files it could not open or verify in a single Japanese-language dialog. Close the document in Office when you finish viewing it.
