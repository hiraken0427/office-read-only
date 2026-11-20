#include "office.h"

#include <shellapi.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <wchar.h>

static const wchar_t app_title[] = L"読み取り専用で開く";

typedef struct {
    const wchar_t *input;
    FilePath path;
    OfficeKind kind;
    DWORD path_error;
    OpenResult result;
} Request;

typedef struct {
    wchar_t *text;
    size_t length;
    size_t capacity;
    bool failed;
} Report;

typedef struct {
    const wchar_t *text;
    HWND edit;
} ReportWindow;

static LRESULT CALLBACK report_window_proc(HWND window, UINT message, WPARAM wparam,
                                           LPARAM lparam)
{
    ReportWindow *view = (ReportWindow *)GetWindowLongPtrW(window, GWLP_USERDATA);
    switch (message) {
    case WM_CREATE: {
        const CREATESTRUCTW *create = (const CREATESTRUCTW *)lparam;
        view = create->lpCreateParams;
        SetWindowLongPtrW(window, GWLP_USERDATA, (LONG_PTR)view);
        view->edit = CreateWindowExW(WS_EX_CLIENTEDGE, L"EDIT", L"",
                                     WS_CHILD | WS_VISIBLE | WS_VSCROLL | WS_HSCROLL |
                                     ES_MULTILINE | ES_READONLY | ES_AUTOHSCROLL |
                                     ES_AUTOVSCROLL | ES_NOHIDESEL,
                                     0, 0, 0, 0, window, NULL, create->hInstance, NULL);
        if (!view->edit || !SetWindowTextW(view->edit, view->text)) {
            return -1;
        }
        SendMessageW(view->edit, WM_SETFONT, (WPARAM)GetStockObject(DEFAULT_GUI_FONT),
                     TRUE);
        return 0;
    }
    case WM_SIZE:
        if (view && view->edit) {
            MoveWindow(view->edit, 0, 0, LOWORD(lparam), HIWORD(lparam), TRUE);
        }
        return 0;
    case WM_SETFOCUS:
        if (view && view->edit) {
            SetFocus(view->edit);
        }
        return 0;
    default:
        return DefWindowProcW(window, message, wparam, lparam);
    }
}

static void show_report(HINSTANCE instance, const wchar_t *text)
{
    if (wcslen(text) <= 420) {
        MessageBoxW(NULL, text, app_title, MB_OK | MB_ICONWARNING);
        return;
    }

    WNDCLASSW window_class = {
        .lpfnWndProc = report_window_proc,
        .hInstance = instance,
        .hCursor = LoadCursorW(NULL, IDC_ARROW),
        .lpszClassName = L"OfficeReadOnlyReport"
    };
    ReportWindow view = {.text = text};
    HWND window = NULL;
    if (RegisterClassW(&window_class)) {
        window = CreateWindowExW(0, window_class.lpszClassName, app_title,
                                 WS_OVERLAPPEDWINDOW,
                                 CW_USEDEFAULT, CW_USEDEFAULT, 820, 460,
                                 NULL, NULL, instance, &view);
    }
    if (!window) {
        MessageBoxW(NULL, text, app_title, MB_OK | MB_ICONWARNING);
        return;
    }

    ShowWindow(window, SW_SHOWNORMAL);
    SetFocus(view.edit);
    MSG message;
    while (IsWindow(window) && GetMessageW(&message, NULL, 0, 0) > 0) {
        if (message.message == WM_KEYDOWN && message.wParam == VK_ESCAPE) {
            DestroyWindow(window);
            break;
        }
        if (message.message == WM_KEYDOWN && message.wParam == 'A' &&
            message.hwnd == view.edit && GetKeyState(VK_CONTROL) < 0) {
            SendMessageW(view.edit, EM_SETSEL, 0, -1);
            continue;
        }
        TranslateMessage(&message);
        DispatchMessageW(&message);
    }
}

static void append_text(Report *report, const wchar_t *text)
{
    if (report->failed || !text) {
        return;
    }

    size_t length = wcslen(text);
    size_t max_chars = SIZE_MAX / sizeof(wchar_t);
    if (length >= max_chars - report->length) {
        report->failed = true;
        return;
    }

    size_t needed = report->length + length + 1;
    if (needed > report->capacity) {
        size_t capacity = report->capacity ? report->capacity : 512;
        while (capacity < needed) {
            capacity = capacity <= max_chars / 2 ? capacity * 2 : needed;
        }
        wchar_t *buffer = realloc(report->text, capacity * sizeof(*buffer));
        if (!buffer) {
            report->failed = true;
            return;
        }
        report->text = buffer;
        report->capacity = capacity;
    }
    memcpy(report->text + report->length, text, (length + 1) * sizeof(wchar_t));
    report->length += length;
}

static void append_com_error(Report *report, const ComError *error)
{
    if (!FAILED(error->hr)) {
        return;
    }

    wchar_t number[64];
    append_text(report, L" [");
    append_text(report, error->member);
    swprintf(number, ARRAYSIZE(number), L" / HRESULT 0x%08lX",
             (unsigned long)(DWORD)error->hr);
    append_text(report, number);
    if (error->argument) {
        swprintf(number, ARRAYSIZE(number), L" / 引数 %u", error->argument);
        append_text(report, number);
    }
    if (error->exception_code) {
        swprintf(number, ARRAYSIZE(number), L" / Office 0x%08lX",
                 (unsigned long)(DWORD)error->exception_code);
        append_text(report, number);
    }
    append_text(report, L"]");
    if (error->description && *error->description) {
        append_text(report, L"\r\n");
        append_text(report, error->description);
    }
}

static const wchar_t *path_error_description(DWORD error)
{
    switch (error) {
    case ERROR_NOT_SUPPORTED:
        return L"対応していない拡張子です。";
    case ERROR_FILE_NOT_FOUND:
    case ERROR_PATH_NOT_FOUND:
    case ERROR_BAD_NETPATH:
    case ERROR_BAD_NET_NAME:
        return L"ファイルまたは共有先が見つかりません。";
    case ERROR_DIRECTORY:
        return L"ファイルを選択してください。";
    case ERROR_INVALID_NAME:
        return L"ファイルパスの形式が正しくありません。";
    case ERROR_ACCESS_DENIED:
        return L"ファイルへのアクセスが拒否されました。";
    case ERROR_FILENAME_EXCED_RANGE:
        return L"ファイルパスが長すぎます。";
    case ERROR_NOT_ENOUGH_MEMORY:
        return L"ファイルパス用のメモリを確保できませんでした。";
    default:
        return L"ファイルを確認できませんでした。";
    }
}

static bool request_failed(const Request *request)
{
    return request->path_error != ERROR_SUCCESS ||
           request->result.status == OPEN_FAILED ||
           FAILED(request->result.restore_error.hr);
}

static int report_requests(HINSTANCE instance, const Request *requests, int count)
{
    int failures = 0;
    for (int i = 0; i < count; ++i) {
        if (request_failed(&requests[i])) {
            ++failures;
        }
    }
    if (!failures) {
        return 0;
    }

    wchar_t summary[96];
    swprintf(summary, ARRAYSIZE(summary), L"成功 %d 件、失敗 %d 件\r\n",
             count - failures, failures);
    Report report = {0};
    append_text(&report, summary);

    for (int i = 0; i < count; ++i) {
        const Request *request = &requests[i];
        if (!request_failed(request)) {
            continue;
        }

        append_text(&report, L"\r\n");
        append_text(&report, request->path.absolute ?
                             request->path.absolute : request->input);
        append_text(&report, L"\r\n");
        if (request->path_error != ERROR_SUCCESS) {
            append_text(&report, path_error_description(request->path_error));
            wchar_t code[40];
            swprintf(code, ARRAYSIZE(code), L" [Win32 %lu]",
                     (unsigned long)request->path_error);
            append_text(&report, code);
        } else {
            append_text(&report, office_name(request->kind));
            append_text(&report, L"  ");
            if (request->result.status == OPEN_FAILED) {
                append_text(&report, request->result.reason);
                append_com_error(&report, &request->result.error);
            } else {
                append_text(&report, L"閲覧ウィンドウを表示しました。");
            }
            if (FAILED(request->result.restore_error.hr)) {
                append_text(&report, L"\r\nマクロ設定を元に戻せませんでした。");
                append_com_error(&report, &request->result.restore_error);
            }
        }
        append_text(&report, L"\r\n");
    }

    show_report(instance, report.failed ? L"処理結果を表示するためのメモリを確保できませんでした。"
                                        : report.text);
    free(report.text);
    return 2;
}

static bool process_office(Request *requests, int count)
{
    HANDLE mutex = CreateMutexW(NULL, FALSE, L"Local\\OfficeReadOnly");
    if (!mutex) {
        MessageBoxW(NULL, L"同時起動を確認できませんでした。", app_title,
                    MB_OK | MB_ICONERROR);
        return false;
    }

    DWORD wait_result = WaitForSingleObject(mutex, 30000);
    if (wait_result != WAIT_OBJECT_0 && wait_result != WAIT_ABANDONED) {
        CloseHandle(mutex);
        MessageBoxW(NULL, wait_result == WAIT_TIMEOUT
                               ? L"先に起動した処理が終わってから、再実行してください。"
                               : L"同時起動の待機中にエラーが発生しました。",
                    app_title, MB_OK | MB_ICONERROR);
        return false;
    }

    const wchar_t *stage = L"COM を初期化できませんでした。";
    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
    if (SUCCEEDED(hr)) {
        IMessageFilter *previous = NULL;
        stage = L"COM のメッセージフィルターを登録できませんでした。";
        hr = message_filter_register(&previous);
        if (SUCCEEDED(hr)) {
            OfficeSession session = {0};
            for (int i = 0; i < count; ++i) {
                if (requests[i].path_error == ERROR_SUCCESS) {
                    requests[i].result = office_open(&session, requests[i].kind,
                                                     &requests[i].path);
                }
            }
            office_release(&session);
            message_filter_restore(previous);
        }
        CoUninitialize();
    }
    ReleaseMutex(mutex);
    CloseHandle(mutex);

    if (FAILED(hr)) {
        wchar_t message[160];
        swprintf(message, ARRAYSIZE(message), L"%ls\r\n[HRESULT 0x%08lX]",
                 stage, (unsigned long)(DWORD)hr);
        MessageBoxW(NULL, message, app_title, MB_OK | MB_ICONERROR);
        return false;
    }
    return true;
}

int WINAPI wWinMain(HINSTANCE instance, HINSTANCE previous, PWSTR command_line,
                    int show_command)
{
    (void)previous;
    (void)command_line;
    (void)show_command;

    int argc = 0;
    LPWSTR *argv = CommandLineToArgvW(GetCommandLineW(), &argc);
    if (!argv) {
        MessageBoxW(NULL, L"ファイルの引数を取得できませんでした。", app_title,
                    MB_OK | MB_ICONERROR);
        return 1;
    }
    if (argc < 2) {
        MessageBoxW(NULL, L"エクスプローラーで Office ファイルを選び、"
                         L"送る → 読み取り専用で開くを選んでください。",
                    app_title, MB_OK | MB_ICONINFORMATION);
        LocalFree(argv);
        return 1;
    }

    int count = argc - 1;
    Request *requests = calloc((size_t)count, sizeof(*requests));
    if (!requests) {
        MessageBoxW(NULL, L"ファイル一覧用のメモリを確保できませんでした。", app_title,
                    MB_OK | MB_ICONERROR);
        LocalFree(argv);
        return 1;
    }

    bool has_valid_path = false;
    for (int i = 0; i < count; ++i) {
        Request *request = &requests[i];
        request->input = argv[i + 1];
        request->path_error = path_prepare(request->input, &request->path, &request->kind);
        if (request->path_error == ERROR_SUCCESS) {
            has_valid_path = true;
        }
    }

    int exit_code = 1;
    if (!has_valid_path || process_office(requests, count)) {
        exit_code = report_requests(instance, requests, count);
    }

    for (int i = 0; i < count; ++i) {
        path_dispose(&requests[i].path);
        com_error_clear(&requests[i].result.error);
        com_error_clear(&requests[i].result.restore_error);
    }
    free(requests);
    LocalFree(argv);
    return exit_code;
}
