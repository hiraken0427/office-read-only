#include "office.h"

#include <wchar.h>

enum {
    MSO_FALSE = 0,
    MSO_TRUE = -1,
    MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3
};

typedef struct {
    const wchar_t *name;
    const wchar_t *prog_id;
    const wchar_t *collection;
    WORD item_flags;
} OfficeSpec;

static const OfficeSpec offices[OFFICE_COUNT] = {
    [OFFICE_EXCEL] = {L"Excel", L"Excel.Application", L"Workbooks", DISPATCH_PROPERTYGET},
    [OFFICE_WORD] = {L"Word", L"Word.Application", L"Documents", DISPATCH_METHOD},
    [OFFICE_POWERPOINT] = {L"PowerPoint", L"PowerPoint.Application", L"Presentations",
                          DISPATCH_METHOD}
};

const wchar_t *office_name(OfficeKind kind)
{
    return kind < OFFICE_COUNT ? offices[kind].name : L"";
}

static HRESULT get_application(OfficeSession *session, OfficeKind kind, ComError *error)
{
    IDispatch **application = &session->applications[kind];
    if (*application) {
        return S_OK;
    }

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(offices[kind].prog_id, &clsid);
    if (FAILED(hr)) {
        return com_error_set(error, L"CLSIDFromProgID", hr);
    }
    IUnknown *running = NULL;
    hr = GetActiveObject(&clsid, NULL, &running);
    if (hr == MK_E_UNAVAILABLE) {
        hr = CoCreateInstance(&clsid, NULL, CLSCTX_LOCAL_SERVER,
                              &IID_IDispatch, (void **)application);
        if (FAILED(hr)) {
            return com_error_set(error, L"CoCreateInstance", hr);
        }
    } else if (SUCCEEDED(hr)) {
        hr = IUnknown_QueryInterface(running, &IID_IDispatch, (void **)application);
        IUnknown_Release(running);
        if (FAILED(hr)) {
            return com_error_set(error, L"QueryInterface", hr);
        }
    } else {
        return com_error_set(error, L"GetActiveObject", hr);
    }
    return S_OK;
}

static HRESULT collection_property(IDispatch *object, const wchar_t *name,
                                   IDispatch **collection, ComError *error)
{
    return dispatch_object(object, name, DISPATCH_PROPERTYGET, NULL, 0,
                           collection, error);
}

static HRESULT collection_item(IDispatch *collection, WORD flags, LONG index,
                               IDispatch **item, ComError *error)
{
    VARIANT argument = variant_int(index);
    return dispatch_object(collection, L"Item", flags, &argument, 1, item, error);
}

static HRESULT collection_count(IDispatch *collection, LONG *count, ComError *error)
{
    VARIANT value;
    HRESULT hr = dispatch_property(collection, L"Count", VT_I4, &value, error);
    if (SUCCEEDED(hr)) {
        *count = V_I4(&value);
        VariantClear(&value);
        if (*count < 0) {
            hr = com_error_set(error, L"Count", E_UNEXPECTED);
        }
    }
    return hr;
}

static HRESULT document_path(IDispatch *document, bool protected_view,
                             const FilePath *path, PathMatch *match, ComError *error)
{
    VARIANT folder, name;
    HRESULT hr = dispatch_property(document, protected_view ? L"SourcePath" : L"Path",
                                    VT_BSTR, &folder, error);
    if (FAILED(hr)) {
        return hr;
    }
    if (!protected_view && SysStringLen(V_BSTR(&folder)) == 0) {
        *match = PATH_DIFFERENT; /* 新規文書はまだファイル名を持たない。 */
    } else {
        hr = dispatch_property(document, protected_view ? L"SourceName" : L"FullName",
                               VT_BSTR, &name, error);
        if (SUCCEEDED(hr)) {
            *match = protected_view
                         ? path_compare_source(path, V_BSTR(&folder), V_BSTR(&name))
                         : path_compare(path, V_BSTR(&name));
            VariantClear(&name);
        }
    }
    VariantClear(&folder);
    return hr;
}

/* S_FALSE はすべての候補が別ファイルと確認できた場合に返す。 */
static HRESULT find_document(IDispatch *collection, OfficeKind kind,
                             bool protected_view, const FilePath *path,
                             IDispatch **match, ComError *error)
{
    *match = NULL;
    LONG count;
    HRESULT hr = collection_count(collection, &count, error);
    if (FAILED(hr)) {
        return hr;
    }

    bool ambiguous = false;
    for (LONG i = 0; i < count; ++i) {
        IDispatch *item;
        hr = collection_item(collection, offices[kind].item_flags, i + 1, &item, error);
        if (FAILED(hr)) {
            return hr;
        }

        PathMatch comparison;
        hr = document_path(item, protected_view, path, &comparison, error);
        if (SUCCEEDED(hr) && comparison == PATH_SAME) {
            *match = item;
            return S_OK;
        }
        IDispatch_Release(item);
        if (FAILED(hr)) {
            return hr;
        }
        if (comparison == PATH_UNKNOWN) {
            ambiguous = true;
        }
    }
    if (ambiguous) {
        return com_error_set(error, protected_view ? L"SourcePath / SourceName" : L"FullName",
                             HRESULT_FROM_WIN32(ERROR_CANT_RESOLVE_FILENAME));
    }
    return S_FALSE;
}

static HRESULT find_protected(IDispatch *application, OfficeKind kind,
                              const FilePath *path, IDispatch **window, ComError *error)
{
    IDispatch *collection;
    HRESULT hr = collection_property(application, L"ProtectedViewWindows", &collection,
                                     error);
    if (SUCCEEDED(hr)) {
        hr = find_document(collection, kind, true, path, window, error);
        IDispatch_Release(collection);
    }
    return hr;
}

static HRESULT read_only_state(IDispatch *document, OfficeKind kind, bool *read_only,
                               ComError *error)
{
    VARIANT value;
    bool powerpoint = kind == OFFICE_POWERPOINT;
    HRESULT hr = dispatch_property(document, L"ReadOnly", powerpoint ? VT_I4 : VT_BOOL,
                                   &value, error);
    if (SUCCEEDED(hr)) {
        LONG state = powerpoint ? V_I4(&value) : V_BOOL(&value);
        LONG yes = powerpoint ? MSO_TRUE : VARIANT_TRUE;
        LONG no = powerpoint ? MSO_FALSE : VARIANT_FALSE;
        VariantClear(&value);
        if (state != yes && state != no) {
            return com_error_set(error, L"ReadOnly", DISP_E_TYPEMISMATCH);
        }
        *read_only = state == yes;
    }
    return hr;
}

static HRESULT set_visible(IDispatch *application, OfficeKind kind, ComError *error)
{
    if (kind == OFFICE_POWERPOINT) {
        return S_OK;
    }
    return dispatch_put(application, L"Visible", variant_bool(TRUE), error);
}

static HRESULT excel_user_control(IDispatch *application, ComError *error)
{
    VARIANT value;
    HRESULT hr = dispatch_property(application, L"UserControl", VT_BOOL, &value, error);
    if (FAILED(hr)) {
        return hr;
    }
    bool controlled = V_BOOL(&value) == VARIANT_TRUE;
    VariantClear(&value);
    if (controlled) {
        return S_OK;
    }
    hr = dispatch_put(application, L"UserControl", variant_bool(TRUE), error);
    if (SUCCEEDED(hr)) {
        hr = dispatch_property(application, L"UserControl", VT_BOOL, &value, error);
        if (SUCCEEDED(hr)) {
            controlled = V_BOOL(&value) == VARIANT_TRUE;
            VariantClear(&value);
            if (!controlled) {
                hr = com_error_set(error, L"UserControl", E_FAIL);
            }
        }
    }
    return hr;
}

static HRESULT activate_window(IDispatch *application, IDispatch *document,
                               OfficeKind kind, bool protected_view, ComError *error)
{
    HRESULT hr = set_visible(application, kind, error);
    if (FAILED(hr)) {
        return hr;
    }

    if (protected_view || kind == OFFICE_WORD) {
        hr = dispatch_invoke(document, L"Activate", DISPATCH_METHOD, NULL, 0, NULL,
                             error);
    } else {
        IDispatch *windows;
        hr = collection_property(document, L"Windows", &windows, error);
        if (FAILED(hr)) {
            return hr;
        }

        IDispatch *window;
        hr = collection_item(windows, offices[kind].item_flags, 1, &window, error);
        IDispatch_Release(windows);
        if (FAILED(hr)) {
            return hr;
        }
        if (kind == OFFICE_EXCEL) {
            hr = dispatch_put(window, L"Visible", variant_bool(TRUE), error);
        }
        if (SUCCEEDED(hr)) {
            hr = dispatch_invoke(window, L"Activate", DISPATCH_METHOD,
                                 NULL, 0, NULL, error);
        }
        IDispatch_Release(window);
    }

    if (SUCCEEDED(hr) && kind == OFFICE_EXCEL) {
        hr = excel_user_control(application, error);
    }
    return hr;
}

static HRESULT invoke_open(IDispatch *collection, OfficeKind kind, BSTR filename,
                           IDispatch **document, ComError *error)
{
    VARIANT arguments[13];
    VARIANT missing;
    VariantInit(&missing);
    V_VT(&missing) = VT_ERROR;
    V_ERROR(&missing) = DISP_E_PARAMNOTFOUND;
    for (size_t i = 0; i < ARRAYSIZE(arguments); ++i) {
        arguments[i] = missing;
    }
    V_VT(&arguments[0]) = VT_BSTR;
    V_BSTR(&arguments[0]) = filename;

    UINT count;
    switch (kind) {
    case OFFICE_EXCEL:
        arguments[1] = variant_int(0);      /* UpdateLinks */
        arguments[2] = variant_bool(TRUE);  /* ReadOnly */
        arguments[6] = variant_bool(TRUE);  /* IgnoreReadOnlyRecommended */
        arguments[10] = variant_bool(FALSE); /* Notify */
        arguments[12] = variant_bool(FALSE); /* AddToMru */
        count = 13;
        break;
    case OFFICE_WORD:
        arguments[1] = variant_bool(FALSE); /* ConfirmConversions */
        arguments[2] = variant_bool(TRUE);  /* ReadOnly */
        arguments[3] = variant_bool(FALSE); /* AddToRecentFiles */
        arguments[6] = variant_bool(FALSE); /* Revert */
        count = 7;
        break;
    case OFFICE_POWERPOINT:
        arguments[1] = variant_int(MSO_TRUE);  /* ReadOnly */
        arguments[2] = variant_int(MSO_FALSE); /* Untitled */
        arguments[3] = variant_int(MSO_TRUE);  /* WithWindow */
        count = 4;
        break;
    default:
        return com_error_set(error, L"Open", E_INVALIDARG);
    }
    return dispatch_object(collection, L"Open", DISPATCH_METHOD, arguments, count,
                           document, error);
}

static HRESULT set_security(IDispatch *application, LONG security, ComError *error)
{
    HRESULT hr = dispatch_put(application, L"AutomationSecurity",
                              variant_int(security), error);
    if (SUCCEEDED(hr)) {
        VARIANT current;
        hr = dispatch_property(application, L"AutomationSecurity", VT_I4, &current,
                               error);
        if (SUCCEEDED(hr)) {
            if (V_I4(&current) != security) {
                hr = com_error_set(error, L"AutomationSecurity", E_FAIL);
            }
            VariantClear(&current);
        }
    }
    return hr;
}

static void open_checked(IDispatch *application, IDispatch *collection,
                         OfficeKind kind, const FilePath *path, OpenResult *result)
{
    BSTR filename = SysAllocString(path->absolute);
    if (!filename) {
        result->reason = L"ファイルを開くためのメモリを確保できませんでした。";
        com_error_set(&result->error, L"Open", E_OUTOFMEMORY);
        return;
    }

    VARIANT original;
    result->reason = L"マクロの実行を抑制する設定を確認できませんでした。";
    HRESULT hr = dispatch_property(application, L"AutomationSecurity", VT_I4,
                                   &original, &result->error);
    if (FAILED(hr)) {
        SysFreeString(filename);
        return;
    }

    IDispatch *document = NULL;
    hr = set_security(application, MSO_AUTOMATION_SECURITY_FORCE_DISABLE,
                      &result->error);
    bool attempted_open = SUCCEEDED(hr);
    if (attempted_open) {
        result->reason = L"ファイルを開けませんでした。";
        hr = invoke_open(collection, kind, filename, &document, &result->error);
    }
    set_security(application, V_I4(&original), &result->restore_error);
    VariantClear(&original);
    SysFreeString(filename);

    if (!attempted_open) {
        return;
    }

    /* Open が失敗を返しても、Office が保護ビューを作成した場合がある。 */
    ComError protected_error = {0};
    IDispatch *protected_window = NULL;
    HRESULT protected_hr = find_protected(application, kind, path, &protected_window,
                                          &protected_error);
    com_error_clear(&protected_error);

    if (protected_hr == S_OK) {
        com_error_clear(&result->error);
        result->reason = L"保護ビューを表示できませんでした。";
        if (SUCCEEDED(activate_window(application, protected_window, kind, true,
                                      &result->error))) {
            result->status = OPEN_PROTECTED;
        }
        IDispatch_Release(protected_window);
    } else if (SUCCEEDED(hr)) {
        PathMatch comparison;
        bool read_only;
        result->reason = L"読み取り専用を確認できませんでした。Office 側のウィンドウを確認してください。";
        hr = document_path(document, false, path, &comparison, &result->error);
        if (SUCCEEDED(hr) && comparison == PATH_SAME &&
            SUCCEEDED(read_only_state(document, kind, &read_only, &result->error)) &&
            read_only) {
            result->reason = L"読み取り専用のウィンドウを表示できませんでした。";
            if (SUCCEEDED(activate_window(application, document, kind, false,
                                          &result->error))) {
                result->status = OPEN_NEW;
            }
        }
    }

    if (document) {
        IDispatch_Release(document);
    }
}

OpenResult office_open(OfficeSession *session, OfficeKind kind, const FilePath *path)
{
    OpenResult result = {0};
    if (session->restore_failed[kind]) {
        result.reason = L"このアプリのマクロ設定を元に戻せなかったため、処理を中止しました。";
        return result;
    }

    result.reason = L"Office を起動または取得できませんでした。";
    if (FAILED(get_application(session, kind, &result.error))) {
        return result;
    }

    IDispatch *application = session->applications[kind];
    IDispatch *collection;
    result.reason = L"すでに開いているファイルの状態を確認できませんでした。";
    if (FAILED(collection_property(application, offices[kind].collection,
                                   &collection, &result.error))) {
        return result;
    }

    IDispatch *document = NULL;
    HRESULT hr = find_document(collection, kind, false, path, &document,
                               &result.error);
    if (hr == S_OK) {
        bool read_only;
        if (SUCCEEDED(read_only_state(document, kind, &read_only, &result.error))) {
            result.reason = L"同じファイルを編集中です。";
            if (read_only) {
                result.reason = L"既存の読み取り専用ウィンドウを表示できませんでした。";
                if (SUCCEEDED(activate_window(application, document, kind, false,
                                              &result.error))) {
                    result.status = OPEN_REUSED;
                }
            }
        }
        IDispatch_Release(document);
    } else if (hr == S_FALSE) {
        IDispatch *protected_window = NULL;
        hr = find_protected(application, kind, path, &protected_window, &result.error);
        if (hr == S_OK) {
            result.reason = L"保護ビューを表示できませんでした。";
            if (SUCCEEDED(activate_window(application, protected_window, kind, true,
                                          &result.error))) {
                result.status = OPEN_PROTECTED;
            }
            IDispatch_Release(protected_window);
        } else if (hr == S_FALSE) {
            result.reason = L"Office のウィンドウを表示できませんでした。";
            if (SUCCEEDED(set_visible(application, kind, &result.error))) {
                open_checked(application, collection, kind, path, &result);
            }
        }
    }
    IDispatch_Release(collection);
    if (FAILED(result.restore_error.hr)) {
        session->restore_failed[kind] = true;
    }
    return result;
}

void office_release(OfficeSession *session)
{
    for (size_t i = 0; i < OFFICE_COUNT; ++i) {
        if (session->applications[i]) {
            IDispatch_Release(session->applications[i]);
            session->applications[i] = NULL;
        }
    }
}
