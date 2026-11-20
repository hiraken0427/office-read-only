#include "dispatch.h"

void com_error_clear(ComError *error)
{
    SysFreeString(error->description);
    *error = (ComError){0};
}

HRESULT com_error_set(ComError *error, const wchar_t *member, HRESULT hr)
{
    com_error_clear(error);
    error->member = member;
    error->hr = hr;
    return hr;
}

HRESULT dispatch_invoke(IDispatch *object, const wchar_t *member, WORD flags,
                        const VARIANT *arguments, UINT count, VARIANT *result,
                        ComError *error)
{
    /* 最も引数が多いのは、AddToMru まで指定する Excel の Workbooks.Open。 */
    VARIANTARG reversed[13];
    if (count > ARRAYSIZE(reversed)) {
        return com_error_set(error, member, E_INVALIDARG);
    }

    DISPID id;
    LPOLESTR name = (LPOLESTR)member;
    HRESULT hr = IDispatch_GetIDsOfNames(object, &IID_NULL, &name, 1,
                                         LOCALE_USER_DEFAULT, &id);
    if (FAILED(hr)) {
        return com_error_set(error, member, hr);
    }

    /* 元の値を Invoke の間だけ借りるため、この配列からは解放しない。 */
    for (UINT i = 0; i < count; ++i) {
        reversed[i] = arguments[count - i - 1];
    }
    DISPPARAMS params = {.rgvarg = count ? reversed : NULL, .cArgs = count};
    DISPID put = DISPID_PROPERTYPUT;
    if (flags == DISPATCH_PROPERTYPUT) {
        params.rgdispidNamedArgs = &put;
        params.cNamedArgs = 1;
    }

    EXCEPINFO exception = {0};
    UINT argument = count;
    hr = IDispatch_Invoke(object, id, &IID_NULL, LOCALE_USER_DEFAULT, flags,
                          &params, result, &exception, &argument);
    if (FAILED(hr)) {
        com_error_set(error, member, hr);
        if (hr == DISP_E_EXCEPTION) {
            if (exception.pfnDeferredFillIn) {
                exception.pfnDeferredFillIn(&exception);
            }
            error->description = exception.bstrDescription;
            exception.bstrDescription = NULL;
            error->exception_code = exception.scode;
            if (!error->exception_code && exception.wCode) {
                error->exception_code = MAKE_HRESULT(SEVERITY_ERROR,
                                                      FACILITY_CONTROL,
                                                      exception.wCode);
            }
        }
        if ((hr == DISP_E_TYPEMISMATCH || hr == DISP_E_PARAMNOTFOUND) &&
            argument < count) {
            error->argument = count - argument;
        }
    }
    SysFreeString(exception.bstrSource);
    SysFreeString(exception.bstrDescription);
    SysFreeString(exception.bstrHelpFile);
    return hr;
}

HRESULT dispatch_object(IDispatch *object, const wchar_t *member, WORD flags,
                        const VARIANT *arguments, UINT count, IDispatch **result,
                        ComError *error)
{
    VARIANT value;
    VariantInit(&value);
    *result = NULL;

    HRESULT hr = dispatch_invoke(object, member, flags, arguments, count,
                                 &value, error);
    if (SUCCEEDED(hr)) {
        if (V_VT(&value) == VT_DISPATCH && V_DISPATCH(&value)) {
            *result = V_DISPATCH(&value);
            IDispatch_AddRef(*result);
        } else if (V_VT(&value) == VT_UNKNOWN && V_UNKNOWN(&value)) {
            hr = IUnknown_QueryInterface(V_UNKNOWN(&value), &IID_IDispatch,
                                         (void **)result);
        } else {
            hr = DISP_E_TYPEMISMATCH;
        }
        if (FAILED(hr)) {
            com_error_set(error, member, hr);
        }
    }
    VariantClear(&value);
    return hr;
}

HRESULT dispatch_property(IDispatch *object, const wchar_t *member, VARTYPE type,
                          VARIANT *result, ComError *error)
{
    VariantInit(result);
    HRESULT hr = dispatch_invoke(object, member, DISPATCH_PROPERTYGET,
                                 NULL, 0, result, error);
    if (SUCCEEDED(hr) && V_VT(result) != type) {
        hr = com_error_set(error, member, DISP_E_TYPEMISMATCH);
    }
    if (FAILED(hr)) {
        VariantClear(result);
    }
    return hr;
}

HRESULT dispatch_put(IDispatch *object, const wchar_t *member, VARIANT value,
                     ComError *error)
{
    return dispatch_invoke(object, member, DISPATCH_PROPERTYPUT, &value, 1,
                           NULL, error);
}

VARIANT variant_bool(BOOL value)
{
    VARIANT result;
    VariantInit(&result);
    V_VT(&result) = VT_BOOL;
    V_BOOL(&result) = value ? VARIANT_TRUE : VARIANT_FALSE;
    return result;
}

VARIANT variant_int(LONG value)
{
    VARIANT result;
    VariantInit(&result);
    V_VT(&result) = VT_I4;
    V_I4(&result) = value;
    return result;
}

static LONG filter_references = 1;

static ULONG STDMETHODCALLTYPE filter_add_ref(IMessageFilter *self)
{
    (void)self;
    return (ULONG)InterlockedIncrement(&filter_references);
}

static ULONG STDMETHODCALLTYPE filter_release(IMessageFilter *self)
{
    (void)self;
    return (ULONG)InterlockedDecrement(&filter_references);
}

static HRESULT STDMETHODCALLTYPE filter_query(IMessageFilter *self,
                                               REFIID iid, void **result)
{
    if (!result) {
        return E_POINTER;
    }
    *result = NULL;
    if (IsEqualIID(iid, &IID_IUnknown) || IsEqualIID(iid, &IID_IMessageFilter)) {
        *result = self;
        filter_add_ref(self);
        return S_OK;
    }
    return E_NOINTERFACE;
}

static DWORD STDMETHODCALLTYPE filter_incoming(IMessageFilter *self, DWORD type,
                                               HTASK caller, DWORD elapsed,
                                               LPINTERFACEINFO info)
{
    (void)self;
    (void)type;
    (void)caller;
    (void)elapsed;
    (void)info;
    return SERVERCALL_ISHANDLED;
}

static DWORD STDMETHODCALLTYPE filter_rejected(IMessageFilter *self, HTASK callee,
                                               DWORD elapsed, DWORD rejection)
{
    (void)self;
    (void)callee;
    const DWORD delay = 200;
    const DWORD limit = 3000;
    if (rejection == SERVERCALL_RETRYLATER && elapsed <= limit - delay) {
        return delay;
    }
    return (DWORD)-1;
}

static DWORD STDMETHODCALLTYPE filter_pending(IMessageFilter *self, HTASK callee,
                                              DWORD elapsed, DWORD pending_type)
{
    (void)self;
    (void)callee;
    (void)elapsed;
    (void)pending_type;
    return PENDINGMSG_WAITDEFPROCESS;
}

static IMessageFilterVtbl filter_vtable = {
    filter_query,
    filter_add_ref,
    filter_release,
    filter_incoming,
    filter_rejected,
    filter_pending
};

static IMessageFilter filter = {&filter_vtable};

HRESULT message_filter_register(IMessageFilter **previous)
{
    return CoRegisterMessageFilter(&filter, previous);
}

void message_filter_restore(IMessageFilter *previous)
{
    IMessageFilter *removed = NULL;
    CoRegisterMessageFilter(previous, &removed);
    if (removed) {
        IMessageFilter_Release(removed);
    }
    if (previous) {
        IMessageFilter_Release(previous);
    }
}
