#ifndef OFFICE_READ_ONLY_DISPATCH_H
#define OFFICE_READ_ONLY_DISPATCH_H

#include <objbase.h>
#include <oleauto.h>

typedef struct {
    const wchar_t *member;
    HRESULT hr;
    HRESULT exception_code;
    UINT argument;
    BSTR description;
} ComError;

void com_error_clear(ComError *error);
HRESULT com_error_set(ComError *error, const wchar_t *member, HRESULT hr);

HRESULT dispatch_invoke(IDispatch *object, const wchar_t *member, WORD flags,
                        const VARIANT *arguments, UINT count, VARIANT *result,
                        ComError *error);
HRESULT dispatch_object(IDispatch *object, const wchar_t *member, WORD flags,
                        const VARIANT *arguments, UINT count, IDispatch **result,
                        ComError *error);
HRESULT dispatch_property(IDispatch *object, const wchar_t *member, VARTYPE type,
                          VARIANT *result, ComError *error);
HRESULT dispatch_put(IDispatch *object, const wchar_t *member, VARIANT value,
                     ComError *error);

VARIANT variant_bool(BOOL value);
VARIANT variant_int(LONG value);
HRESULT message_filter_register(IMessageFilter **previous);
void message_filter_restore(IMessageFilter *previous);

#endif
