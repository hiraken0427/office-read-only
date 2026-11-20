#ifndef OFFICE_READ_ONLY_OFFICE_H
#define OFFICE_READ_ONLY_OFFICE_H

#include "dispatch.h"
#include "path.h"

typedef struct {
    IDispatch *applications[OFFICE_COUNT];
    bool restore_failed[OFFICE_COUNT];
} OfficeSession;

typedef enum {
    OPEN_FAILED,
    OPEN_NEW,
    OPEN_REUSED,
    OPEN_PROTECTED
} OpenStatus;

typedef struct {
    OpenStatus status;
    const wchar_t *reason;
    ComError error;
    ComError restore_error;
} OpenResult;

const wchar_t *office_name(OfficeKind kind);
OpenResult office_open(OfficeSession *session, OfficeKind kind, const FilePath *path);
void office_release(OfficeSession *session);

#endif
