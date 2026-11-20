#ifndef OFFICE_READ_ONLY_PATH_H
#define OFFICE_READ_ONLY_PATH_H

#include <stdbool.h>
#include <windows.h>

typedef enum {
    OFFICE_EXCEL,
    OFFICE_WORD,
    OFFICE_POWERPOINT,
    OFFICE_COUNT
} OfficeKind;

typedef struct {
    wchar_t *absolute;
    wchar_t *final_name;
    FILE_ID_INFO id;
    bool has_id;
} FilePath;

typedef enum { PATH_DIFFERENT, PATH_SAME, PATH_UNKNOWN } PathMatch;

DWORD path_prepare(const wchar_t *input, FilePath *path, OfficeKind *kind);
void path_dispose(FilePath *path);
const wchar_t *path_basename(const wchar_t *path);
PathMatch path_compare(const FilePath *path, const wchar_t *candidate);
PathMatch path_compare_source(const FilePath *path, const wchar_t *folder,
                              const wchar_t *name);

#endif
