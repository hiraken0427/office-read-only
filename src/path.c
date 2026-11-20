#include "path.h"

#include <stdlib.h>
#include <string.h>
#include <wchar.h>

static bool is_separator(wchar_t ch)
{
    return ch == L'\\' || ch == L'/';
}

static bool drive_prefix(const wchar_t *name)
{
    return ((name[0] >= L'A' && name[0] <= L'Z') ||
            (name[0] >= L'a' && name[0] <= L'z')) && name[1] == L':';
}

static bool filesystem_syntax(const wchar_t *name)
{
    if (!*name) {
        return false;
    }

    /* コロンはドライブ文字の直後だけに許可し、代替データストリームを除外する。 */
    const wchar_t *tail = drive_prefix(name) ? name + 2 : name;
    if (wcspbrk(tail, L":*?\"<>|")) {
        return false;
    }

    if (is_separator(name[0]) && is_separator(name[1]) &&
        (name[2] == L'.' || name[2] == L'?') && is_separator(name[3])) {
        return false;
    }
    return wcsncmp(name, L"\\??\\", 4) != 0;
}

static bool is_absolute(const wchar_t *name)
{
    return (drive_prefix(name) && is_separator(name[2])) ||
           (is_separator(name[0]) && is_separator(name[1]));
}

const wchar_t *path_basename(const wchar_t *path)
{
    const wchar_t *result = path;
    for (; *path; ++path) {
        if (is_separator(*path)) {
            result = path + 1;
        }
    }
    return result;
}

static DWORD full_path(const wchar_t *input, wchar_t **result)
{
    if (!filesystem_syntax(input)) {
        return ERROR_INVALID_NAME;
    }

    DWORD capacity = GetFullPathNameW(input, 0, NULL, NULL);
    while (capacity) {
        wchar_t *buffer = malloc((size_t)capacity * sizeof(*buffer));
        if (!buffer) {
            return ERROR_NOT_ENOUGH_MEMORY;
        }
        DWORD size = GetFullPathNameW(input, capacity, buffer, NULL);
        if (!size) {
            DWORD error = GetLastError();
            free(buffer);
            return error;
        }
        if (size < capacity) {
            if (!filesystem_syntax(buffer) || !is_absolute(buffer)) {
                free(buffer);
                return ERROR_INVALID_NAME;
            }
            *result = buffer;
            return ERROR_SUCCESS;
        }
        free(buffer);
        capacity = size;
    }
    return GetLastError();
}

static wchar_t *final_path(HANDLE handle, DWORD volume_format)
{
    DWORD capacity = GetFinalPathNameByHandleW(handle, NULL, 0, volume_format);
    if (!capacity) {
        return NULL;
    }
    wchar_t *buffer = malloc((size_t)capacity * sizeof(*buffer));
    if (!buffer) {
        return NULL;
    }
    DWORD size = GetFinalPathNameByHandleW(handle, buffer, capacity, volume_format);
    if (!size || size >= capacity) {
        free(buffer);
        return NULL;
    }
    return buffer;
}

static void read_identity(FilePath *path)
{
    HANDLE handle = CreateFileW(path->absolute, FILE_READ_ATTRIBUTES,
                                FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE,
                                NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);
    if (handle == INVALID_HANDLE_VALUE) {
        return;
    }

    path->has_id = GetFileInformationByHandleEx(handle, FileIdInfo, &path->id,
                                               (DWORD)sizeof(path->id)) != FALSE;
    path->final_name = final_path(handle, VOLUME_NAME_GUID);
    if (!path->final_name) {
        path->final_name = final_path(handle, VOLUME_NAME_DOS);
    }
    CloseHandle(handle);
}

DWORD path_prepare(const wchar_t *input, FilePath *path, OfficeKind *kind)
{
    static const struct {
        const wchar_t *extension;
        OfficeKind kind;
    } formats[] = {
        {L".xlsx", OFFICE_EXCEL}, {L".xlsm", OFFICE_EXCEL},
        {L".xlsb", OFFICE_EXCEL}, {L".xls", OFFICE_EXCEL},
        {L".docx", OFFICE_WORD}, {L".docm", OFFICE_WORD}, {L".doc", OFFICE_WORD},
        {L".pptx", OFFICE_POWERPOINT}, {L".pptm", OFFICE_POWERPOINT},
        {L".ppt", OFFICE_POWERPOINT}
    };

    *kind = OFFICE_COUNT;
    DWORD error = full_path(input, &path->absolute);
    if (error != ERROR_SUCCESS) {
        return error;
    }

    DWORD attributes = GetFileAttributesW(path->absolute);
    if (attributes == INVALID_FILE_ATTRIBUTES) {
        return GetLastError();
    }
    if (attributes & (FILE_ATTRIBUTE_DIRECTORY | FILE_ATTRIBUTE_DEVICE)) {
        return ERROR_DIRECTORY;
    }

    const wchar_t *extension = wcsrchr(path_basename(path->absolute), L'.');
    if (extension) {
        for (size_t i = 0; i < ARRAYSIZE(formats); ++i) {
            if (CompareStringOrdinal(extension, -1, formats[i].extension, -1,
                                     TRUE) == CSTR_EQUAL) {
                *kind = formats[i].kind;
                read_identity(path);
                return ERROR_SUCCESS;
            }
        }
    }
    return ERROR_NOT_SUPPORTED;
}

void path_dispose(FilePath *path)
{
    free(path->final_name);
    free(path->absolute);
    *path = (FilePath){0};
}

static int volume_length(const wchar_t *path)
{
    if (!path) {
        return 0;
    }
    if (wcsncmp(path, L"\\\\?\\Volume{", 11) == 0) {
        const wchar_t *end = wcschr(path, L'}');
        return end ? (int)(end + 1 - path) : 0;
    }
    if (wcsncmp(path, L"\\\\?\\UNC\\", 8) == 0) {
        const wchar_t *share = wcschr(path + 8, L'\\');
        const wchar_t *end = share ? wcschr(share + 1, L'\\') : NULL;
        return end ? (int)(end - path) : 0;
    }
    if (wcsncmp(path, L"\\\\?\\", 4) == 0 && drive_prefix(path + 4)) {
        return 6;
    }
    return 0;
}

static bool names_equal(const wchar_t *first, const wchar_t *second, BOOL ignore_case)
{
    return first && second &&
           CompareStringOrdinal(first, -1, second, -1, ignore_case) == CSTR_EQUAL;
}

static PathMatch compare_paths(const FilePath *first, const FilePath *second)
{
    int first_volume = volume_length(first->final_name);
    int second_volume = volume_length(second->final_name);
    bool same_volume = first_volume && first_volume == second_volume &&
                       CompareStringOrdinal(first->final_name, first_volume,
                                            second->final_name, second_volume,
                                            TRUE) == CSTR_EQUAL;

    if (same_volume && first->has_id && second->has_id) {
        return first->id.VolumeSerialNumber == second->id.VolumeSerialNumber &&
               memcmp(&first->id.FileId, &second->id.FileId,
                      sizeof(first->id.FileId)) == 0 ? PATH_SAME : PATH_DIFFERENT;
    }
    if (names_equal(first->final_name, second->final_name, FALSE) ||
        names_equal(first->absolute, second->absolute, FALSE)) {
        return PATH_SAME;
    }
    if (names_equal(first->final_name, second->final_name, TRUE) ||
        names_equal(first->absolute, second->absolute, TRUE)) {
        return PATH_UNKNOWN;
    }

    /* 共有先と割り当てドライブは、異なるパスで同じファイルを指すことがある。 */
    if (!same_volume &&
        names_equal(path_basename(first->absolute), path_basename(second->absolute),
                    TRUE)) {
        if (first_volume && second_volume &&
            wcsncmp(first->final_name, L"\\\\?\\Volume{", 11) == 0 &&
            wcsncmp(second->final_name, L"\\\\?\\Volume{", 11) == 0) {
            return PATH_DIFFERENT;
        }
        return PATH_UNKNOWN;
    }
    return PATH_DIFFERENT;
}

PathMatch path_compare(const FilePath *path, const wchar_t *candidate)
{
    FilePath other = {0};
    if (!candidate || !is_absolute(candidate) ||
        full_path(candidate, &other.absolute) != ERROR_SUCCESS) {
        return PATH_UNKNOWN;
    }

    read_identity(&other);
    PathMatch match = compare_paths(path, &other);
    path_dispose(&other);
    return match;
}

PathMatch path_compare_source(const FilePath *path, const wchar_t *folder,
                              const wchar_t *name)
{
    if (name && is_absolute(name)) {
        return path_compare(path, name);
    }
    if (!folder || !*folder || !name || !*name || path_basename(name) != name) {
        return PATH_UNKNOWN;
    }

    size_t folder_length = wcslen(folder);
    size_t size = folder_length + wcslen(name) + 2;
    wchar_t *combined = malloc(size * sizeof(*combined));
    if (!combined) {
        return PATH_UNKNOWN;
    }
    wcscpy(combined, folder);
    if (!is_separator(folder[folder_length - 1])) {
        wcscat(combined, L"\\");
    }
    wcscat(combined, name);
    PathMatch match = path_compare(path, combined);
    free(combined);
    return match;
}
