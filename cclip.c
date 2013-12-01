/* cclip
 * A tool for the Microsoft Windows clipboard
 * Copyright (c) 2013 Alexander F. Mayer
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 * GNU General Public License for more details.
 *
 * You should have received a copy of the GNU General Public License
 * along with this program.  If not, see <http://www.gnu.org/licenses/>.
 */

#ifdef _MSC_VER
#define _CRT_SECURE_NO_WARNINGS
#define snprintf _snprintf
#endif /* #ifdef _MSC_VER */

#include <windows.h>
#include <stdio.h>
#include <string.h>

typedef struct ErrBlock_
{
    unsigned int functionSpecificErrorCode;
    char errDescription[256];
} ErrBlock;

typedef enum TagType_
{
    TagTypeUnderscore,
    TagTypeFgBlue,
    TagTypeFgGreen,
    TagTypeFgRed,
    TagTypeBgBlue,
    TagTypeBgGreen,
    TagTypeBgRed
    // TODO extend TagType enum values
} TagType;

typedef struct FormatInfo_
{
    unsigned int numberOfTags;
    struct
    {
        unsigned int characterPos;
        TagType type;
        unsigned int parameter;
        unsigned int yClose;
    } tags[1];
} FormatInfo;

void ShowUsage(char *pArgv0)
{
    // TODO implement ShowUsage();
}

typedef struct CmdLineOptions_
{
    unsigned int yCodepageOverride;
    unsigned int codepage;
    unsigned int yInputBufferSizeStepOverride;
    unsigned int inputBufferSizeStep;
} CmdLineOptions;

void ParseCommandLineOptions(int argc, const char *argv[],
                             CmdLineOptions *pOptions)
{
    int i;

    memset(pOptions, 0, sizeof(*pOptions));
    for (i = 1; i < argc; i++)
    {
        if (strncmp(argv[i], "-c", 3) == 0 ||
            strncmp(argv[i], "-cp", 4) == 0 ||
            strncmp(argv[i], "-codepage", 10) == 0)
        {
            if (argc > i+1)
            {
                int val;
                // TODO CHECK FOR CP_ACP, CP_OEMCP, ?CP_MACCP?,
                //   CP_THREAD_ACP, CP_SYMBOL, CP_UTF7 and CP_UTF8
                i++;
                val = strtol(argv[i], NULL, 0);
                if (val >= 0)
                {
                    pOptions->codepage = (unsigned int)val;
                    pOptions->yCodepageOverride = 1;
                }
            }
        }
        else if (strncmp(argv[i], "-bufstep", 9) == 0)
        {
            if (argc > i+1)
            {
                int val;
                i++;
                val = strtol(argv[i], NULL, 0);
                if (val > 0)
                {
                    pOptions->inputBufferSizeStep = (unsigned int)val;
                    pOptions->yInputBufferSizeStepOverride = 1;
                }
            }
        }
        else
        {
            // TODO handle unsupported command line switches
        }
    }
}

/* ReadFileToNewBuffer()
 *
 * Allocate a buffer and fill it with data from a given file handle. Stores
 * the address of the allocated buffer (which must be released by the caller)
 * and the number of read bytes in output variables. The size of the allocated
 * buffer is not returned but will be the next multiple of bufferSizeStep
 * greater than or equal to the returned number of read bytes.
 *
 * Returns zero on success or -1 in case of an error. In case of an error
 * and when the error block pointer is not NULL the error block is filled with
 * an error description. When an error occurrs no buffer must be released by
 * the caller and the values of the output pointers are undefined.
 */
int ReadFileToNewBuffer(HANDLE fileHandle, unsigned int bufferSizeStep,
                     void **ppAllocatedBuffer, unsigned int *pReadBytes,
                     ErrBlock *pEb)
{
    char *pInputBuffer = NULL;
    unsigned int inputBufferSize = 0;
    unsigned int yEndOfInput = 0;
    unsigned int totalReadBytes = 0;

    while (!yEndOfInput)
    {
        unsigned int bufferRemainingBytes;
        char *pReadPointer;

        pInputBuffer = realloc(pInputBuffer, inputBufferSize + bufferSizeStep);
        pReadPointer = pInputBuffer + inputBufferSize;
        inputBufferSize += bufferSizeStep;
        bufferRemainingBytes = bufferSizeStep;
        if (pInputBuffer == NULL)
        {
            if (pEb != NULL)
            {
                snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                    "Could not allocate memory for input buffer");
                pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
                pEb->functionSpecificErrorCode = 1;
            }
            return -1;
        }

        do
        {
            unsigned int readBytes;
            if (!ReadFile(fileHandle, pReadPointer, bufferRemainingBytes,
                &readBytes, NULL))
            {
                unsigned int err = GetLastError();
                if (err == ERROR_BROKEN_PIPE)
                {
                    yEndOfInput = 1;
                    break;
                }
                else
                {
                    free(pInputBuffer);
                    if (pEb != NULL)
                    {
                        // TODO multiple ErrBlock types + function for output
                        //  - description in errDescription[]
                        //  - GetLastError() number stored, later printed via FormatMessageW()
                        //  - more?
                        snprintf(pEb->errDescription,
                            sizeof(pEb->errDescription),
                            "ReadFile() failed, GetLastError() = 0x%X",
                            GetLastError());
                        pEb->errDescription[
                            sizeof(pEb->errDescription) - 1] = '\0';
                        pEb->functionSpecificErrorCode = 2;
                    }
                    return -1;
                }
            }

            pReadPointer += readBytes;
            bufferRemainingBytes -= readBytes;

            if (readBytes == 0)
            {
                yEndOfInput = 1;
                break;
            }
        } while (bufferRemainingBytes != 0);

        totalReadBytes += (bufferSizeStep - bufferRemainingBytes);
    }

    /* success */
    *ppAllocatedBuffer = pInputBuffer;
    *pReadBytes = totalReadBytes;
    return 0;
}

// TODO change ConvToZeroTerminatedWideCharNewBuffer() to not include a NUL terminator character, adapt main()
/* ConvToZeroTerminatedWideCharNewBuffer()
 *
 * Convert a given string (not zero terminated) in a given codepage to a wide
 * character string and store it in an allocated buffer, adding a zero
 * termination character. Stores the address of the allocated buffer (which
 * must be released by the caller) and the size of the allocated buffer in
 * output variables.
 *
 * Returns zero on success or -1 in case of an error. In case of an error
 * and when the error block pointer is not NULL the error block is filled with
 * an error description. When an error occurrs no buffer must be released by
 * the caller and the values of the output pointers are undefined.
 */
int ConvToZeroTerminatedWideCharNewBuffer(const char *pInputBuffer,
                                          unsigned int numberOfInputBytes,
                                          unsigned int codepage,
                                          wchar_t **ppAllocatedWideCharBuffer,
                                          unsigned int *pAllocatedBufSizeBytes,
                                          ErrBlock *pEb)
{
    int retval;
    int numberOfWideCharacters;
    wchar_t *pWideCharBuf;

    if (numberOfInputBytes != 0)
    {
        numberOfWideCharacters = MultiByteToWideChar(codepage, 0, pInputBuffer,
            numberOfInputBytes, NULL, 0);
        if (numberOfWideCharacters == 0)
        {
            if (pEb != NULL)
            {
                snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                    "MultiByteToWideChar() space detection failed, "
                    "GetLastError() = 0x%X", GetLastError());
                pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
                pEb->functionSpecificErrorCode = 1;
            }
            return -1;
        }
    }
    else
    {
        numberOfWideCharacters = 0;
    }

    pWideCharBuf = malloc((numberOfWideCharacters + 1) * sizeof(wchar_t));
    if (pWideCharBuf == NULL)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "Could not allocate conversion buffer");
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 2;
        }
        return -1;
    }

    if (numberOfInputBytes != 0)
    {
        retval = MultiByteToWideChar(codepage, 0, pInputBuffer,
            numberOfInputBytes, pWideCharBuf, numberOfWideCharacters);
        if (retval == 0)
        {
            if (pEb != NULL)
            {
                snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                    "MultiByteToWideChar() conversion failed, GetLastError() "
                    "= 0x%X", GetLastError());
                pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
                pEb->functionSpecificErrorCode = 3;
            }
            free(pWideCharBuf);
            return -1;
        }
    }
    pWideCharBuf[numberOfWideCharacters] = L'\0';

    /* success */
    *ppAllocatedWideCharBuffer = pWideCharBuf;
    *pAllocatedBufSizeBytes = (numberOfWideCharacters + 1) * sizeof(wchar_t);
    return 0;
}

// TODO documentation for WriteToClipboard()
int WriteToClipboard(unsigned int format, const void *pData,
                     unsigned int sizeBytes, ErrBlock *pEb)
{
    HGLOBAL hGlobalMem;
    unsigned char *pGlobalMem;

    if (!OpenClipboard(0))
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "OpenClipboard() failed, GetLastError() = 0x%X",
                GetLastError());
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 1;
        }
        return -1;
    }

    if (!EmptyClipboard())
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "EmptyClipboard() failed, GetLastError() = 0x%X",
                GetLastError());
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 2;
        }
        CloseClipboard();
        return -1;
    }

    hGlobalMem = GlobalAlloc(GMEM_MOVEABLE, sizeBytes);
    if (hGlobalMem == NULL)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "GlobalAlloc() failed, GetLastError() = 0x%X",
                GetLastError());
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 3;
        }
        CloseClipboard();
        return -1;
    }

    pGlobalMem = GlobalLock(hGlobalMem);
    if (pGlobalMem == NULL)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "GlobalLock() failed, GetLastError() = 0x%X",
                GetLastError());
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 4;
        }
        GlobalFree(hGlobalMem);
        CloseClipboard();
        return -1;
    }

    CopyMemory(pGlobalMem, pData, sizeBytes);
    GlobalUnlock(hGlobalMem);

    if (!SetClipboardData(format, hGlobalMem))
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "SetClipboardData() failed, GetLastError() = 0x%X",
                GetLastError());
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 5;
        }
        GlobalFree(hGlobalMem);
        CloseClipboard();
        return -1;
    }

    CloseClipboard();
    GlobalFree(hGlobalMem);

    return 0;
}

/* GenerateClipboardHtml()
 *
 * Generate HTML code in the CF_HTML clipboard format (not NUL terminated) from
 * a wide character input buffer and an optional FormatInfo structure and store
 * it in an allocated buffer. Stores the address of the allocated buffer (which
 * must be released by the caller) and the size of the allocated buffer in
 * output variables. When the FormatInfo pointer is NULL no formatting is
 * applied to the HTML output.
 *
 * Returns zero on success or -1 in case of an error. In case of an error
 * and when the error block pointer is not NULL the error block is filled with
 * an error description. When an error occurrs no buffer must be released by
 * the caller and the values of the output pointers are undefined.
 */
int GenerateClipboardHtml(const wchar_t *pInputBuffer,
                          unsigned int inputBufSizeBytes,
                          const FormatInfo *pFormatInfo,
                          char **ppAllocatedHtmlBuffer,
                          unsigned int *pAllocatedHtmlBufSizeBytes,
                          ErrBlock *pEb)
{
    // TODO implement GenerateClipboardHtml()
    return -1;
}

int main(int argc, char *argv[])
{
    HANDLE standardin = GetStdHandle(STD_INPUT_HANDLE);
    char *pInputBuffer;
    wchar_t *pWideCharBuf;
    unsigned int inputBufferSizeStep;
    unsigned int totalReadBytes;
    unsigned int wideCharBufSizeBytes;
    unsigned int codepage;
    int retval;
    ErrBlock eb;
    CmdLineOptions opt;

    ParseCommandLineOptions(argc, argv, &opt);

    if (opt.yInputBufferSizeStepOverride)
        inputBufferSizeStep = opt.inputBufferSizeStep;
    else
        inputBufferSizeStep = 4096;

    if (standardin == INVALID_HANDLE_VALUE)
    {
        fprintf(stderr, "Could not open standard input handle\n");
        exit(1);
    }

    if (opt.yCodepageOverride)
    {
        codepage = opt.codepage;
    }
    else
    {
        unsigned int fileType = GetFileType(standardin);
        if (fileType == FILE_TYPE_DISK)
        {
            /* stdin is redirected to a file - use system default codepage */
            codepage = GetACP();
        }
        else if (fileType == FILE_TYPE_CHAR)
        {
            /* standard input is connected to a console */
            codepage = GetConsoleCP();
            // TODO handle standard input connected to console
        }
        else if (fileType == FILE_TYPE_PIPE)
        {
            /* stdin is connected to a pipe - use console codepage */
            codepage = GetConsoleCP();
        }
        else if (fileType == FILE_TYPE_UNKNOWN)
        {
            if (GetLastError() != NO_ERROR)
            {
                // TODO Handle error in GetFileType()
            }
            // TODO handle FILE_TYPE_UNKNOWN
            codepage = GetConsoleCP();
        }
    }

    retval = ReadFileToNewBuffer(standardin, inputBufferSizeStep,
        &pInputBuffer, &totalReadBytes, &eb);
    if (retval != 0)
    {
        fprintf(stderr, "ERROR: ReadFileToNewBuffer() returned %d\n    %s\n",
            retval, eb.errDescription);
        exit(1);
    }

    retval = ConvToZeroTerminatedWideCharNewBuffer(pInputBuffer,
        totalReadBytes, codepage, &pWideCharBuf, &wideCharBufSizeBytes, &eb);
    if (retval != 0)
    {
        fprintf(stderr, "ERROR: ConvToZeroTerminatedWideCharNewBuffer() "
            "returned %d\n    %s\n", retval, eb.errDescription);
        free(pInputBuffer);
        exit(1);
    }

    free(pInputBuffer);

    retval = WriteToClipboard(CF_UNICODETEXT, pWideCharBuf,
        wideCharBufSizeBytes, &eb);
    if (retval != 0)
    {
        fprintf(stderr, "ERROR: WriteToClipboard() returned %d\n    %s\n",
            retval, eb.errDescription);
        free(pWideCharBuf);
        exit(1);
    }

    free(pWideCharBuf);
}
