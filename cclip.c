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

typedef struct ErrBlock_
{
    char errDescription[256];
} ErrBlock;

/* ReadFileToNewBuffer()
 *
 * Allocate a buffer and fill it with data from a given file handle. Returns
 * the address of the allocated buffer (which must be released by the caller)
 * and the number of read bytes. The size of the allocated buffer is not
 * returned but will be the next multiple of bufferSizeStep greater than the
 * returned number of read bytes.
 *
 * Returns zero on success and a negative number otherwise. In case of an error
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
                        snprintf(pEb->errDescription,
                            sizeof(pEb->errDescription),
                            "ReadFile() failed, GetLastError() = 0x%X",
                            GetLastError()); // TODO use FormatMessage()
                        pEb->errDescription[
                            sizeof(pEb->errDescription) - 1] = '\0';
                    }
                    return -2;
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
        }
        CloseClipboard();
        return -2;
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
        }
        CloseClipboard();
        return -3;
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
        }
        GlobalFree(hGlobalMem);
        CloseClipboard();
        return -4;
    }

    CopyMemory(pGlobalMem, pData, sizeBytes);
    GlobalUnlock(hGlobalMem);

    if (!SetClipboardData(format, hGlobalMem))
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "GlobalLock() failed, GetLastError() = 0x%X",
                GetLastError());
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
        }
        GlobalFree(hGlobalMem);
        CloseClipboard();
        return -5;
    }

    CloseClipboard();
    GlobalFree(hGlobalMem);

    return 0;
}

int main(int argc, char *argv[])
{
    HANDLE standardin = GetStdHandle(STD_INPUT_HANDLE);
    char *pInputBuffer;
    wchar_t *pUnicodeBuffer;
    unsigned int inputBufferSizeStep;
    unsigned int unicodeBufferCharacters;
    unsigned int totalReadBytes;
    unsigned int codepage;
    unsigned int fileType;
    int retval;
    ErrBlock eb;

    inputBufferSizeStep = 4096; // TODO buffer size step command line switch

    if (standardin == INVALID_HANDLE_VALUE)
    {
        fprintf(stderr, "Could not open standard input handle\n");
        exit(1);
    }

    codepage = GetConsoleCP();
    fileType = GetFileType(standardin);
    if (fileType == FILE_TYPE_DISK)
    {
        /* stdin is redirected to a file - use system default codepage */
        codepage = GetACP();
    }
    else if (fileType == FILE_TYPE_CHAR)
    {
        /* standard input is connected to a console */
        // TODO handle standard input connected to console
    }
    else if (fileType == FILE_TYPE_PIPE)
    {
        /* stdin is connected to a pipe - use console codepage */
    }
    else if (fileType == FILE_TYPE_UNKNOWN)
    {
        if (GetLastError() != NO_ERROR)
        {
            // TODO Handle error in GetFileType()
        }
        // TODO handle FILE_TYPE_UNKNOWN
    }

    // TODO command line switch for input codepage override
    //   allow default codepage identifiers CP_ACP, CP_OEMCP, ?CP_MACCP?,
    //   CP_THREAD_ACP, CP_SYMBOL, CP_UTF7 and CP_UTF8

    retval = ReadFileToNewBuffer(standardin, inputBufferSizeStep,
        &pInputBuffer, &totalReadBytes, &eb);
    if (retval != 0)
    {
        fprintf(stderr, "ERROR: ReadFileToNewBuffer() returned %d\n    %s\n",
            retval, eb.errDescription);
        exit(1);
    }

    unicodeBufferCharacters =
        MultiByteToWideChar(codepage, 0, pInputBuffer, totalReadBytes, NULL, 0);
    if (unicodeBufferCharacters == 0)
    {
        fprintf(stderr, "MultiByteToWideChar() failed, GetLastError() = %X\n",
            GetLastError());
        free(pInputBuffer);
        exit(1);
    }
    pUnicodeBuffer = malloc((unicodeBufferCharacters + 1) * sizeof(wchar_t));
    if (pUnicodeBuffer == NULL)
    {
        fprintf(stderr, "Could not allocate Unicode conversion buffer\n");
        free(pInputBuffer);
        exit(1);
    }

    retval = MultiByteToWideChar(codepage, 0, pInputBuffer, totalReadBytes,
        pUnicodeBuffer, unicodeBufferCharacters);
    if (retval == 0)
    {
        fprintf(stderr, "MultiByteToWideChar() failed, GetLastError() = %X\n",
            GetLastError());
        free(pUnicodeBuffer);
        free(pInputBuffer);
        exit(1);
    }
    pUnicodeBuffer[unicodeBufferCharacters] = L'\0';

    free(pInputBuffer);

    retval = WriteToClipboard(CF_UNICODETEXT, pUnicodeBuffer,
        (unicodeBufferCharacters + 1) * sizeof(wchar_t), &eb);
    if (retval != 0)
    {
        fprintf(stderr, "ERROR: WriteToClipboard() returned %d\n    %s\n",
            retval, eb.errDescription);
        free(pUnicodeBuffer);
        exit(1);
    }

    free(pUnicodeBuffer);
}
