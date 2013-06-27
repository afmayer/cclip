#include <windows.h>
#include <stdio.h>

/* ReadFileToBuffer()
 *
 * Allocate a buffer fill it with data from a given file handle. Returns the
 * address of the allocated buffer (which must be released by the caller) and
 * the number of read bytes. The size of the allocated buffer is not returned
 * but will be the next multiple of bufferSizeStep greater than the returned
 * number of read bytes.
 *
 * Returns zero on success. Returns -1 when the buffer could not be
 * (re-)allocated and -2 when reading from the handle fails. When an error
 * occurrs no buffer must be released by the caller and the values of the output
 * pointers are undefined.
 */
int ReadFileToBuffer(HANDLE fileHandle, unsigned int bufferSizeStep,
                     void **ppAllocatedBuffer, unsigned int *pReadBytes)
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
            return -1;

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

int WriteToClipboard(unsigned int format, const void *pData,
                     unsigned int sizeBytes)
{
    HGLOBAL hGlobalMem;
    unsigned char *pGlobalMem;

    if (!OpenClipboard(0))
        return -1;

    if (!EmptyClipboard())
    {
        DWORD err = GetLastError();
        CloseClipboard();
        return -2;
    }

    hGlobalMem = GlobalAlloc(GMEM_MOVEABLE, sizeBytes);
    if (hGlobalMem == NULL)
    {
        CloseClipboard();
        return -3;
    }

    pGlobalMem = GlobalLock(hGlobalMem);
    if (pGlobalMem == NULL)
    {
        GlobalFree(hGlobalMem);
        CloseClipboard();
        return -4;
    }

    CopyMemory(pGlobalMem, pData, sizeBytes);
    GlobalUnlock(hGlobalMem);

    if (!SetClipboardData(format, hGlobalMem))
    {
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

    retval = ReadFileToBuffer(standardin, inputBufferSizeStep, &pInputBuffer,
        &totalReadBytes);
    if (retval != 0)
    {
        fprintf(stderr, "ReadFileToBuffer() returned %d\n", retval);
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
        (unicodeBufferCharacters + 1) * sizeof(wchar_t));
    if (retval != 0)
    {
        fprintf(stderr, "WriteToClipboard() returned %d\n", retval);
        free(pUnicodeBuffer);
        exit(1);
    }

    free(pUnicodeBuffer);
}
