#include <windows.h>
#include <stdio.h>

int WriteToClipboard(unsigned int format, const void *pData, unsigned int sizeBytes)
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
    unsigned int inputBufferSize;
    unsigned int inputBufferSizeStep;
    unsigned int unicodeBufferCharacters;
    unsigned int totalReadBytes;
    unsigned int codepage;
    unsigned int fileType;
    int retval;

    inputBufferSize = 0;
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
        /* standard input is redirected to a file - use system default codepage */
        codepage = GetACP();
    }
    else if (fileType == FILE_TYPE_CHAR)
    {
        /* standard input is connected to a console */
        // TODO handle standard input connected to console
    }
    else if (fileType == FILE_TYPE_PIPE)
    {
        /* standard input is connected to a pipe - use process console codepage */
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

    pInputBuffer = NULL;
    totalReadBytes = 0;
    while (1)
    {
        unsigned int readbytes;
        char *pReadPointer;

        pInputBuffer = realloc(pInputBuffer, inputBufferSize + inputBufferSizeStep);
        pReadPointer = pInputBuffer + inputBufferSize;
        inputBufferSize += inputBufferSizeStep;
        if (pInputBuffer == NULL)
        {
            fprintf(stderr, "Could not allocate input buffer\n");
            exit(1);
        }

        // TODO double buffer size for each call and read everyting into one buffer
        if (!ReadFile(standardin, pReadPointer, inputBufferSizeStep, &readbytes, NULL))
        {
            unsigned int err = GetLastError();
            if (err == ERROR_BROKEN_PIPE && GetFileType(standardin) == FILE_TYPE_PIPE)
            {
                break;
            }
            else
            {
                fprintf(stderr, "ReadFile() failed, GetLastError() = 0x%X\n", GetLastError());
                free(pInputBuffer);
                exit(1);
            }
        }

        if (readbytes == 0)
            break;

        totalReadBytes += readbytes;
    }

    unicodeBufferCharacters =
        MultiByteToWideChar(codepage, 0, pInputBuffer, totalReadBytes, NULL, 0);
    if (unicodeBufferCharacters == 0)
    {
        fprintf(stderr, "MultiByteToWideChar() failed, GetLastError() = %X\n", GetLastError());
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
        fprintf(stderr, "MultiByteToWideChar() failed, GetLastError() = %X\n", GetLastError());
        free(pUnicodeBuffer);
        free(pInputBuffer);
        exit(1);
    }
    pUnicodeBuffer[unicodeBufferCharacters] = L'\0';

    free(pInputBuffer);

    retval = WriteToClipboard(CF_UNICODETEXT, pUnicodeBuffer, retval * sizeof(wchar_t));
    if (retval != 0)
    {
        fprintf(stderr, "WriteToClipboard() returned %d\n", retval);
        free(pUnicodeBuffer);
        exit(1);
    }

    free(pUnicodeBuffer);
}
