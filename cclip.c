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
    unsigned char *pInputBuffer;
    wchar_t *pUnicodeBuffer;
    unsigned int inputBufferSize;
    unsigned int unicodeBufferSize;
    unsigned int totalReadBytes;
    unsigned int codepage;
    int retval;

    inputBufferSize = 4096; // TODO buffer size command line switch

    if (standardin == INVALID_HANDLE_VALUE)
    {
        fprintf(stderr, "Could not open standard input handle\n");
        exit(1);
    }

    // TODO use GetFileType() to determine type
    //   FILE_TYPE_DISK --> invoked as cclip.exe < somefile.txt
    //   FILE_TYPE_CHAR --> invoked as cclip.exe
    //   FILE_TYPE_PIPE --> invoked as someprocess.exe | cclip.exe

    codepage = GetConsoleCP(); // TODO command line switch for input codepage override

    pInputBuffer = malloc(inputBufferSize);
    if (pInputBuffer == NULL)
    {
        fprintf(stderr, "Could not allocate input buffer\n");
        exit(1);
    }

    totalReadBytes = 0;
    while (1)
    {
        unsigned int readbytes;
        // TODO double buffer size for each call and read everyting into one buffer
        if (!ReadFile(standardin, pInputBuffer, inputBufferSize, &readbytes, NULL))
        {
            unsigned int err = GetLastError();
            if (err == ERROR_BROKEN_PIPE && GetFileType(standardin) == FILE_TYPE_PIPE)
            {
                break;
            }
            else
            {
                fprintf(stderr, "ReadFile() failed, GetLastError() = 0x%X\n", GetLastError());
                exit(1);
            }
        }

        if (readbytes == 0)
            break;

        totalReadBytes += readbytes;
    }

    unicodeBufferSize = totalReadBytes * sizeof(wchar_t);
    pUnicodeBuffer = malloc(unicodeBufferSize);
    if (pUnicodeBuffer == NULL)
    {
        fprintf(stderr, "Could not allocate output buffer\n");
        exit(1);
    }

    retval = MultiByteToWideChar(codepage, 0, pInputBuffer, totalReadBytes, pUnicodeBuffer,
        unicodeBufferSize / sizeof(wchar_t));
    if (retval == 0)
    {
        fprintf(stderr, "MultiByteToWideChar() failed, GetLastError() = %X\n", GetLastError());
        exit(1);
    }

    retval = WriteToClipboard(CF_UNICODETEXT, pUnicodeBuffer, retval * sizeof(wchar_t));
    if (retval != 0)
    {
        fprintf(stderr, "WriteToClipboard() returned %d\n", retval);
        exit(1);
    }
}
