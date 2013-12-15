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
    TagTypePreWithAttributes,
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

/* SearchForStringList()
 *
 * Search for the first occurrence of any of an array of strings within a given
 * input string. Earlier entries in the array of search strings take higher
 * priority.
 *
 * Returns the character index (within the input string) of the first
 * occurrence of a found string. Returns -1 if none of the strings was found.
 * Stores the index of the search string within the given array of search
 * strings in an output variable if a string is found, otherwise the value of
 * the output pointer is undefined.
 */
int SearchForStringList(const wchar_t *pInputString,
                        unsigned int inputStringSizeBytes,
                        const wchar_t **ppSearchStrings,
                        unsigned int numberOfSearchStrings,
                        unsigned int *pHitSearchStringIndex)
{
    unsigned int i;
    unsigned int charPos = 0;

    for (charPos = 0; charPos < inputStringSizeBytes / sizeof(wchar_t);
        charPos++)
    {
        for (i = 0; i < numberOfSearchStrings; i++)
        {
            unsigned int cmpPos = 0;
            while (ppSearchStrings[i][cmpPos] == pInputString[charPos+cmpPos]
                && charPos + cmpPos < inputStringSizeBytes / sizeof(wchar_t))
            {
                cmpPos++;
            }
            if (ppSearchStrings[i][cmpPos] == L'\0')
            {
                /* found a search string */
                *pHitSearchStringIndex = i;
                return charPos;
            }
        }
    }

    /* none of the search strings was found */
    return -1;
}

void ShiftFormatInfoPositions(FormatInfo *pFormatInfo,
                              unsigned int startCharPos,
                              unsigned int charactersDeleted,
                              unsigned int charactersInserted)
{
    unsigned int i;

    for (i = 0; i < pFormatInfo->numberOfTags; i++)
    {
        if (pFormatInfo->tags[i].characterPos > startCharPos)
        {
            /* tags that are within a deleted region:
               move to the beginning of the deleted region */
            if (pFormatInfo->tags[i].characterPos <
                startCharPos + charactersDeleted)
            {
                pFormatInfo->tags[i].characterPos = startCharPos;
            }
            /* tags that are after a deleted region:
               move right when more characters are inserted than deleted, move
               left otherwise */
            else
            {
                pFormatInfo->tags[i].characterPos +=
                    charactersInserted - charactersDeleted;
            }
        }
    }
}

// TODO ReplaceCharacters() documentation
int ReplaceCharacters(const wchar_t *pInputBuffer,
                      unsigned int inputBufSizeBytes,
                      FormatInfo *pFormatInfo,
                      const wchar_t **ppSearchStrings,
                      const wchar_t **ppReplaceStrings,
                      wchar_t **ppAllocatedBuffer,
                      unsigned int *pAllocatedBufSizeBytes,
                      ErrBlock *pEb)
{
    wchar_t *pOutputBuffer;
    unsigned int numOfSearchStrings = 0;
    unsigned int outputCharacters = 0;
    unsigned int inputCharacterPos = 0;
    unsigned int outputCharPos = 0;

    while (ppSearchStrings[numOfSearchStrings] != NULL)
        numOfSearchStrings++;

    /* determine output size */
    while (1)
    {
        unsigned int hitSearchStringIndex;
        unsigned int searchStringCharacters;
        unsigned int replaceCharacters;
        int index;
        index = SearchForStringList(pInputBuffer + inputCharacterPos,
            inputBufSizeBytes - inputCharacterPos * sizeof(wchar_t),
            ppSearchStrings, numOfSearchStrings, &hitSearchStringIndex);
        if (index == -1)
        {
            outputCharacters +=
                inputBufSizeBytes / sizeof(wchar_t) - inputCharacterPos;
            break;
        }
        searchStringCharacters = wcslen(ppSearchStrings[hitSearchStringIndex]);
        replaceCharacters = wcslen(ppReplaceStrings[hitSearchStringIndex]);
        outputCharacters += index + replaceCharacters;
        inputCharacterPos += index + searchStringCharacters;
    }

    pOutputBuffer = malloc(outputCharacters * sizeof(wchar_t));
    if (pOutputBuffer == NULL)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "Could not allocate buffer for replacement text");
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 1;
        }
        return -1;
    }

    inputCharacterPos = 0;
    while (1)
    {
        unsigned int hitSearchStringIndex;
        unsigned int searchStringCharacters;
        unsigned int replaceCharacters;
        int index;
        index = SearchForStringList(pInputBuffer + inputCharacterPos,
            inputBufSizeBytes - inputCharacterPos * sizeof(wchar_t),
            ppSearchStrings, numOfSearchStrings, &hitSearchStringIndex);
        if (index == -1)
        {
            /* no more strings to replace - insert remaining input string */
            memcpy(pOutputBuffer + outputCharPos,
                pInputBuffer + inputCharacterPos,
                inputBufSizeBytes - inputCharacterPos * sizeof(wchar_t));
            outputCharPos +=
                inputBufSizeBytes / sizeof(wchar_t) - inputCharacterPos;
            break;
        }

        searchStringCharacters = wcslen(ppSearchStrings[hitSearchStringIndex]);
        replaceCharacters = wcslen(ppReplaceStrings[hitSearchStringIndex]);

        /* adapt FormatInfo structure tag positions */
        ShiftFormatInfoPositions(pFormatInfo, index + outputCharPos,
            searchStringCharacters, replaceCharacters);

        /* insert input string up to characters to be replaced */
        memcpy(pOutputBuffer + outputCharPos,
            pInputBuffer + inputCharacterPos,
            index * sizeof(wchar_t));
        outputCharPos += index;

        /* insert replacement string */
        memcpy(pOutputBuffer + outputCharPos,
            ppReplaceStrings[hitSearchStringIndex],
            replaceCharacters * sizeof(wchar_t));
        outputCharPos += replaceCharacters;
        inputCharacterPos += index + searchStringCharacters;
    }

    if (outputCharacters != outputCharPos)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "Error in internal buffer size calculation");
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 2;
        }
        free(pOutputBuffer);
        return -1;
    }

    /* success */
    *ppAllocatedBuffer = pOutputBuffer;
    *pAllocatedBufSizeBytes = outputCharacters * sizeof(wchar_t);
    return 0;
}

/* GenerateHtmlMarkupFromFormatInfoTag()
 *
 * Generate HTML code in UTF8 (without a zero termination byte) from a TagType,
 * a parameter and a boolean flag (indicating if the tag is a closing tag) and
 * store it at a specified address. When the specified buffer size is 0 no data
 * is written (for size calculation). Otherwise the function fails if the
 * specified buffer size if smaller than the generated HTML code.
 *
 * Returns the number of bytes required by the generated UTF8 string (without
 * a zero termination byte) on success or -1 in case of an error.
 */
int GenerateHtmlMarkupFromFormatInfoTag(TagType type, unsigned parameter,
                                        unsigned int yClose,
                                        char *pOutputBuffer,
                                        unsigned int bufferSizeBytes)
{
    int returnValue;
    char *pTag = NULL;
    unsigned int yFreeTagPointer = 0;

    if (type == TagTypePreWithAttributes)
    {
        if (!yClose)
            pTag = "<pre>";
        else
            pTag = "</pre>";
    }
    else if (type == TagTypeUnderscore)
    {
        if (!yClose)
            pTag = "<u>";
        else
            pTag = "</u>";
    }
    else if (type == TagTypeBgBlue)
    {
        // TODO more tag types in GenerateHtmlMarkupFromFormatInfoTag()
        pTag = "<blue>";
    }

    if (pTag == NULL)
        return -1;

    returnValue = strlen(pTag);

    if (bufferSizeBytes != 0)
    {
        if (bufferSizeBytes >= (unsigned int)returnValue)
            snprintf(pOutputBuffer, returnValue, pTag);
        else
            returnValue = -1;
    }

    if (yFreeTagPointer)
        free(pTag);

    return returnValue;

    //<span style="color: rgb(0, 0, 0);
    //    font-family: Verdana, Arial, Helvetica, sans-serif;
    //    font-size: medium; font-style: normal;
    //    font-variant: normal; font-weight: normal;
    //    letter-spacing: normal;
    //    line-height: 18px;
    //    orphans: auto;
    //    text-align: left;
    //    text-indent: 0px;
    //    text-transform: none;
    //    white-space: normal;
    //    widows: auto;
    //    word-spacing: 0px;
    //    -webkit-text-stroke-width: 0px;
    //    background-color: rgb(255, 255, 255);
    //    display: inline !important;
    //    float: none;">intention o</span>
}

/* GenerateClipboardHtml()
 *
 * Generate HTML code in the CF_HTML clipboard format (not zero terminated)
 * from a wide character input buffer (does not need to be zero terminated) and
 * an optional FormatInfo structure and store it in an allocated buffer. Stores
 * the address of the allocated buffer (which must be released by the caller)
 * and the size of the allocated buffer in output variables. When the
 * FormatInfo pointer is NULL no formatting is applied to the HTML output.
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
    char *pStartString = "Version:0.9\r\nStartHTML:0000000105\r\n"
        "EndHTML:0000000000\r\nStartFragment:0000000141\r\n"
        "EndFragment:0000000000\r\n<html>\r\n<body>\r\n<!--StartFragment-->";
    char *pEndString = "<!--EndFragment-->\r\n</body>\r\n</html>";
    char *pOutputBuffer;
    FormatInfo *pOwnFormatInfo;
    unsigned int formatInfoTotalTags;
    unsigned int formatInfoTagIndex = 0;
    unsigned int inputCharacterPos = 0;
    unsigned int nextTagSearchStartPos = 0;
    unsigned int outputBufWriteIndex = 0;
    unsigned int outputBufRemainingBytes;
    unsigned int htmlSizeBytes;
    unsigned int i;
    int iReturnedSize;
    // TODO zero terminate GenerateClipboardHtml() output?

    /* create FormatInfo structure derived from pFormatInfo parameter */
    formatInfoTotalTags = (
            /* tags from pFormatInfo input parameter */
            (pFormatInfo == NULL ? 0 : pFormatInfo->numberOfTags) +
            /* opening and closing <pre> tag */
            2);
    pOwnFormatInfo = malloc(
            /* space for FormatInfo structure without any tags */
            sizeof(*pFormatInfo) - sizeof(pFormatInfo->tags) +
            /* space for tags */
            formatInfoTotalTags * sizeof(pFormatInfo->tags));
    if (pOwnFormatInfo == NULL)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "Could not allocate buffer for FormatInfo structure");
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 1;
        }
        return -1;
    }
    pOwnFormatInfo->numberOfTags = formatInfoTotalTags;

    /* insert <pre> tag around everything (open) */
    pOwnFormatInfo->tags[formatInfoTagIndex].characterPos = 0;
    pOwnFormatInfo->tags[formatInfoTagIndex].type = TagTypePreWithAttributes;
    pOwnFormatInfo->tags[formatInfoTagIndex].parameter = 0;
    pOwnFormatInfo->tags[formatInfoTagIndex].yClose = 0;
    formatInfoTagIndex++;
    // TODO set attributes for <pre> tag

    /* copy tags from pFormatInfo input parameter */
    if (pFormatInfo != NULL)
    {
        memcpy(&(pOwnFormatInfo->tags[formatInfoTagIndex]), pFormatInfo->tags,
            pFormatInfo->numberOfTags * sizeof(pFormatInfo->tags));
        formatInfoTagIndex += pFormatInfo->numberOfTags;
    }

    /* insert <pre> tag around everything (close) */
    pOwnFormatInfo->tags[formatInfoTagIndex].characterPos =
        inputBufSizeBytes / sizeof(wchar_t);
    pOwnFormatInfo->tags[formatInfoTagIndex].type = TagTypePreWithAttributes;
    pOwnFormatInfo->tags[formatInfoTagIndex].parameter = 0;
    pOwnFormatInfo->tags[formatInfoTagIndex].yClose = 1;
    formatInfoTagIndex++;

    /* sanity check: FormatInfo structure filled correctly? */
    if (formatInfoTagIndex != formatInfoTotalTags)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "Error in internal FormatInfo size calculation");
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 2;
        }
        free(pOwnFormatInfo);
        return -1;
    }

    // TODO replace HTML characters, and optionally linefeeds with <br>
    //  "   ==>  &quot;     &#34;
    //  &   ==>  &amp;      &#38;
    //  <   ==>  &lt;       &#60;
    //  >   ==>  &gt;       &#62;

    /* determine output size: input string as UTF8 */
    iReturnedSize = WideCharToMultiByte(CP_UTF8, 0, pInputBuffer,
        inputBufSizeBytes / sizeof(wchar_t), NULL, 0, NULL, NULL);
    if (iReturnedSize == 0)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "WideCharToMultiByte() space detection failed, "
                "GetLastError() = 0x%X", GetLastError());
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 3;
        }
        free(pOwnFormatInfo);
        return -1;
    }
    htmlSizeBytes = (unsigned int)iReturnedSize;

    /* determine output size: description + fixed start and end HTML code */
    htmlSizeBytes += strlen(pStartString) + strlen(pEndString);

    /* determine output size: generated HTML tags */
    for (i = 0; i < pOwnFormatInfo->numberOfTags; i++)
    {
        iReturnedSize = GenerateHtmlMarkupFromFormatInfoTag(
            pOwnFormatInfo->tags[i].type,
            pOwnFormatInfo->tags[i].parameter,
            pOwnFormatInfo->tags[i].yClose,
            NULL,
            0);
        if (iReturnedSize == -1)
        {
            if (pEb != NULL)
            {
                snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                    "HTML tag space detection for tag type 0x%X with "
                    "parameter 0x%X failed", pOwnFormatInfo->tags[i].type,
                    pOwnFormatInfo->tags[i].parameter);
                pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
                pEb->functionSpecificErrorCode = 4;
            }
            free(pOwnFormatInfo);
            return -1;
        }
        htmlSizeBytes += (unsigned int)iReturnedSize;
    }

    /* allocate the output buffer */
    outputBufRemainingBytes = htmlSizeBytes;
    pOutputBuffer = malloc(htmlSizeBytes);
    if (pOutputBuffer == NULL)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "Could not allocate buffer for HTML data");
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 5;
        }
        free(pOwnFormatInfo);
        return -1;
    }

    /* fill buffer: description and HTML until <!--StartFragment--> */
    iReturnedSize = (int)strlen(pStartString);
    strncpy(pOutputBuffer + outputBufWriteIndex, pStartString, iReturnedSize);
    outputBufWriteIndex += iReturnedSize;
    outputBufRemainingBytes -= iReturnedSize;

    while (1)
    {
        unsigned int yFoundNextTag = 0;
        unsigned int nextTagCharacter;
        unsigned int inputCharsToConvert;

        nextTagCharacter = inputBufSizeBytes / sizeof(wchar_t);
        for (i = 0; i < pOwnFormatInfo->numberOfTags; i++)
        {
            if (pOwnFormatInfo->tags[i].characterPos <= nextTagCharacter &&
                pOwnFormatInfo->tags[i].characterPos >= nextTagSearchStartPos)
            {
                nextTagCharacter = pOwnFormatInfo->tags[i].characterPos;
                yFoundNextTag = 1;
            }
        }

        if (yFoundNextTag)
        {
            /* convert input characters up to next tag */
            inputCharsToConvert = nextTagCharacter - inputCharacterPos;
            nextTagSearchStartPos = nextTagCharacter + 1;
        }
        else
        {
            /* convert all remaining input characters */
            inputCharsToConvert =
                (inputBufSizeBytes / sizeof(wchar_t)) - inputCharacterPos;
        }

        /* fill buffer: convert input characters */
        if (inputCharsToConvert != 0)
        {
            iReturnedSize = WideCharToMultiByte(CP_UTF8, 0, pInputBuffer +
                    inputCharacterPos, inputCharsToConvert, pOutputBuffer +
                    outputBufWriteIndex, outputBufRemainingBytes, NULL, NULL);
            if (iReturnedSize == 0)
            {
                if (pEb != NULL)
                {
                    snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                        "WideCharToMultiByte() conversion failed, "
                        "GetLastError() = 0x%X", GetLastError());
                    pEb->errDescription[
                        sizeof(pEb->errDescription) - 1] = '\0';
                    pEb->functionSpecificErrorCode = 6;
                }
                free(pOutputBuffer);
                free(pOwnFormatInfo);
                return -1;
            }
            outputBufWriteIndex += iReturnedSize;
            outputBufRemainingBytes -= iReturnedSize;
        }
        inputCharacterPos += inputCharsToConvert;

        /* exit the loop if there are no more tags to write */
        if (!yFoundNextTag)
            break;

        /* fill buffer: insert all tags at that position */
        for (i = 0; i < pOwnFormatInfo->numberOfTags; i++)
        {
            if (pOwnFormatInfo->tags[i].characterPos == inputCharacterPos)
            {
                iReturnedSize = GenerateHtmlMarkupFromFormatInfoTag(
                        pOwnFormatInfo->tags[i].type,
                        pOwnFormatInfo->tags[i].parameter,
                        pOwnFormatInfo->tags[i].yClose,
                        pOutputBuffer + outputBufWriteIndex,
                        outputBufRemainingBytes);
                if (iReturnedSize == -1)
                {
                    if (pEb != NULL)
                    {
                        snprintf(pEb->errDescription,
                            sizeof(pEb->errDescription),
                            "HTML tag generation for tag type 0x%X with "
                            "parameter 0x%X failed",
                            pOwnFormatInfo->tags[i].type,
                            pOwnFormatInfo->tags[i].parameter);
                        pEb->errDescription[
                            sizeof(pEb->errDescription) - 1] = '\0';
                        pEb->functionSpecificErrorCode = 7;
                    }
                    free(pOutputBuffer);
                    free(pOwnFormatInfo);
                    return -1;
                }
                outputBufWriteIndex += iReturnedSize;
                outputBufRemainingBytes -= iReturnedSize;
            }
        }
    }

    /* the FormatInfo structure is not needed anymore */
    free(pOwnFormatInfo);

    /* fill buffer: <!--EndFragment--> and closing HTML tags */
    iReturnedSize = (int)strlen(pEndString);
    strncpy(pOutputBuffer + outputBufWriteIndex, pEndString, iReturnedSize);
    outputBufWriteIndex += iReturnedSize;
    outputBufRemainingBytes -= iReturnedSize;

    /* correct EndHTML and EndFragment in the description */
    snprintf(pOutputBuffer + 0x2B, 10, "%010d", htmlSizeBytes);
    snprintf(pOutputBuffer + 0x5D, 10, "%010d", htmlSizeBytes - 36);

    /* sanity check: buffer filled exactly to the end? */
    if (outputBufWriteIndex != htmlSizeBytes || outputBufRemainingBytes != 0)
    {
        if (pEb != NULL)
        {
            snprintf(pEb->errDescription, sizeof(pEb->errDescription),
                "Error in internal buffer size calculation");
            pEb->errDescription[sizeof(pEb->errDescription) - 1] = '\0';
            pEb->functionSpecificErrorCode = 8;
        }
        free(pOutputBuffer);
        return -1;
    }

    /* success */
    *ppAllocatedHtmlBuffer = pOutputBuffer;
    *pAllocatedHtmlBufSizeBytes = htmlSizeBytes;
    return 0;
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
