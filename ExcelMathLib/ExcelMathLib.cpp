// MathLibrary.cpp : Defines the exported functions for the DLL.
#include "pch.h" // use stdafx.h in Visual Studio 2017 and earlier
#include <OleAuto.h> //must be included for BSTR to cstring conversions
//#include <iostream>
#include <utility>
#include <algorithm>
#include <cstring>
#include <limits.h>
#include "ExcelMathLib.h"

/*
these includes are probably important.
*/
#include <string>
#include <string.h>
#include <wtypes.h>
#include <comutil.h> //must be included for the _bstr_t type to work
#include <comdef.h>
#include <Windows.h>

/* Following typedefs allow for ease of reading (I think). BSTR refer to 
OLECHAR pointers specifically*/

typedef WCHAR OLECHAR;
typedef OLECHAR* BSTR;
typedef BSTR* LPBSTR;

double WINAPI get_square(double* x) //returns square of a given number
{
    return *x * *x;
}

double WINAPI add_nums(double* x, double* y) //adds two numbers together
{
    return *x + *y;
}

double WINAPI divide_nums(double* x, double* y) 
//divides two numbers, with first input as dividend, and second number
//is the divisor
{
    return *x / *y;
}

double WINAPI add_two(double* x) //adds 2.0 to the given input x
{
    return *x + 2.0;
}

double WINAPI power(double* x, double* y) // raises x to the power of y.
{
    return pow(*x, *y);
}
/*
* Below are functions dealing with strings.
*/

/*
BSTR WINAPI wordFunc() //returns a VBA string/BSTR
{
    return SysAllocString(L"test");
}
*/

/*
* another converter found at
* https://gist.github.com/kuba-orlik/999ca634dba613ba6a1c
*/
/*
std::string bstr_to_str(BSTR source) {
    //source = L"lol2inside";
    _bstr_t wrapped_bstr = _bstr_t(source);
    int length = wrapped_bstr.length();
    char* char_array = new char[length];
    strcpy_s(char_array, length + 1, wrapped_bstr);
    return char_array;
}
*/

/*
* Helper function for ConvertBSTRToMBS. Probably do not call directly.
* original found at https://stackoverflow.com/questions/6284524/bstr-to-stdstring-stdwstring-and-vice-versa 
*/
std::string ConvertWCSToMBS(const wchar_t* pstr, long wslen)
{
    int len = ::WideCharToMultiByte(CP_ACP, 0, pstr, wslen, NULL, 0, NULL, NULL);

    std::string dblstr(len, '\0');
    len = ::WideCharToMultiByte(CP_ACP, 0 /* no flags */,
        pstr, wslen /* not necessary NULL-terminated */,
        &dblstr[0], len,
        NULL, NULL /* no default char */);

    return dblstr;
}

/*
* Call this function to convert a given BSTR into std::string
* ConvertWCSToMBS is a helper function and should probably not be called directly 
* by any other method
* function found at: https://stackoverflow.com/questions/6284524/bstr-to-stdstring-stdwstring-and-vice-versa 
*/
std::string ConvertBSTRToMBS(BSTR bstr)
{
    int wslen = ::SysStringLen(bstr);
    return ConvertWCSToMBS((wchar_t*)bstr, wslen);
}

BSTR ConvertMBSToBSTR(const std::string& str) 
//found online at 
//https://stackoverflow.com/questions/6284524/bstr-to-stdstring-stdwstring-and-vice-versa 
//for converting wstrings and strings to BSTRs
{
    int wslen = ::MultiByteToWideChar(CP_ACP, 0 /* no flags */,
        str.data(), str.length(),
        NULL, 0);

    BSTR wsdata = ::SysAllocStringLen(NULL, wslen);
    ::MultiByteToWideChar(CP_ACP, 0 /* no flags */,
        str.data(), str.length(),
        wsdata, wslen);
    return wsdata;
}

/*
* ChatGPT generated codes for BSTR to string conversion.
*/
std::string bstr_to_string(BSTR bstrVariable)
{
    std::wstring wstrVariable(bstrVariable);
    std::string strVariable(wstrVariable.begin(), wstrVariable.end());
    return strVariable;
}

std::string BstrToString(BSTR bstr)
{
    std::string result;

    int len = SysStringLen(bstr);
    if (len > 0)
    {
        int size = WideCharToMultiByte(CP_UTF8, 0, bstr, len, nullptr, 0, nullptr, nullptr);
        if (size > 0)
        {
            char* buffer = new char[size + 1];
            WideCharToMultiByte(CP_UTF8, 0, bstr, len, buffer, size, nullptr, nullptr);
            buffer[size] = '\0';
            result = buffer;
            delete[] buffer;
        }
    }

    return result;
}

/*
* Another BSTR to wstring converter, found at https://stackoverflow.com/questions/6284524/bstr-to-stdstring-stdwstring-and-vice-versa
* in the answers.
*/
std::wstring BSTR_to_wstring(BSTR bs)
{
    assert(bs != nullptr);
    std::wstring ws(bs, SysStringLen(bs));
    return ws;
}

BSTR WINAPI wordFunc() // returns a VBA string of "test"
{
    return ConvertMBSToBSTR("test");
}

BSTR WINAPI reverse_word(BSTR excelW)
{
    std::string str = ConvertBSTRToMBS(excelW);
    reverse(str.begin(), str.end());
    return ConvertMBSToBSTR(str);
}

BSTR WINAPI is_apple(BSTR excelW) 
{
    //std::string str = ConvertBSTRToMBS(excelW);
    //std::string str = bstr_to_str(excelW);
    //std::wstring str = BSTR_to_wstring(excelW);
    //std::string str = bstr_to_string(excelW);
    std::string str = BstrToString(excelW);
    //std::cout << "Converted string: " << str << std::endl;
    //if (str == L"apple") {
    if (str.compare("apple") == 0) {
        return ConvertMBSToBSTR("Yes");
    }
    return ConvertMBSToBSTR("No");
}

BSTR WINAPI bstr_str_back(BSTR excelW)
{
    //std::string str = bstr_to_string(excelW);
    std::string str = BstrToString(excelW);
    return ConvertMBSToBSTR(str);
}

/*
* Checks to see if num is greater than target. Returns "True" or "False" as BSTRs
*/
BSTR WINAPI greater_than_check(double* num, double* target)
{
    if (*num > *target) {
        return ConvertMBSToBSTR("True");
    }
    return ConvertMBSToBSTR("False");
}