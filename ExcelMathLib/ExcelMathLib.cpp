// MathLibrary.cpp : Defines the exported functions for the DLL.
#include "pch.h" // use stdafx.h in Visual Studio 2017 and earlier
#include <OleAuto.h> //must be included for BSTR to cstring conversions
#include <utility>
#include <limits.h>
#include "ExcelMathLib.h"
#include <string>

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
BSTR WINAPI wordFunc() //returns a VBA string/BSTR
{
    return SysAllocString(L"test");
}
*/

BSTR ConvertMBSToBSTR(const std::string& str) 
//found online at 
//https://stackoverflow.com/questions/6284524/bstr-to-stdstring-stdwstring-and-vice-versa 
//for converting wstrings to BSTRs
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

BSTR WINAPI wordFunc() // returns a VBA string of "test"
{
    return ConvertMBSToBSTR("test");
}