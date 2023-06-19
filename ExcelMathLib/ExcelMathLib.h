// ExcelMathLib.h - contains declarations of functions for the excel library
#pragma once

#ifdef EXCELMATHLIB_EXPORTS
#define EXCELMATHLIB_API __declspec(dllexport)
#else
#define EXCELMATHLIB_API __declspec(dllimport)
#endif

extern "C" EXCELMATHLIB_API double get_square(double* x);
extern "C" EXCELMATHLIB_API double add_nums(double* x, double* y);
extern "C" EXCELMATHLIB_API double divide_nums(double* x, double* y);
extern "C" EXCELMATHLIB_API double add_two(double* x);
extern "C" EXCELMATHLIB_API double power(double* x, double* y);
extern "C" EXCELMATHLIB_API BSTR wordFunc();
extern "C" EXCELMATHLIB_API BSTR reverse_word(BSTR excelW);
extern "C" EXCELMATHLIB_API BSTR is_apple(BSTR excelW);
extern "C" EXCELMATHLIB_API BSTR bstr_str_back(BSTR excelW);
extern "C" EXCELMATHLIB_API BSTR greater_than_check(double* num, double* target);
extern "C" EXCELMATHLIB_API double add_byVal(double x, double y);
//extern "C" EXCELMATHLIB_API BSTR intToString(int* arr, int* size);
