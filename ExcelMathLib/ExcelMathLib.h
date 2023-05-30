// MathLibrary.h - Contains declarations of math functions
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
