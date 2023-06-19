#include "pch.h" // use stdafx.h in Visual Studio 2017 and earlier
#include <stdlib.h>
#include <OleAuto.h> //must be included for BSTR to cstring conversions
//#include <iostream>
#include <utility>
#include <algorithm>
#include <cstring>
#include <limits.h>
#include "adder.h"
#include "AddNum.h"
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

struct adder {
	void* obj;
};

void* WINAPI adder_create()
{
	void* a;
	AddNum obj;
	a = &obj;

	return a;
}

adder_t* WINAPI adder_create_2(double start)
{
	adder_t* a;
	AddNum* obj;

	a = (decltype(a))malloc(sizeof(*a));
	obj = new AddNum(start);
	a->obj = obj;

	return a;
}

adder_t* WINAPI adder_create_3(double start1, double start2)
{
	adder_t* a;
	AddNum* obj;

	a = (decltype(a))malloc(sizeof(*a));
	obj = new AddNum(start1, start2);
	a->obj = obj;

	return a;
}

void WINAPI adder_destroy(void* a)
{
	if (a == NULL)
		return;
	delete (AddNum*)a;
}

double WINAPI adder_sum(void* a, double x, double y)
{
	AddNum* obj = (AddNum*)a;
	if (obj == NULL)
		return -1.0;
	obj->sum(x, y);
}