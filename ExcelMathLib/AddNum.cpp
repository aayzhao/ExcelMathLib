#include "pch.h"
#include "AddNum.h"

AddNum::AddNum(int x)
{
	num1 = 0.0 + x;
	num2 = 0.0 + x;
}

double AddNum::addNums()
{
	return num1 + num2;
}
