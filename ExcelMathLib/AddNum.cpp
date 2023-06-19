//definitions for the AddNum (C++) class
#include "pch.h"
#include "AddNum.h"

AddNum::AddNum()
{
	num1 = 0.0;
	num2 = 0.0;
}

AddNum::AddNum(double x)
{
	num1 = 0.0 + x;
	num2 = 0.0 + x;
}

AddNum::AddNum(double x, double y)
{
	num1 = 0.0 + x;
	num2 = 0.0 + y;
}

double AddNum::addNums()
{
	return num1 + num2;
}

double AddNum::addNums(double num)
{
	return num1 + num;
}

double AddNum::sum(double x, double y)
{
	return x + y;
}

double AddNum::get_num1()
{
	return num1;
}

double AddNum::get_num2()
{
	return num2;
}
