//C++ class AddNum will be wrapped as "adder" for export
#pragma once
class AddNum
{
public: 
	AddNum();
	AddNum(double x);
	AddNum(double x, double y);
	double addNums();
	double addNums(double num);
	double sum(double x, double y);
	double get_num1();
	double get_num2();
	
private:
	double num1;
	double num2;
};