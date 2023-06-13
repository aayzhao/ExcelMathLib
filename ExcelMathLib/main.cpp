#include "pch.h"
#include "AddNum.h"
#include <iostream>

using namespace std;

int main()
{
	AddNum ob(0);

	cout << ob.addNums() << "\n";

	ob.num1 = 10.0;
	ob.num2 = 20.1;

	cout << ob.addNums() << "\n";

	return 0;
}