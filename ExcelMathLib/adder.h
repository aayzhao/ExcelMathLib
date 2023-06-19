#pragma once
#ifdef ADDER_EXCELMATHLIB_EXPORTS
#define ADDER_API __declspec(dllexport)
#else
#define ADDER_API __declspec(dllimport)
#endif

extern "C"
{
	struct adder;
	typedef struct adder adder_t;

	ADDER_API void* adder_create();
	ADDER_API adder_t* adder_create_2(double start);
	ADDER_API adder_t* adder_create_3(double start1, double start2);
	ADDER_API void adder_destroy(adder_t* a);
	ADDER_API double adder_add(adder_t* a);
	ADDER_API double adder_add_2(adder_t* a, double val);
	ADDER_API double adder_sum(adder_t* a, double x, double y);
	ADDER_API double adder_get_one(adder_t*);
	ADDER_API double adder_get_two(adder_t*);

}
