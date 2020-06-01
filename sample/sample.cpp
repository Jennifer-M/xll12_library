// sample.cpp - Simple example of using AddIn.
#include <cmath>
#include "../xll/xll.h"
#include "../xll/shfb/entities.h"


#include <ql/quantlib.hpp>
#include <iostream>

using namespace xll;

AddIn xai_sample(
	Document(L"sample") // top level documentation
	.Documentation(LR"(
This object will generate a Sandcastle Helpfile Builder project file.
)"));

AddIn xal_sample_category(
	Document(L"Example")
	.Documentation(LR"(
This object will generate documentation for the Example category.
)")
);

// Information Excel needs to register add-in.
AddIn xai_function(
	// Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
	// Don't forget prepend a question mark to the C++ name.
	//                     |
	//                     v
	Function(XLL_LPOPER, L"?xll_function", L"XLL.FUNCTION")
	// First argument is a double called x with an argument description and default value of 2
	.Arg(XLL_DOUBLE, L"x", L"is the first double argument.", L"2")
	// Paste function category.
	.Category(L"Example")
	// Insert Function description.
	.FunctionHelp(L"Help on XLL.FUNCTION goes here.")
	// Create entry for this function in Sandcastle Help File Builder project file.
	.Alias(L"XLL.FUNCTION.ALIAS") // alternate name
	.Documentation(
		PARA_(
			L"Free-form documentation for " C_(L"XLL.FUNCTION") L" goes here."
		)
		PARA_(L"But you can include MAML directives.")
		PARA_(L"This is " B_(L"bold") " and so is " B_("this"))
		PARA_(L"Math: " MATH_(int_ SUB_(minus_ infin_) SUP_(infin_) L"e" SUP_(L"x" sup2_ L"/2") L" dx"))
	)
	.Remarks(
		L"This is a remark. "
		L"This is " B_(L"bold") L" and this is " I_(L"italic")
		PARA_(L"This is a paragraph.")
	)
	.Examples(LR"(
This is an example.
</para><para>
It has two paragraphs.
)")
);
// Calling convention *must* be WINAPI (aka __stdcall) for Excel.
/*LPOPER WINAPI xll_function(double x)
{
	// Be sure to export your function.
#pragma XLLEXPORT
	static OPER result;

	try {
		ensure(x >= 0);
		result = sqrt(x); // OPER's act like Excel cells.
	}
	catch (const std::exception & ex) {
		XLL_ERROR(ex.what());

		result = OPER(xlerr::Num);
	}

	return &result;
}*/
/**
std::string wtoa(const std::wstring& Text) {
    std::string s(WideCharToMultiByte(CP_UTF8, 0, Text.c_str(), Text.size(), NULL, NULL, NULL, NULL), '\0');
    s.resize(WideCharToMultiByte(CP_UTF8, 0, Text.c_str(), Text.size(), &s[0], s.size(), NULL, NULL));
    return s;
}

TimeUnit convertTime(std::string x) {
    if (x == "Months") {
        return Months;
    }
    else if (x == "Years") {
        return Years;
    }
    else if (x == "Days") {
        return Days;
    }
    else if (x == "Weeks") {
        return Weeks;
    }
    else if (x == "Hours") {
        return Hours;
    }
    else if (x == "Minutes") {
        return Minutes;
    }
    else if (x == "Seconds") {
        return Seconds;
    }
    else if (x == "Milliseconds") {
        return Milliseconds;
    }
    else {
        return Microseconds;
    }

}
template <class T, class I, template<class C> class B>
LPOPER WINAPI testLinearZeroConsistency(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;
    //BOOST_TEST_MESSAGE(
    //    "Testing consistency of piecewise-linear zero-yield curve...");

    using namespace piecewise_yield_curve_test;
    // int sizedepositData = depositData1
     //Datum  depositData;
     //for (int i)

    Real tolerance = 1.0e-9;
    int sizeDepositData = (int)depositData1->val.array.rows;
    std::vector<Datum> depositData(sizeDepositData);
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = depositData1->val.array.lparray[i].val.array.lparray[0].val.w;
        std::string bb = wtoa(depositData1->val.array.lparray[i].val.array.lparray[1].val.str);
        TimeUnit b = convertTime(bb);
        Rate c = depositData1->val.array.lparray[i].val.array.lparray[2].val.w;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData(sizefraData);
    for (int i = 0; i < sizefraData; i++) {
        Integer a = fraData1->val.array.lparray[i].val.array.lparray[0].val.w;
        std::string bb = wtoa(fraData1->val.array.lparray[i].val.array.lparray[1].val.str);
        TimeUnit b = convertTime(bb);
        Rate c = fraData1->val.array.lparray[i].val.array.lparray[2].val.w;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData(sizeimmFutData);
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = immFutData1->val.array.lparray[i].val.array.lparray[0].val.w;
        std::string bb = wtoa(immFutData1->val.array.lparray[i].val.array.lparray[1].val.str);
        TimeUnit b = convertTime(bb);
        Rate c = immFutData1->val.array.lparray[i].val.array.lparray[2].val.w;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData(sizeasxFutData);
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = asxFutData1->val.array.lparray[i].val.array.lparray[0].val.w;
        std::string bb = wtoa(asxFutData1->val.array.lparray[i].val.array.lparray[1].val.str);
        TimeUnit b = convertTime(bb);
        Rate c = asxFutData1->val.array.lparray[i].val.array.lparray[2].val.w;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData(sizeswapData);
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = swapData1->val.array.lparray[i].val.array.lparray[0].val.w;
        std::string bb = wtoa(swapData1->val.array.lparray[i].val.array.lparray[1].val.str);
        TimeUnit b = convertTime(bb);
        Rate c = swapData1->val.array.lparray[i].val.array.lparray[2].val.w;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData(sizebondData);
    for (int i = 0; i < sizebondData; i++) {
        Integer a = bondData1->val.array.lparray[i].val.array.lparray[0].val.w;
        std::string bb = wtoa(bondData1->val.array.lparray[i].val.array.lparray[1].val.str);
        TimeUnit b = convertTime(bb);
        Rate c = bondData1->val.array.lparray[i].val.array.lparray[2].val.w;
        BondDatum sub = { a, b, c };
        bondData.push_back(sub);
    }

    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = bmaData1->val.array.lparray[i].val.array.lparray[0].val.w;
        std::string bb = wtoa(bmaData1->val.array.lparray[i].val.array.lparray[1].val.str);
        TimeUnit b = convertTime(bb);
        Rate c = bmaData1->val.array.lparray[i].val.array.lparray[2].val.w;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);

    //   testCurveConsistency<ZeroYield, Linear, IterativeBootstrap>(vars);
    RelinkableHandle<YieldTermStructure> curveHandle;

    curveHandle.linkTo(ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<T, I, B>(vars.settlement, vars.instruments,
            Actual360(),
            interpolator)));

    sizeDepositData = LENGTH(depositData);
    int a[sizeDepositData];
    vector<Rate> subExpectedRate;
    for (Size i = 0; i < sizeDepositData; i++) {
        Euribor index(depositData[i].n * depositData[i].units, curveHandle);
        Rate expectedRate = depositData[i].rate / 100,
            estimatedRate = index.fixing(calendar.adjust(Date::todaysDate()););
        if (std::fabs(expectedRate - estimatedRate) > tolerance) {
            std::cout <<
                depositData[i].n << " "
                << (depositData[i].units == Weeks ? "week(s)" : "month(s)")
                << " deposit:"
                //<< std::setprecision(8)
                << "\n    estimated rate: " << io::rate(estimatedRate)
                << "\n    expected rate:  " << io::rate(expectedRate);
        }
    }
    

    //    testBMACurveConsistency<ZeroYield, Linear, IterativeBootstrap>(vars);
    return &result;
} **/

LPOPER WINAPI xll_function(double x)
{
#pragma XLLEXPORT
    static OPER result;

    QuantLib::Calendar myCal = QuantLib::UnitedKingdom();
    QuantLib::Date newYearsEve(31, QuantLib::Dec, 2008);

    std::cout << "Name: " << myCal.name() << std::endl;
    std::cout << "New Year is Holiday: " << myCal.isHoliday(newYearsEve) << std::endl;
    std::cout << "New Year is Business Day: " << myCal.isBusinessDay(newYearsEve) << std::endl;

    std::cout << "--------------- Date Counter --------------------" << std::endl;

    QuantLib::Date date1(28, QuantLib::Dec, 2008);
    QuantLib::Date date2(04, QuantLib::Jan, 2009);

    std::cout << "First Date: " << date1 << std::endl;
    std::cout << "Second Date: " << date2 << std::endl;
    std::cout << "Business Days Betweeen: " << myCal.businessDaysBetween(date1, date2) << std::endl;
    std::cout << "End of Month 1. Date: " << myCal.endOfMonth(date1) << std::endl;
    std::cout << "End of Month 2. Date: " << myCal.endOfMonth(date2) << std::endl;

    double tmp;
    std::cin >> tmp;

    return &result;
}