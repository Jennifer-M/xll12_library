//#define BOOST_TEST_MODULE sample
//#define BOOST_TEST_DYN_LINK


#include <ql/quantlib.hpp>
#include <ql/time/timeunit.hpp>
#include <ql/time/frequency.hpp>
#include <test-suite/piecewiseyieldcurve.hpp>
#include <test-suite/utilities.hpp>
#include <ql/indexes/ibor/euribor.hpp>

//#include <boost/test/unit_test.hpp>
//#include <ql/time/date.hpp> 
//#include <ql/time/calendar.hpp>
//#include <boost/move/detail/meta_utils.hpp>

// sample.cpp - Simple example of using AddIn.
#include <cmath>
#include "../xll/xll.h"
#include "../xll/shfb/entities.h"

#include <iostream>
#include <algorithm>
using namespace xll;
using namespace QuantLib;

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
	Function(XLL_LPOPER, L"?depositDataTest", L"XLL.depositDataTest")
	// First argument is a double called x with an argument description and default value of 2
	.Arg(XLL_LPOPER, L"depositData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"fraData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"immFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"asxFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"swapData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bondData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bmaData1", L"is the first double argument.", L"2")
    /*
    .Arg(XLL_DOUBLE, L"n", L"is the first n argument.", L"2")
    .Arg(XLL_LPOPER, L"units", L"is the second unit argument.", L"2")
    .Arg(XLL_DOUBLE, L"rate", L"is the third rate argument.", L"2")
    */
    /*
    .Arg(XLL_DOUBLE, L"x", L"is the first double argument.", L"2")
    */
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
/*
LPOPER WINAPI xll_function(double x)
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
}
*/

namespace piecewise_yield_curve_test {

    struct Datum {
        Integer n;
        TimeUnit units;
        Rate rate;
    };

    struct BondDatum {
        Integer n;
        TimeUnit units;
        Integer length;
        Frequency frequency;
        Rate coupon;
        Real price;
    };
    struct CommonVars {
        // global variables
        Calendar calendar;
        Natural settlementDays;
        Date today, settlement;
        BusinessDayConvention fixedLegConvention;
        Frequency fixedLegFrequency;
        DayCounter fixedLegDayCounter;
        Natural bondSettlementDays;
        DayCounter bondDayCounter;
        BusinessDayConvention bondConvention;
        Real bondRedemption;
        Frequency bmaFrequency;
        BusinessDayConvention bmaConvention;
        DayCounter bmaDayCounter;

        Size deposits, fras, immFuts, asxFuts, swaps, bonds, bmas;
        std::vector<ext::shared_ptr<SimpleQuote> > rates, fraRates,
            immFutPrices, asxFutPrices,
            prices, fractions;
        std::vector<ext::shared_ptr<RateHelper> > instruments, fraHelpers,
            immFutHelpers, asxFutHelpers,
            bondHelpers, bmaHelpers;
        std::vector<Schedule> schedules;
        ext::shared_ptr<YieldTermStructure> termStructure;

        // cleanup
        SavedSettings backup;
        //QuantLib::IndexHistoryCleaner cleaner;

        // setup
        CommonVars(std::vector<Datum>& depositData, std::vector<Datum>& fraData, std::vector<Datum>& immFutData,
            std::vector<Datum>& asxFutData, std::vector<Datum>& swapData, std::vector<BondDatum>& bondData,
            std::vector<Datum>& bmaData) {
            // data
            calendar = TARGET();
            settlementDays = 2;
            today = calendar.adjust(Date::todaysDate());
            Settings::instance().evaluationDate() = today;
            settlement = calendar.advance(today, settlementDays, Days);
            fixedLegConvention = Unadjusted;
            fixedLegFrequency = Annual;
            fixedLegDayCounter = Thirty360();
            bondSettlementDays = 3;
            bondDayCounter = ActualActual();
            bondConvention = Following;
            bondRedemption = 100.0;
            bmaFrequency = Quarterly;
            bmaConvention = Following;
            bmaDayCounter = ActualActual();

            deposits = (int)depositData.size();
            fras = (int)fraData.size();
            immFuts = (int)immFutData.size();
            asxFuts = (int)asxFutData.size();
            swaps = (int)swapData.size();
            bonds = (int)bondData.size();
            bmas = (int)bmaData.size();

            // market elements
            rates =
                std::vector<ext::shared_ptr<SimpleQuote> >(deposits + swaps);
            fraRates = std::vector<ext::shared_ptr<SimpleQuote> >(fras);
            immFutPrices = std::vector<ext::shared_ptr<SimpleQuote> >(immFuts);
            asxFutPrices = std::vector<ext::shared_ptr<SimpleQuote> >(asxFuts);
            prices = std::vector<ext::shared_ptr<SimpleQuote> >(bonds);
            fractions = std::vector<ext::shared_ptr<SimpleQuote> >(bmas);
            for (Size i = 0; i < deposits; i++) {
                rates[i] = ext::make_shared<SimpleQuote>(
                    depositData[i].rate / 100);
            }
            for (Size i = 0; i < swaps; i++) {
                rates[i + deposits] = ext::make_shared<SimpleQuote>(
                    swapData[i].rate / 100);
            }
            for (Size i = 0; i < fras; i++) {
                fraRates[i] = ext::make_shared<SimpleQuote>(
                    fraData[i].rate / 100);
            }
            for (Size i = 0; i < bonds; i++) {
                prices[i] = ext::make_shared<SimpleQuote>(
                    bondData[i].price);
            }
            for (Size i = 0; i < immFuts; i++) {
                immFutPrices[i] = ext::make_shared<SimpleQuote>(
                    100.0 - immFutData[i].rate);
            }
            for (Size i = 0; i < asxFuts; i++) {
                asxFutPrices[i] = ext::make_shared<SimpleQuote>(
                    100.0 - asxFutData[i].rate);
            }
            for (Size i = 0; i < bmas; i++) {
                fractions[i] = ext::make_shared<SimpleQuote>(
                    bmaData[i].rate / 100);
            }

            // rate helpers
            instruments =
                std::vector<ext::shared_ptr<RateHelper> >(deposits + swaps);
            fraHelpers = std::vector<ext::shared_ptr<RateHelper> >(fras);
            immFutHelpers = std::vector<ext::shared_ptr<RateHelper> >(immFuts);
            asxFutHelpers = std::vector<ext::shared_ptr<RateHelper> >();
            bondHelpers = std::vector<ext::shared_ptr<RateHelper> >(bonds);
            schedules = std::vector<Schedule>(bonds);
            bmaHelpers = std::vector<ext::shared_ptr<RateHelper> >(bmas);

            ext::shared_ptr<IborIndex> euribor6m(new Euribor6M);
            for (Size i = 0; i < deposits; i++) {
                Handle<Quote> r(rates[i]);
                instruments[i] = ext::shared_ptr<RateHelper>(new
                    DepositRateHelper(r,
                        ext::make_shared<Euribor>(
                            depositData[i].n * depositData[i].units)));
            }
            for (Size i = 0; i < swaps; i++) {
                Handle<Quote> r(rates[i + deposits]);
                instruments[i + deposits] = ext::shared_ptr<RateHelper>(new
                    SwapRateHelper(r, swapData[i].n * swapData[i].units,
                        calendar,
                        fixedLegFrequency, fixedLegConvention,
                        fixedLegDayCounter, euribor6m));
            }


#ifdef QL_USE_INDEXED_COUPON
            bool useIndexedFra = false;
#else
            bool useIndexedFra = true;
#endif

            ext::shared_ptr<IborIndex> euribor3m(new Euribor3M());
            for (Size i = 0; i < fras; i++) {
                Handle<Quote> r(fraRates[i]);
                fraHelpers[i] = ext::shared_ptr<RateHelper>(new
                    FraRateHelper(r, fraData[i].n, fraData[i].n + 3,
                        euribor3m->fixingDays(),
                        euribor3m->fixingCalendar(),
                        euribor3m->businessDayConvention(),
                        euribor3m->endOfMonth(),
                        euribor3m->dayCounter(),
                        Pillar::LastRelevantDate,
                        Date(),
                        useIndexedFra));
            }
            Date immDate = Date();
            for (Size i = 0; i < immFuts; i++) {
                Handle<Quote> r(immFutPrices[i]);
                immDate = IMM::nextDate(immDate, false);
                // if the fixing is before the evaluation date, we
                // just jump forward by one future maturity
                if (euribor3m->fixingDate(immDate) <
                    Settings::instance().evaluationDate())
                    immDate = IMM::nextDate(immDate, false);
                immFutHelpers[i] = ext::shared_ptr<RateHelper>(new
                    FuturesRateHelper(r, immDate, euribor3m, Handle<Quote>(),
                        Futures::IMM));
            }
            Date asxDate = Date();
            for (Size i = 0; i < asxFuts; i++) {
                Handle<Quote> r(asxFutPrices[i]);
                asxDate = ASX::nextDate(asxDate, false);
                // if the fixing is before the evaluation date, we
                // just jump forward by one future maturity
                if (euribor3m->fixingDate(asxDate) <
                    Settings::instance().evaluationDate())
                    asxDate = ASX::nextDate(asxDate, false);
                if (euribor3m->fixingCalendar().isBusinessDay(asxDate))
                    asxFutHelpers.push_back(ext::shared_ptr<RateHelper>(new
                        FuturesRateHelper(r, asxDate, euribor3m,
                            Handle<Quote>(), Futures::ASX)));
            }

            for (Size i = 0; i < bonds; i++) {
                Handle<Quote> p(prices[i]);
                Date maturity =
                    calendar.advance(today, bondData[i].n, bondData[i].units);
                Date issue =
                    calendar.advance(maturity, -bondData[i].length, Years);
                std::vector<Rate> coupons(1, bondData[i].coupon / 100.0);
                schedules[i] = Schedule(issue, maturity,
                    Period(bondData[i].frequency),
                    calendar,
                    bondConvention, bondConvention,
                    DateGeneration::Backward, false);
                bondHelpers[i] = ext::shared_ptr<RateHelper>(new
                    FixedRateBondHelper(p,
                        bondSettlementDays,
                        bondRedemption, schedules[i],
                        coupons, bondDayCounter,
                        bondConvention,
                        bondRedemption, issue));
            }
        }
    };
}

std::string wtoa(const std::wstring& Text) {

    std::string s(WideCharToMultiByte(CP_UTF8, 0, Text.c_str(), Text.size(), NULL, NULL, NULL, NULL), '\0');
    s.resize(WideCharToMultiByte(CP_UTF8, 0, Text.c_str(), Text.size(), &s[0], s.size(), NULL, NULL));
    return s;
};

TimeUnit convertTime(std::string x) {
    int xsize = x.length();
    if (x.find("Months")<xsize) {
        return Months;
    }
    else if (x.find("Years") < xsize) {
        return Years;
    }
    else if (x.find("Days") < xsize) {
        return Days;
    }
    else if (x.find("Weeks") < xsize) {
        return Weeks;
    }
    else if (x.find("Hours") < xsize) {
        return Hours;
    }
    else if (x.find("Minutes") < xsize) {
        return Minutes;
    }
    else if (x.find("Seconds") < xsize) {
        return Seconds;
    }
    else if (x.find("Milliseconds") < xsize) {
        return Milliseconds;
    }
    else if (x.find("Microseconds") < xsize) {
        return Microseconds;
    }
    else {
        std::cout << "error with TimeUnit input" << std::endl;
        return Microseconds;
    }

};
Frequency convertFreq(std::string x) {
    int xsize = x.length();
    if (x.find("NoFrequency") < xsize) {
        return NoFrequency;
    }
    else if (x.find("Once") < xsize) {
        return Once;
    }
    else if (x.find("Annual") < xsize) {
        return Annual;
    }
    else if (x.find("Semiannual") < xsize) {
        return Semiannual;
    }
    else if (x.find("EveryFourthMonth") < xsize) {
        return EveryFourthMonth;
    }
    else if (x.find("Quarterly") < xsize) {
        return Quarterly;
    }
    else if (x.find("Bimonthly") < xsize) {
        return Bimonthly;
    }
    else if (x.find("Monthly") < xsize) {
        return Monthly;
    }
    else if (x.find("EveryFourthWeek") < xsize) {
        return EveryFourthWeek;
    }
    else if (x.find("Biweekly") < xsize) {
        return Biweekly;
    }
    else if (x.find("Weekly") < xsize) {
        return Weekly;
    }
    else if (x.find("Daily") < xsize) {
        return Daily;
    }
    else {
        return OtherFrequency;
    }

};
LPOPER WINAPI depositDataTest(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;

    using namespace piecewise_yield_curve_test;

    Real tolerance = 1.0e-9;
    int sizeDepositData = (int)depositData1->val.array.rows;
    int count = 0;
    std::vector<Datum> depositData;
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = (int)depositData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(depositData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = depositData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    count = 0;
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData;
    for (int i = 0; i < sizefraData; i++) {
        Integer a = (int)fraData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(fraData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = fraData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    count = 0;
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData;
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = (int)immFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(immFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = immFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    count = 0;
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData;
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = (int)asxFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(asxFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = asxFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    count = 0;
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData;
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = (int)swapData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(swapData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = swapData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    count = 0;
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData;
    for (int i = 0; i < sizebondData; i++) {
        Integer a = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bondData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Integer c = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string dd = wtoa(bondData1->val.array.lparray[count].val.str);
        Frequency d = convertFreq(dd);
        count++;
        Rate e = bondData1->val.array.lparray[count].val.num;
        count++;
        Real f = bondData1->val.array.lparray[count].val.num;
        count++;
        BondDatum sub = { a, b, c, d, e, f };
        bondData.push_back(sub);
    }

    count = 0;
    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = (int)bmaData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bmaData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = bmaData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);

    QuantLib::Linear interpolator = Linear();
    RelinkableHandle<YieldTermStructure> curveHandle;

    curveHandle.linkTo(ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<ZeroYield, Linear, IterativeBootstrap>(vars.settlement, vars.instruments,
            Actual360(),
            interpolator)));

    sizeDepositData = (int)depositData.size();
    std::vector<Rate> expectedRate(sizeDepositData);
    std::vector<Rate> estimatedRate(sizeDepositData);
    for (Size i = 0; i < sizeDepositData; i++) {
        Euribor index(depositData[i].n * depositData[i].units, curveHandle);
        expectedRate[i] = depositData[i].rate / 100;
            estimatedRate[i] = index.fixing(vars.today);
        if (std::fabs(expectedRate[i] - estimatedRate[i]) > tolerance) {
            BOOST_ERROR(
                depositData[i].n << " "
                << (depositData[i].units == Weeks ? "week(s)" : "month(s)")
                << " deposit:"
                //<< std::setprecision(8)
                << "\n    estimated rate: " << io::rate(estimatedRate[i])
                << "\n    expected rate:  " << io::rate(expectedRate[i]));
        }
    }
    count = 0;
    OPER sub(sizeDepositData,3);
    for (int j = 0; j < sizeDepositData; j++) {
        sub[count] = expectedRate[j];
        count++;
        sub[count] = estimatedRate[j];
        count++;
        sub[count] = std::fabs(expectedRate[j] - estimatedRate[j]) < tolerance ? true : false;
        count++;
    }
    
    result = sub;
    return &result;
}

AddIn xai_function1(
    // Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
    // Don't forget prepend a question mark to the C++ name.
    Function(XLL_LPOPER, L"?swapDataTest", L"XLL.swapDataTest")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"depositData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"fraData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"immFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"asxFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"swapData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bondData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bmaData1", L"is the first double argument.", L"2")
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
);
LPOPER WINAPI swapDataTest(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;

    using namespace piecewise_yield_curve_test;
    Real tolerance = 1.0e-9;
    int sizeDepositData = (int)depositData1->val.array.rows;
    int count = 0;
    std::vector<Datum> depositData;
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = (int)depositData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(depositData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = depositData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    count = 0;
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData;
    for (int i = 0; i < sizefraData; i++) {
        Integer a = (int)fraData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(fraData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = fraData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    count = 0;
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData;
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = (int)immFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(immFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = immFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    count = 0;
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData;
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = (int)asxFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(asxFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = asxFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    count = 0;
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData;
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = (int)swapData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(swapData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = swapData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    count = 0;
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData;
    for (int i = 0; i < sizebondData; i++) {
        Integer a = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bondData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Integer c = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string dd = wtoa(bondData1->val.array.lparray[count].val.str);
        Frequency d = convertFreq(dd);
        count++;
        Rate e = bondData1->val.array.lparray[count].val.num;
        count++;
        Real f = bondData1->val.array.lparray[count].val.num;
        count++;
        BondDatum sub = { a, b, c, d, e, f };
        bondData.push_back(sub);
    }

    count = 0;
    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = (int)bmaData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bmaData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = bmaData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);

    QuantLib::Linear interpolator = Linear();
    RelinkableHandle<YieldTermStructure> curveHandle;

    curveHandle.linkTo(ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<ZeroYield, Linear, IterativeBootstrap>(vars.settlement, vars.instruments,
            Actual360(),
            interpolator)));

    std::vector<Rate> expectedRate(vars.swaps);
    std::vector<Rate> estimatedRate(vars.swaps);

    ext::shared_ptr<IborIndex> euribor6m(new Euribor6M(curveHandle));
    for (Size i = 0; i < vars.swaps; i++) {
        Period tenor = swapData[i].n * swapData[i].units;

        VanillaSwap swap = MakeVanillaSwap(tenor, euribor6m, 0.0)
            .withEffectiveDate(vars.settlement)
            .withFixedLegDayCount(vars.fixedLegDayCounter)
            .withFixedLegTenor(Period(vars.fixedLegFrequency))
            .withFixedLegConvention(vars.fixedLegConvention)
            .withFixedLegTerminationDateConvention(vars.fixedLegConvention);

        expectedRate[i] = swapData[i].rate / 100;
            estimatedRate[i] = swap.fairRate();
        Spread error = std::fabs(expectedRate[i] - estimatedRate[i]);
        if (error > tolerance) {
            BOOST_ERROR(
                swapData[i].n << " year(s) swap:\n"
                << std::setprecision(8)
                << "\n estimated rate: " << io::rate(estimatedRate[i])
                << "\n expected rate:  " << io::rate(expectedRate[i])
                << "\n error:          " << io::rate(error)
                << "\n tolerance:      " << io::rate(tolerance));
        }
    }

    count = 0;
    OPER sub(vars.swaps, 3);
    for (int j = 0; j < vars.swaps; j++) {
        sub[count] = expectedRate[j];
        count++;
        sub[count] = estimatedRate[j];
        count++;
        sub[count] = std::fabs(expectedRate[j] - estimatedRate[j]) < tolerance ? true : false;
        count++;
    }

    result = sub;
    return &result;
};

AddIn xai_function2(
    // Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
    // Don't forget prepend a question mark to the C++ name.
    Function(XLL_LPOPER, L"?bondDataTest", L"XLL.bondDataTest")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"depositData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"fraData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"immFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"asxFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"swapData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bondData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bmaData1", L"is the first double argument.", L"2")
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
);

LPOPER WINAPI bondDataTest(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;

    using namespace piecewise_yield_curve_test;
    Real tolerance = 1.0e-9;
    int sizeDepositData = (int)depositData1->val.array.rows;
    int count = 0;
    std::vector<Datum> depositData;
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = (int)depositData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(depositData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = depositData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    count = 0;
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData;
    for (int i = 0; i < sizefraData; i++) {
        Integer a = (int)fraData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(fraData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = fraData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    count = 0;
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData;
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = (int)immFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(immFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = immFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    count = 0;
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData;
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = (int)asxFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(asxFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = asxFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    count = 0;
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData;
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = (int)swapData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(swapData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = swapData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    count = 0;
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData;
    for (int i = 0; i < sizebondData; i++) {
        Integer a = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bondData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Integer c = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string dd = wtoa(bondData1->val.array.lparray[count].val.str);
        Frequency d = convertFreq(dd);
        count++;
        Rate e = bondData1->val.array.lparray[count].val.num;
        count++;
        Real f = bondData1->val.array.lparray[count].val.num;
        count++;
        BondDatum sub = { a, b, c, d, e, f };
        bondData.push_back(sub);
    }

    count = 0;
    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = (int)bmaData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bmaData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = bmaData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);
    QuantLib::Linear interpolator = Linear();
    RelinkableHandle<YieldTermStructure> curveHandle;

    // check bonds
    vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<ZeroYield, Linear, IterativeBootstrap>(vars.settlement, vars.bondHelpers,
            Actual360(),
            interpolator));
    curveHandle.linkTo(vars.termStructure);

    std::vector<Real> expectedPrice(vars.bonds);
    std::vector<Real> estimatedPrice(vars.bonds);

    for (Size i = 0; i < vars.bonds; i++) {
        Date maturity = vars.calendar.advance(vars.today,
            bondData[i].n,
            bondData[i].units);
        Date issue = vars.calendar.advance(maturity,
            -bondData[i].length,
            Years);
        std::vector<Rate> coupons(1, bondData[i].coupon / 100.0);

        FixedRateBond bond(vars.bondSettlementDays, 100.0,
            vars.schedules[i], coupons,
            vars.bondDayCounter, vars.bondConvention,
            vars.bondRedemption, issue);

        ext::shared_ptr<PricingEngine> bondEngine(
            new DiscountingBondEngine(curveHandle));
        bond.setPricingEngine(bondEngine);

        expectedPrice[i] = bondData[i].price;
            estimatedPrice[i] = bond.cleanPrice();
        Real error = std::fabs(expectedPrice[i] - estimatedPrice[i]);
        if (error > tolerance) {
            BOOST_ERROR(io::ordinal(i + 1) << " bond failure:" <<
                std::setprecision(8) <<
                "\n  estimated price: " << estimatedPrice[i] <<
                "\n  expected price:  " << expectedPrice[i] <<
                "\n  error:           " << error);
        }
    }

    count = 0;
    OPER sub(vars.bonds, 3);
    for (int j = 0; j < vars.bonds; j++) {
        sub[count] = expectedPrice[j];
        count++;
        sub[count] = estimatedPrice[j];
        count++;
        sub[count] = std::fabs(expectedPrice[j] - estimatedPrice[j]) < tolerance ? true : false;
        count++;
    }

    result = sub;
    return &result;
};

AddIn xai_function3(
    // Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
    // Don't forget prepend a question mark to the C++ name.
    Function(XLL_LPOPER, L"?fraDataTest", L"XLL.fraDataTest")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"depositData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"fraData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"immFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"asxFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"swapData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bondData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bmaData1", L"is the first double argument.", L"2")
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
);

LPOPER WINAPI fraDataTest(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;

    using namespace piecewise_yield_curve_test;
    Real tolerance = 1.0e-9;
    int sizeDepositData = (int)depositData1->val.array.rows;
    int count = 0;
    std::vector<Datum> depositData;
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = (int)depositData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(depositData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = depositData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    count = 0;
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData;
    for (int i = 0; i < sizefraData; i++) {
        Integer a = (int)fraData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(fraData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = fraData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    count = 0;
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData;
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = (int)immFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(immFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = immFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    count = 0;
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData;
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = (int)asxFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(asxFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = asxFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    count = 0;
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData;
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = (int)swapData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(swapData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = swapData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    count = 0;
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData;
    for (int i = 0; i < sizebondData; i++) {
        Integer a = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bondData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Integer c = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string dd = wtoa(bondData1->val.array.lparray[count].val.str);
        Frequency d = convertFreq(dd);
        count++;
        Rate e = bondData1->val.array.lparray[count].val.num;
        count++;
        Real f = bondData1->val.array.lparray[count].val.num;
        count++;
        BondDatum sub = { a, b, c, d, e, f };
        bondData.push_back(sub);
    }

    count = 0;
    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = (int)bmaData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bmaData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = bmaData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);
    QuantLib::Linear interpolator = Linear();
    RelinkableHandle<YieldTermStructure> curveHandle;

    vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<ZeroYield, Linear>(vars.settlement, vars.fraHelpers,
            Actual360(),
            interpolator));
    curveHandle.linkTo(vars.termStructure);

#ifdef QL_USE_INDEXED_COUPON
    bool useIndexedFra = false;
#else
    bool useIndexedFra = true;
#endif
    std::vector<Rate> expectedRate(vars.fras);
    std::vector<Rate> estimatedRate(vars.fras);
    ext::shared_ptr<IborIndex> euribor3m(new Euribor3M(curveHandle));
    for (Size i = 0; i < vars.fras; i++) {
        Date start =
            vars.calendar.advance(vars.settlement,
                fraData[i].n,
                fraData[i].units,
                euribor3m->businessDayConvention(),
                euribor3m->endOfMonth());
        BOOST_REQUIRE(fraData[i].units == Months);
        Date end = vars.calendar.advance(vars.settlement, 3 + fraData[i].n, Months,
            euribor3m->businessDayConvention(),
            euribor3m->endOfMonth());

        ForwardRateAgreement fra(start, end, Position::Long,
            fraData[i].rate / 100, 100.0,
            euribor3m, curveHandle,
            useIndexedFra);
        expectedRate[i] = fraData[i].rate / 100;
        estimatedRate[i] = fra.forwardRate();
        if (std::fabs(expectedRate[i] - estimatedRate[i]) > tolerance) {
            BOOST_ERROR(io::ordinal(i + 1) << " FRA failure:" <<
                std::setprecision(8) <<
                "\n  estimated rate: " << io::rate(estimatedRate[i]) <<
                "\n  expected rate:  " << io::rate(expectedRate[i]));
        }
    }

    count = 0;
    OPER sub(vars.fras, 3);
    for (int j = 0; j < vars.fras; j++) {
        sub[count] = expectedRate[j];
        count++;
        sub[count] = estimatedRate[j];
        count++;
        sub[count] = std::fabs(expectedRate[j] - estimatedRate[j]) < tolerance ? true : false;
        count++;
    }

    result = sub;
    return &result;
}

AddIn xai_function4(
    // Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
    // Don't forget prepend a question mark to the C++ name.
    Function(XLL_LPOPER, L"?immFutsDataTest", L"XLL.immFutsDataTest")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"depositData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"fraData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"immFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"asxFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"swapData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bondData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bmaData1", L"is the first double argument.", L"2")
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
);

LPOPER WINAPI immFutsDataTest(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;

    using namespace piecewise_yield_curve_test;
    Real tolerance = 1.0e-9;
    int sizeDepositData = (int)depositData1->val.array.rows;
    int count = 0;
    std::vector<Datum> depositData;
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = (int)depositData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(depositData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = depositData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    count = 0;
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData;
    for (int i = 0; i < sizefraData; i++) {
        Integer a = (int)fraData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(fraData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = fraData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    count = 0;
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData;
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = (int)immFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(immFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = immFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    count = 0;
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData;
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = (int)asxFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(asxFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = asxFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    count = 0;
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData;
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = (int)swapData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(swapData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = swapData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    count = 0;
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData;
    for (int i = 0; i < sizebondData; i++) {
        Integer a = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bondData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Integer c = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string dd = wtoa(bondData1->val.array.lparray[count].val.str);
        Frequency d = convertFreq(dd);
        count++;
        Rate e = bondData1->val.array.lparray[count].val.num;
        count++;
        Real f = bondData1->val.array.lparray[count].val.num;
        count++;
        BondDatum sub = { a, b, c, d, e, f };
        bondData.push_back(sub);
    }

    count = 0;
    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = (int)bmaData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bmaData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = bmaData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);
    
    QuantLib::Linear interpolator = Linear();
    RelinkableHandle<YieldTermStructure> curveHandle;
    vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<ZeroYield, Linear>(vars.settlement, vars.immFutHelpers,
            Actual360(),
            interpolator));
    curveHandle.linkTo(vars.termStructure);

    ext::shared_ptr<IborIndex> euribor3m(new Euribor3M(curveHandle));

    std::vector<Rate> expectedRate(vars.immFuts);
    std::vector<Rate> estimatedRate(vars.immFuts);
    Date immStart = Date();
    for (Size i = 0; i < vars.immFuts; i++) {
        immStart = IMM::nextDate(immStart, false);
        // if the fixing is before the evaluation date, we
        // just jump forward by one future maturity
        if (euribor3m->fixingDate(immStart) <
            Settings::instance().evaluationDate())
            immStart = IMM::nextDate(immStart, false);
        Date end = vars.calendar.advance(immStart, 3, Months,
            euribor3m->businessDayConvention(),
            euribor3m->endOfMonth());

        ForwardRateAgreement immFut(immStart, end, Position::Long,
            immFutData[i].rate / 100, 100.0,
            euribor3m, curveHandle);
        expectedRate[i] = immFutData[i].rate / 100;
            estimatedRate[i] = immFut.forwardRate();
        if (std::fabs(expectedRate[i] - estimatedRate[i]) > tolerance) {
            BOOST_ERROR(io::ordinal(i + 1) << " IMM futures failure:" <<
                std::setprecision(8) <<
                "\n  estimated rate: " << io::rate(estimatedRate[i]) <<
                "\n  expected rate:  " << io::rate(expectedRate[i]));
        }
    }

    count = 0;
    OPER sub(vars.immFuts, 3);
    for (int j = 0; j < vars.immFuts; j++) {
        sub[count] = expectedRate[j];
        count++;
        sub[count] = estimatedRate[j];
        count++;
        sub[count] = std::fabs(expectedRate[j] - estimatedRate[j]) < tolerance ? true : false;
        count++;
    }

    result = sub;
    return &result;
}

AddIn xai_function5(
    // Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
    // Don't forget prepend a question mark to the C++ name.
    Function(XLL_LPOPER, L"?asxFutsDataTest", L"XLL.asxFutsDataTest")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"depositData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"fraData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"immFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"asxFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"swapData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bondData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bmaData1", L"is the first double argument.", L"2")
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
);

LPOPER WINAPI asxFutsDataTest(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;

    using namespace piecewise_yield_curve_test;
    Real tolerance = 1.0e-9;
    int sizeDepositData = (int)depositData1->val.array.rows;
    int count = 0;
    std::vector<Datum> depositData;
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = (int)depositData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(depositData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = depositData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    count = 0;
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData;
    for (int i = 0; i < sizefraData; i++) {
        Integer a = (int)fraData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(fraData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = fraData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    count = 0;
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData;
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = (int)immFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(immFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = immFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    count = 0;
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData;
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = (int)asxFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(asxFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = asxFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    count = 0;
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData;
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = (int)swapData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(swapData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = swapData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    count = 0;
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData;
    for (int i = 0; i < sizebondData; i++) {
        Integer a = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bondData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Integer c = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string dd = wtoa(bondData1->val.array.lparray[count].val.str);
        Frequency d = convertFreq(dd);
        count++;
        Rate e = bondData1->val.array.lparray[count].val.num;
        count++;
        Real f = bondData1->val.array.lparray[count].val.num;
        count++;
        BondDatum sub = { a, b, c, d, e, f };
        bondData.push_back(sub);
    }

    count = 0;
    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = (int)bmaData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bmaData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = bmaData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);
    QuantLib::Linear interpolator = Linear();
    RelinkableHandle<YieldTermStructure> curveHandle;


    vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<ZeroYield, Linear>(vars.settlement, vars.asxFutHelpers,
            Actual360(),
            interpolator));
    curveHandle.linkTo(vars.termStructure);

    ext::shared_ptr<IborIndex> euribor3m(new Euribor3M(curveHandle));

    std::vector<Rate> expectedRate(vars.immFuts);
    std::vector<Rate> estimatedRate(vars.immFuts);

    Date asxStart = Date();
    for (Size i = 0; i < vars.asxFuts; i++) {
        asxStart = ASX::nextDate(asxStart, false);
        // if the fixing is before the evaluation date, we
        // just jump forward by one future maturity
        if (euribor3m->fixingDate(asxStart) <
            Settings::instance().evaluationDate())
            asxStart = ASX::nextDate(asxStart, false);
        if (euribor3m->fixingCalendar().isHoliday(asxStart))
            continue;
        Date end = vars.calendar.advance(asxStart, 3, Months,
            euribor3m->businessDayConvention(),
            euribor3m->endOfMonth());

        ForwardRateAgreement asxFut(asxStart, end, Position::Long,
            asxFutData[i].rate / 100, 100.0,
            euribor3m, curveHandle);
        expectedRate[i] = asxFutData[i].rate / 100;
            estimatedRate[i] = asxFut.forwardRate();
        if (std::fabs(expectedRate[i] - estimatedRate[i]) > tolerance) {
            BOOST_ERROR(io::ordinal(i + 1) << " ASX futures failure:" <<
                std::setprecision(8) <<
                "\n  estimated rate: " << io::rate(estimatedRate[i]) <<
                "\n  expected rate:  " << io::rate(expectedRate[i]));
        }
    }
    count = 0;
    OPER sub(vars.asxFuts, 3);
    for (int j = 0; j < vars.asxFuts; j++) {
        sub[count] = expectedRate[j];
        count++;
        sub[count] = estimatedRate[j];
        count++;
        sub[count] = std::fabs(expectedRate[j] - estimatedRate[j]) < tolerance ? true : false;
        count++;
    }

    result = sub;
    return &result;
}

AddIn xai_functionbma(
    // Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
    // Don't forget prepend a question mark to the C++ name.
    Function(XLL_LPOPER, L"?bmaTest", L"XLL.bmaTest")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"depositData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"fraData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"immFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"asxFutData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"swapData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bondData1", L"is the first double argument.", L"2")
    .Arg(XLL_LPOPER, L"bmaData1", L"is the first double argument.", L"2")
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
);

LPOPER WINAPI bmaTest(XLOPER12* depositData1,
    XLOPER12* fraData1,
    XLOPER12* immFutData1,
    XLOPER12* asxFutData1,
    XLOPER12* swapData1,
    XLOPER12* bondData1,
    XLOPER12* bmaData1) {
#pragma XLLEXPORT
    static OPER result;

    using namespace piecewise_yield_curve_test;
    Real tolerance = 1.0e-9;

    int sizeDepositData = (int)depositData1->val.array.rows;
    int count = 0;
    std::vector<Datum> depositData;
    for (int i = 0; i < sizeDepositData; i++) {
        Integer a = (int)depositData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(depositData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = depositData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        depositData.push_back(sub);
    }
    count = 0;
    int sizefraData = (int)fraData1->val.array.rows;
    std::vector<Datum> fraData;
    for (int i = 0; i < sizefraData; i++) {
        Integer a = (int)fraData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(fraData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = fraData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        fraData.push_back(sub);
    }
    count = 0;
    int sizeimmFutData = (int)immFutData1->val.array.rows;
    std::vector<Datum> immFutData;
    for (int i = 0; i < sizeimmFutData; i++) {
        Integer a = (int)immFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(immFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = immFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        immFutData.push_back(sub);
    }
    count = 0;
    int sizeasxFutData = (int)asxFutData1->val.array.rows;
    std::vector<Datum> asxFutData;
    for (int i = 0; i < sizeasxFutData; i++) {
        Integer a = (int)asxFutData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(asxFutData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = asxFutData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        asxFutData.push_back(sub);
    }
    count = 0;
    int sizeswapData = (int)swapData1->val.array.rows;
    std::vector<Datum> swapData;
    for (int i = 0; i < sizeswapData; i++) {
        Integer a = (int)swapData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(swapData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = swapData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        swapData.push_back(sub);
    }
    count = 0;
    int sizebondData = (int)bondData1->val.array.rows;
    std::vector<BondDatum> bondData;
    for (int i = 0; i < sizebondData; i++) {
        Integer a = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bondData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Integer c = (int)bondData1->val.array.lparray[count].val.num;
        count++;
        std::string dd = wtoa(bondData1->val.array.lparray[count].val.str);
        Frequency d = convertFreq(dd);
        count++;
        Rate e = bondData1->val.array.lparray[count].val.num;
        count++;
        Real f = bondData1->val.array.lparray[count].val.num;
        count++;
        BondDatum sub = { a, b, c, d, e, f };
        bondData.push_back(sub);
    }

    count = 0;
    int sizebmaData = (int)bmaData1->val.array.rows;
    std::vector<Datum> bmaData(sizebmaData);
    for (int i = 0; i < sizebmaData; i++) {
        Integer a = (int)bmaData1->val.array.lparray[count].val.num;
        count++;
        std::string bb = wtoa(bmaData1->val.array.lparray[count].val.str);
        TimeUnit b = convertTime(bb);
        count++;
        Rate c = bmaData1->val.array.lparray[count].val.num;
        count++;
        Datum sub = { a, b, c };
        bmaData.push_back(sub);
    }
    CommonVars vars(depositData, fraData, immFutData, asxFutData, swapData, bondData, bmaData);

    // re-adjust settlement
    vars.calendar = JointCalendar(BMAIndex().fixingCalendar(),
        USDLibor(3 * Months).fixingCalendar(),
        JoinHolidays);
    vars.today = vars.calendar.adjust(Date::todaysDate());
    Settings::instance().evaluationDate() = vars.today;
    vars.settlement =
        vars.calendar.advance(vars.today, vars.settlementDays, Days);


    Handle<YieldTermStructure> riskFreeCurve(
        ext::shared_ptr<YieldTermStructure>(
            new FlatForward(vars.settlement, 0.04, Actual360())));

    ext::shared_ptr<BMAIndex> bmaIndex(new BMAIndex);
    ext::shared_ptr<IborIndex> liborIndex(
        new USDLibor(3 * Months, riskFreeCurve));
    for (Size i = 0; i < vars.bmas; ++i) {
        Handle<Quote> f(vars.fractions[i]);
        vars.bmaHelpers[i] = ext::shared_ptr<RateHelper>(
            new BMASwapRateHelper(f, bmaData[i].n * bmaData[i].units,
                vars.settlementDays,
                vars.calendar,
                Period(vars.bmaFrequency),
                vars.bmaConvention,
                vars.bmaDayCounter,
                bmaIndex,
                liborIndex));
    }

    Weekday w = vars.today.weekday();
    Date lastWednesday =
        (w >= 4) ? vars.today - (w - 4) : vars.today + (4 - w - 7);
    Date lastFixing = bmaIndex->fixingCalendar().adjust(lastWednesday);
    bmaIndex->addFixing(lastFixing, 0.03);

    QuantLib::Linear interpolator = Linear();

    vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
        PiecewiseYieldCurve<ZeroYield, Linear, IterativeBootstrap>(vars.today, vars.bmaHelpers,
            Actual360(),
            interpolator));

    RelinkableHandle<YieldTermStructure> curveHandle;
    curveHandle.linkTo(vars.termStructure);

    // check BMA swaps
    std::vector<Real> expectedFraction(vars.bmas);
    std::vector<Real> estimatedFraction(vars.bmas);

    ext::shared_ptr<BMAIndex> bma(new BMAIndex(curveHandle));
    ext::shared_ptr<IborIndex> libor3m(new USDLibor(3 * Months,
        riskFreeCurve));
    for (Size i = 0; i < vars.bmas; i++) {
        Period tenor = bmaData[i].n * bmaData[i].units;

        Schedule bmaSchedule =
            MakeSchedule().from(vars.settlement)
            .to(vars.settlement + tenor)
            .withFrequency(vars.bmaFrequency)
            .withCalendar(bma->fixingCalendar())
            .withConvention(vars.bmaConvention)
            .backwards();
        Schedule liborSchedule =
            MakeSchedule().from(vars.settlement)
            .to(vars.settlement + tenor)
            .withTenor(libor3m->tenor())
            .withCalendar(libor3m->fixingCalendar())
            .withConvention(libor3m->businessDayConvention())
            .endOfMonth(libor3m->endOfMonth())
            .backwards();


        BMASwap swap(BMASwap::Payer, 100.0,
            liborSchedule, 0.75, 0.0,
            libor3m, libor3m->dayCounter(),
            bmaSchedule, bma, vars.bmaDayCounter);
        swap.setPricingEngine(ext::shared_ptr<PricingEngine>(
            new DiscountingSwapEngine(libor3m->forwardingTermStructure())));

        expectedFraction[i] = bmaData[i].rate / 100;
            estimatedFraction[i] = swap.fairLiborFraction();
        Real error = std::fabs(expectedFraction[i] - estimatedFraction[i]);
        if (error > tolerance) {
            BOOST_ERROR(bmaData[i].n << " year(s) BMA swap:\n"
                << std::setprecision(8)
                << "\n estimated libor fraction: " << estimatedFraction[i]
                << "\n expected libor fraction:  " << expectedFraction[i]
                << "\n error:          " << error
                << "\n tolerance:      " << tolerance);
        }
    }

    count = 0;
    OPER sub(vars.bmas, 3);
    for (int j = 0; j < vars.bmas; j++) {
        sub[count] = expectedFraction[j];
        count++;
        sub[count] = estimatedFraction[j];
        count++;
        sub[count] = std::fabs(expectedFraction[j] - estimatedFraction[j]) < tolerance ? true : false;
        count++;
    }

    result = sub;
    return &result;
}