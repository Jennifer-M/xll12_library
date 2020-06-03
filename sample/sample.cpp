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
	Function(XLL_LPOPER, L"?xll_function", L"XLL.FUNCTION")
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

            deposits = LENGTH(depositData);
            fras = LENGTH(fraData);
            immFuts = LENGTH(immFutData);
            asxFuts = LENGTH(asxFutData);
            swaps = LENGTH(swapData);
            bonds = LENGTH(bondData);
            bmas = LENGTH(bmaData);

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
LPOPER WINAPI xll_function(XLOPER12* depositData1,
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

    sizeDepositData = LENGTH(depositData);
    std::vector<Rate> expectedRate(sizeDepositData);
    std::vector<Rate> estimatedRate(sizeDepositData);
    std::vector<Rate> subExpectedRate;
    for (Size i = 0; i < sizeDepositData; i++) {
        Euribor index(depositData[i].n * depositData[i].units, curveHandle);
        expectedRate[i] = depositData[i].rate / 100;
            estimatedRate[i] = index.fixing(vars.today);
        if (std::fabs(expectedRate[i] - estimatedRate[i]) > tolerance) {
            std::cout <<
                depositData[i].n << " "
                << (depositData[i].units == Weeks ? "week(s)" : "month(s)")
                << " deposit:"
                //<< std::setprecision(8)
                << "\n    estimated rate: " << io::rate(estimatedRate[i])
                << "\n    expected rate:  " << io::rate(expectedRate[i]);
        }
    }
    Rate sub[50];
    for (int j = 0; j < sizeDepositData; j++) {
        sub[j] = expectedRate[j];
    }
    
    result = sub;
    return &result;
}

/*
template <class T, class I, template<class C> class B>
LPOPER WINAPI xll_function(XLOPER12* depositData1,
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
    double expectedRate[sizeDepositData];
    vector<Rate> subExpectedRate;
    for (Size i = 0; i < sizeDepositData; i++) {
        Euribor index(depositData[i].n * depositData[i].units, curveHandle);
        Rate expectedRate[i] = depositData[i].rate / 100,
            estimatedRate = index.fixing(calendar.adjust(Date::todaysDate()););
        if (std::fabs(expectedRate[i] - estimatedRate) > tolerance) {
            std::cout <<
                depositData[i].n << " "
                << (depositData[i].units == Weeks ? "week(s)" : "month(s)")
                << " deposit:"
                //<< std::setprecision(8)
                << "\n    estimated rate: " << io::rate(estimatedRate)
                << "\n    expected rate:  " << io::rate(expectedRate[i]);
        }
    }
    result = expectedRate;

    //    testBMACurveConsistency<ZeroYield, Linear, IterativeBootstrap>(vars);
    return &result;
} */



/*
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
*/