#include <test-suite/swap.hpp>
#include <test-suite/utilities.hpp>
#include <ql/handle.hpp>
#include <ql/quantlib.hpp>
#include <ql/instruments/vanillaswap.hpp>
#include <ql/pricingengines/swap/discountingswapengine.hpp>
#include <ql/termstructures/yield/flatforward.hpp>
#include <ql/time/calendars/nullcalendar.hpp>
#include <ql/time/daycounters/thirty360.hpp>
#include <ql/time/daycounters/actual365fixed.hpp>
#include <ql/time/daycounters/simpledaycounter.hpp>
#include <ql/time/schedule.hpp>
#include <ql/indexes/ibor/euribor.hpp>
#include <ql/cashflows/iborcoupon.hpp>
#include <ql/cashflows/cashflowvectors.hpp>
#include <ql/termstructures/volatility/optionlet/constantoptionletvol.hpp>
#include <ql/utilities/dataformatters.hpp>
#include <ql/cashflows/cashflows.hpp>
#include <ql/cashflows/couponpricer.hpp>
#include <ql/currencies/europe.hpp>


#include "../xll/xll.h"
#include "../xll/shfb/entities.h"

#include <iostream>
#include <algorithm>

using namespace xll;
using namespace QuantLib;

// Information Excel needs to register add-in.
AddIn swap_function(
	// Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
	// Don't forget prepend a question mark to the C++ name.
	//                     |
	//                     v
	Function(XLL_LPOPER, L"?swapTest", L"XLL.swapTest")
	// First argument is a double called x with an argument description and default value of 2
	.Arg(XLL_LPOPER, L"variables", L"is the first integer argument.", L"2")
	// Paste function category.
	.Category(L"Example1")
	// Insert Function description.
	.FunctionHelp(L"Help on XLL.FUNCTION goes here.")
	// Create entry for this function in Sandcastle Help File Builder project file.
	.Alias(L"XLL.FUNCTION.ALIAS.swap") // alternate name
);

struct CommonVars {
    // global data
    Date today, settlement;
    VanillaSwap::Type type;
    Real nominal;
    Calendar calendar;
    BusinessDayConvention fixedConvention, floatingConvention;
    Frequency fixedFrequency, floatingFrequency;
    DayCounter fixedDayCount;
    ext::shared_ptr<IborIndex> index;
    Natural settlementDays;
    RelinkableHandle<YieldTermStructure> termStructure;

    // cleanup
    SavedSettings backup;

    // utilities
    ext::shared_ptr<VanillaSwap> makeSwap(Integer length, Rate fixedRate,
        Spread floatingSpread) {
        Date maturity = calendar.advance(settlement, length, Years,
            floatingConvention);
        Schedule fixedSchedule(settlement, maturity, Period(fixedFrequency),
            calendar, fixedConvention, fixedConvention,
            DateGeneration::Forward, false);
        Schedule floatSchedule(settlement, maturity,
            Period(floatingFrequency),
            calendar, floatingConvention,
            floatingConvention,
            DateGeneration::Forward, false);
        ext::shared_ptr<VanillaSwap> swap(
            new VanillaSwap(type, nominal,
                fixedSchedule, fixedRate, fixedDayCount,
                floatSchedule, index, floatingSpread,
                index->dayCounter()));
        swap->setPricingEngine(ext::shared_ptr<PricingEngine>(
            new DiscountingSwapEngine(termStructure)));
        return swap;
    }

    CommonVars() {
        type = VanillaSwap::Payer;
        settlementDays = 2;
        nominal = 100.0;
        fixedConvention = Unadjusted;
        floatingConvention = ModifiedFollowing;
        fixedFrequency = Annual;
        floatingFrequency = Semiannual;
        fixedDayCount = Thirty360();
        index = ext::shared_ptr<IborIndex>(new
            Euribor(Period(floatingFrequency), termStructure));
        calendar = index->fixingCalendar();
        today = calendar.adjust(Settings::instance().evaluationDate());
        settlement = calendar.advance(today, settlementDays, Days);
        termStructure.linkTo(flatRate(settlement, 0.05, Actual365Fixed()));
    }
};

LPOPER WINAPI swapTest(XLOPER12* variables) {

#pragma XLLEXPORT
    static OPER result;

    CommonVars vars;
    Integer length = (int)variables->val.array.lparray[0].val.num;
    Rate fixedRate = variables->val.array.lparray[1].val.num;
    Rate floatingSpread = variables->val.array.lparray[2].val.num;
    ext::shared_ptr<VanillaSwap> swap = vars.makeSwap(length, fixedRate, floatingSpread);
    result = swap->NPV();
    return &result;

}

