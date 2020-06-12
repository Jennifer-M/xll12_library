#include <test-suite/utilities.hpp>
#include <test-suite/swap.hpp>
#include <ql/termstructures/yield/flatforward.hpp>
#include <ql/handle.hpp>
#include <ql/quantlib.hpp>

#include "../xll/xll.h"
#include "../xll/shfb/entities.h"

#include <iostream>
#include <algorithm>

using namespace xll;
using namespace QuantLib;

AddIn xai_sample1(
    Document(L"sample1") // top level documentation
    .Documentation(LR"(
This object will generate a Sandcastle Helpfile Builder project file.
)"));

AddIn xal_sample_category1(
    Document(L"Example1")
    .Documentation(LR"(
This object will generate documentation for the Example category.
)")
);

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

struct CommonVars1 {
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

    CommonVars1() {
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

        Date today = settlement;
        ext::shared_ptr<Quote> forward = ext::shared_ptr<Quote>(new SimpleQuote(0.05));
        DayCounter dc = Actual365Fixed();
        ext::shared_ptr<YieldTermStructure> a = 
        ext::shared_ptr<YieldTermStructure>(
            new FlatForward(today, Handle<Quote>(forward), dc));

        termStructure.linkTo(a);
    }
};

LPOPER WINAPI swapTest(XLOPER12* variables) {

#pragma XLLEXPORT
    static OPER result;

    CommonVars1 vars;
    Integer length = (int)variables->val.array.lparray[0].val.num;
    Rate fixedRate = variables->val.array.lparray[1].val.num;
    Rate floatingSpread = variables->val.array.lparray[2].val.num;
    ext::shared_ptr<VanillaSwap> swap = vars.makeSwap(length, fixedRate, floatingSpread);
    result = swap->NPV();
    return &result;

}
