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

    //a being in the set{flatForward,piecewiseyield,InterpolatedDiscountCurve}
    void linkTermStructure(ext::shared_ptr<YieldTermStructure>& a) {
        termStructure.linkTo(a);
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

        ext::shared_ptr<Quote> forward = ext::shared_ptr<Quote>(new SimpleQuote(0.05));
        DayCounter dc = Actual365Fixed();
        ext::shared_ptr<YieldTermStructure> a = 
        ext::shared_ptr<YieldTermStructure>(
            new FlatForward(settlement, Handle<Quote>(forward), dc));

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

//having one curve to fit all price or multiple curve with each fitting one pirce
AddIn swap_function_multi(
    // Function returning a pointer to an OPER with C++ name xll_function and Excel name XLL.FUNCTION.
    // Don't forget prepend a question mark to the C++ name.
    //                     |
    //                     v
    Function(XLL_LPOPER, L"?swapTest_multi", L"XLL.swapTest.multi")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"variables", L"is the first integer argument.", L"2")
    .Arg(XLL_LPOPER, L"curve_variables", L"is the first integer argument.", L"2")
    .Uncalced()//needs to be there for declaring handle
    // Paste function category.
    .Category(L"Example1")
    // Insert Function description.
    .FunctionHelp(L"Help on XLL.FUNCTION goes here.")
    // Create entry for this function in Sandcastle Help File Builder project file.
    .Alias(L"XLL.FUNCTION.ALIAS.swap") // alternate name
);

struct CurveVal{
//flatfoward input
    Date referenceDate;
    Rate forward;
    DayCounter dayCounter;

//piecewiseyieldcurve input
//    Traits trait;
//    Interpolator interpolat;
//    Booststrap = booststrap;

    //Date referenceDate;
//    ext::std::vector<ext::shared_ptr<typename Traits::helper> > helper;
    //DayCounter& dayCounter;
//    Interpolator iInput;

//InterpolatedDiscountCurve
    std::vector<Date> dates;
    std::vector<DiscountFactor> dfs;
    //DayCounter& dayCounter;
    Calendar calendar;
    //Interpolator& interpolator;

//FittedBondDiscountCurve
    Natural settlementDays;
    //Calendar& calendar;
    std::vector<ext::shared_ptr<BondHelper> > bondHelpers;
    //DayCounter& dayCounter;
//    FittingMethod& fittingMethod;
    Real accuracy;
    Size maxEvaluations;
    Array guess;
    Real simplexLambda;
    Size maxStationaryStateIterations;

    CurveVal() {
        settlementDays = 2;
    };
};


//FOr converting wstring to string for string input usage
std::string wtoa1(const std::wstring& Text) {

    std::string s(WideCharToMultiByte(CP_UTF8, 0, Text.c_str(), Text.size(), NULL, NULL, NULL, NULL), '\0');
    s.resize(WideCharToMultiByte(CP_UTF8, 0, Text.c_str(), Text.size(), &s[0], s.size(), NULL, NULL));
    return s;
};

int convertTsType(std::string str) {
    int strsize = str.length();
    if (str.find("flatForward") < strsize) {
        return 1;
    }
    return -1;
}


struct IstV {
    Integer length;
    Rate fixedRate;
    Spread floatingSpread;
};


//Input:HANDLEX xhdlcv (length,fixedRate,floatingSpread from input)
//       HANDLEX xhdlyts (termstructures)
//Output:vector<double>* pointer to vector of double with price in it
//Expectation: we expect to give multiple input of the termstrcuture and makeswap and calculate the price of the swap
std::vector<double> termToPriceMulit(CommonVars1 & vars,HANDLEX xhdlist, HANDLEX xhdlyts) {
    //CommonVars1 vars;
    std::vector<double> r;
    std::vector<IstV>* ist = dynamic_cast<std::vector<IstV>*>(handle<std::vector<IstV>>(xhdlyts).ptr());
    std::vector<ext::shared_ptr<YieldTermStructure>>* yts = dynamic_cast<std::vector<ext::shared_ptr<YieldTermStructure>>*>(
                                                    handle<std::vector<ext::shared_ptr<YieldTermStructure>>>(xhdlyts).ptr());

    int Ssize = ist->size();
    for (int i = 0; i < Ssize; i++) {
        vars.linkTermStructure(yts->at(i));
        Integer length = ist->at(i).length;
        Rate fixedRate = ist->at(i).fixedRate;
        Spread floatingSpread = ist->at(i).floatingSpread;
        ext::shared_ptr<VanillaSwap> swap = vars.makeSwap(length, fixedRate, floatingSpread);
        r.push_back(swap->NPV());
    }
    return r;
}
//Input:HANDLEX xhdlcv (length,fixedRate,floatingSpread from input)
//       HANDLEX xhdlyts (termstructures)
//Output:vector<double>* pointer to vector of double with price in it
//Expectation: we expect to give a single input of the termstrcuture and multiple input of makeswap and calculate the price of the swap
std::vector<double> termToPriceSingle(HANDLEX xhdlist, HANDLEX xhdlyts) {

    CommonVars1 vars;
    std::vector<double> r;

    std::vector<IstV>* ist = dynamic_cast<std::vector<IstV>*>(handle<std::vector<IstV>>(xhdlyts).ptr());
    std::vector<ext::shared_ptr<YieldTermStructure>>* yts = dynamic_cast<std::vector<ext::shared_ptr<YieldTermStructure>>*>(
                                                    handle<std::vector<ext::shared_ptr<YieldTermStructure>>>(xhdlyts).ptr());

    vars.linkTermStructure(yts->at(0));
    int Ssize = ist->size();
    for (int i = 0; i < Ssize; i++) {
        Integer length = ist->at(i).length;
        Rate fixedRate = ist->at(i).fixedRate;
        Spread floatingSpread = ist->at(i).floatingSpread;
        ext::shared_ptr<VanillaSwap> swap = vars.makeSwap(length, fixedRate, floatingSpread);
        r.push_back(swap->NPV());
    }
    return r;
}



class curve {
private:
    ext::shared_ptr<YieldTermStructure> yts;

public:
    curve() {}
    curve(ext::shared_ptr<YieldTermStructure> y):yts(y){}
    ext::shared_ptr<YieldTermStructure> getYts() {
        return yts;
    }
    void setYts(ext::shared_ptr<YieldTermStructure>& y) {
        yts = y;
    }
};

class flatforward : public curve {
private:
    Date settlement;
    Rate rate;
    DayCounter dc;
public: 
    flatforward(Date settlement, Rate rate, DayCounter dc):
                settlement(settlement),rate(rate),dc(dc){

        ext::shared_ptr<YieldTermStructure> sub =
            ext::shared_ptr<YieldTermStructure>(new FlatForward(settlement,
                Handle<Quote>(ext::shared_ptr<Quote>(new SimpleQuote(rate))),
                dc));
        setYts(sub);
    };
    Rate getRate() {
        return rate;
    }

};

void test1() {
   Calendar calendar = TARGET();
   Date today = calendar.adjust(Date::todaysDate());

    flatforward* a = new flatforward(today, 0.05, Actual365Fixed());
    std::vector<curve*> sub;
    sub.push_back(a);

    std::pair<int, curve*> pa = std::make_pair(2, new flatforward(today, 0.05, Actual365Fixed()));
    handle <flatforward> simplehandle(a);
    handlex sub1;
    sub1 = simplehandle.get();

    flatforward* pd = dynamic_cast<flatforward*>(handle<flatforward>(sub1).ptr());

    Rate pddata = pd->getRate();

    handle<std::vector<curve*>> subhandle(&sub);
    handlex xhandle;
    xhandle = subhandle.get();
    ext::shared_ptr<YieldTermStructure> testYts = subhandle->at(0)->getYts();
}
LPOPER WINAPI swapTest_multi(XLOPER12* variables,XLOPER12* curve_variables) {

#pragma XLLEXPORT
    static OPER result;
//testing handle
 //   test1();

    int curve_size = (int)curve_variables->val.array.rows;
    int data_size = (int)variables->val.array.rows;

    CommonVars1 vars;
    std::vector<double> res;

    int size =curve_size;
    std::vector<ext::shared_ptr<YieldTermStructure>> ytsvec;
    int count = 0;
    for (int i = 0; i < size; i++) {
        std::string str = wtoa1(curve_variables->val.array.lparray[count].val.str);
        count++;
        int tsType = convertTsType(str);
        ext::shared_ptr<YieldTermStructure> a;
        if (tsType == 1) {
            Integer settlementDays = (int)curve_variables->val.array.lparray[count].val.num;
            count++;
            Rate forward = curve_variables->val.array.lparray[count].val.num;
            count++;
            ext::shared_ptr<IborIndex> index = ext::shared_ptr<IborIndex>(new
                Euribor(Period(Semiannual), vars.termStructure));
            Calendar calendar = index->fixingCalendar();
            Date today = calendar.adjust(Settings::instance().evaluationDate());
            Date settlement = calendar.advance(today, settlementDays, Days);
            a = ext::shared_ptr<YieldTermStructure>(new FlatForward(settlement,
                Handle<Quote>(ext::shared_ptr<Quote>(new SimpleQuote(forward))),
                Actual365Fixed()));

        }
        ytsvec.push_back(a);
    }
    handle<std::vector<ext::shared_ptr<YieldTermStructure>>> hdlyts(&ytsvec);
    handlex xhdlyts;
    xhdlyts = hdlyts.get();

    size = data_size;
    count = 0;
    std::vector<IstV> ist;
    for (int i = 0; i < size; i++) {
        Integer length = (int)variables->val.array.lparray[count].val.num;
        count++;
        Rate fixedRate = variables->val.array.lparray[count].val.num;
        count++;
        Spread floatingSpread = variables->val.array.lparray[count].val.num;
        count++;
        IstV a = { length,fixedRate,floatingSpread };
        ist.push_back(a);
    }
    handle<std::vector<IstV>> hdlist(&ist);
    handlex xhdlist;
    xhdlist = hdlist.get();

    if (curve_size == 1) {
        res = termToPriceSingle(xhdlist, xhdlyts);
    }
    else {
        res = termToPriceMulit(vars, xhdlist, xhdlyts);
    }

    OPER sub(data_size, 1);
    for (int j = 0; j < data_size; j++) {
        sub[j] = res[j];
    }
    result = sub;
    return &result;

}