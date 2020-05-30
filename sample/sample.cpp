// sample.cpp - Simple example of using AddIn.
#include <cmath>

#include <iostream>
#include "../xll/xll.h"
#include "../xll/shfb/entities.h"

//include quantlib libraries
#include <test-suite/piecewiseyieldcurve.hpp>
#include <test-suite/utilities.hpp>
//#include <ql/quantlib.hpp>

#include <ql/cashflows/iborcoupon.hpp>
#include <ql/termstructures/globalbootstrap.hpp>
#include <ql/termstructures/yield/piecewiseyieldcurve.hpp>
#include <ql/termstructures/yield/ratehelpers.hpp>
#include <ql/termstructures/yield/bondhelpers.hpp>
#include <ql/termstructures/yield/flatforward.hpp>
#include <ql/time/calendars/target.hpp>
#include <ql/time/calendars/japan.hpp>
#include <ql/time/calendars/weekendsonly.hpp>
#include <ql/time/calendars/jointcalendar.hpp>
#include <ql/time/daycounters/actual360.hpp>
#include <ql/time/daycounters/actualactual.hpp>
#include <ql/time/daycounters/thirty360.hpp>
#include <ql/time/imm.hpp>
#include <ql/time/asx.hpp>
#include <ql/indexes/ibor/euribor.hpp>
#include <ql/indexes/ibor/usdlibor.hpp>
#include <ql/indexes/ibor/jpylibor.hpp>
#include <ql/indexes/bmaindex.hpp>
#include <ql/indexes/indexmanager.hpp>
#include <ql/instruments/forwardrateagreement.hpp>
#include <ql/instruments/makevanillaswap.hpp>
#include <ql/math/interpolations/linearinterpolation.hpp>
#include <ql/math/interpolations/loginterpolation.hpp>
#include <ql/math/interpolations/backwardflatinterpolation.hpp>
#include <ql/math/interpolations/cubicinterpolation.hpp>
#include <ql/math/interpolations/convexmonotoneinterpolation.hpp>
#include <ql/math/comparison.hpp>
#include <ql/quotes/simplequote.hpp>
#include <ql/utilities/dataformatters.hpp>
#include <ql/pricingengines/bond/discountingbondengine.hpp>
#include <ql/pricingengines/swap/discountingswapengine.hpp>
#include <iomanip>
#include <boost/test/unit_test.hpp>
//#include <boost/shared_ptr.hpp>
//#include <boost/assign/list_of.hpp>




//quantlib using
using namespace QuantLib;
using namespace boost::unit_test_framework;

using namespace xll;

AddIn xai_sample(
	Document(L"sample") // top level documentation
	.Category(L"sample")
	.Documentation(LR"(
This object will generate a Sandcastle Helpfile Builder project file.
)"));

AddIn xal_sample_category(
	Document(L"Example")
	.Category(L"Example")
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
    Function(XLL_LPOPER, L"?testLinearZeroConsistency", L"XLL.FUNCTIONs")
    // First argument is a double called x with an argument description and default value of 2
    .Arg(XLL_LPOPER, L"depositData1", L"deposit data", L"")
    .Arg(XLL_LPOPER,L"fraData1",L"fra data",L"")
    .Arg(XLL_LPOPER, L"immFutData1", L"immFut data", L"")
    .Arg(XLL_LPOPER, L"asxFutData1", L"asxFut data", L"")
    .Arg(XLL_LPOPER, L"swapData1", L"swap data", L"")
    .Arg(XLL_LPOPER, L"bondData1", L"bond data", L"")
    .Arg(XLL_LPOPER, L"bmaData1", L"bma data", L"")
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

#define QL_REAL double
typedef QL_REAL Real;

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
        QuantLib::IndexHistoryCleaner cleaner;

        // setup
        CommonVars(std::vector<Datum> &depositData, std::vector<Datum> &fraData, std::vector<Datum> &immFutData, 
            std::vector<Datum>& asxFutData, std::vector<Datum>& swapData, std::vector<BondDatum>& bondData, 
            std::vector<Datum>&bmaData) {
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


    template <class T, class I, template<class C> class B>
    void testCurveConsistency(CommonVars& vars,
        const I& interpolator = I(),
        Real tolerance = 1.0e-9) {

        vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
            PiecewiseYieldCurve<T, I, B>(vars.settlement, vars.instruments,
                Actual360(),
                interpolator));

        RelinkableHandle<YieldTermStructure> curveHandle;
        curveHandle.linkTo(vars.termStructure);

        // check deposits
        for (Size i = 0; i < vars.deposits; i++) {
            Euribor index(depositData[i].n * depositData[i].units, curveHandle);
            Rate expectedRate = depositData[i].rate / 100,
                estimatedRate = index.fixing(vars.today);
            if (std::fabs(expectedRate - estimatedRate) > tolerance) {
                BOOST_ERROR(
                    depositData[i].n << " "
                    << (depositData[i].units == Weeks ? "week(s)" : "month(s)")
                    << " deposit:"
                    << std::setprecision(8)
                    << "\n    estimated rate: " << io::rate(estimatedRate)
                    << "\n    expected rate:  " << io::rate(expectedRate));
            }
        }
        return expectedRate;

        // check swaps
        ext::shared_ptr<IborIndex> euribor6m(new Euribor6M(curveHandle));
        for (Size i = 0; i < vars.swaps; i++) {
            Period tenor = swapData[i].n * swapData[i].units;

            VanillaSwap swap = MakeVanillaSwap(tenor, euribor6m, 0.0)
                .withEffectiveDate(vars.settlement)
                .withFixedLegDayCount(vars.fixedLegDayCounter)
                .withFixedLegTenor(Period(vars.fixedLegFrequency))
                .withFixedLegConvention(vars.fixedLegConvention)
                .withFixedLegTerminationDateConvention(vars.fixedLegConvention);

            Rate expectedRate = swapData[i].rate / 100,
                estimatedRate = swap.fairRate();
            Spread error = std::fabs(expectedRate - estimatedRate);
            if (error > tolerance) {
                BOOST_ERROR(
                    swapData[i].n << " year(s) swap:\n"
                    << std::setprecision(8)
                    << "\n estimated rate: " << io::rate(estimatedRate)
                    << "\n expected rate:  " << io::rate(expectedRate)
                    << "\n error:          " << io::rate(error)
                    << "\n tolerance:      " << io::rate(tolerance));
            }
        }

        // check bonds
        vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
            PiecewiseYieldCurve<T, I, B>(vars.settlement, vars.bondHelpers,
                Actual360(),
                interpolator));
        curveHandle.linkTo(vars.termStructure);

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

            Real expectedPrice = bondData[i].price,
                estimatedPrice = bond.cleanPrice();
            Real error = std::fabs(expectedPrice - estimatedPrice);
            if (error > tolerance) {
                BOOST_ERROR(io::ordinal(i + 1) << " bond failure:" <<
                    std::setprecision(8) <<
                    "\n  estimated price: " << estimatedPrice <<
                    "\n  expected price:  " << expectedPrice <<
                    "\n  error:           " << error);
            }
        }

        // check FRA
        vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
            PiecewiseYieldCurve<T, I>(vars.settlement, vars.fraHelpers,
                Actual360(),
                interpolator));
        curveHandle.linkTo(vars.termStructure);

#ifdef QL_USE_INDEXED_COUPON
        bool useIndexedFra = false;
#else
        bool useIndexedFra = true;
#endif

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
            Rate expectedRate = fraData[i].rate / 100,
                estimatedRate = fra.forwardRate();
            if (std::fabs(expectedRate - estimatedRate) > tolerance) {
                BOOST_ERROR(io::ordinal(i + 1) << " FRA failure:" <<
                    std::setprecision(8) <<
                    "\n  estimated rate: " << io::rate(estimatedRate) <<
                    "\n  expected rate:  " << io::rate(expectedRate));
            }
        }

        // check immFuts
        vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
            PiecewiseYieldCurve<T, I>(vars.settlement, vars.immFutHelpers,
                Actual360(),
                interpolator));
        curveHandle.linkTo(vars.termStructure);

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
            Rate expectedRate = immFutData[i].rate / 100,
                estimatedRate = immFut.forwardRate();
            if (std::fabs(expectedRate - estimatedRate) > tolerance) {
                BOOST_ERROR(io::ordinal(i + 1) << " IMM futures failure:" <<
                    std::setprecision(8) <<
                    "\n  estimated rate: " << io::rate(estimatedRate) <<
                    "\n  expected rate:  " << io::rate(expectedRate));
            }
        }

        // check asxFuts
        vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
            PiecewiseYieldCurve<T, I>(vars.settlement, vars.asxFutHelpers,
                Actual360(),
                interpolator));
        curveHandle.linkTo(vars.termStructure);

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
            Rate expectedRate = asxFutData[i].rate / 100,
                estimatedRate = asxFut.forwardRate();
            if (std::fabs(expectedRate - estimatedRate) > tolerance) {
                BOOST_ERROR(io::ordinal(i + 1) << " ASX futures failure:" <<
                    std::setprecision(8) <<
                    "\n  estimated rate: " << io::rate(estimatedRate) <<
                    "\n  expected rate:  " << io::rate(expectedRate));
            }
        }

        // end checks
    }

    template <class T, class I, template<class C> class B>
    void testBMACurveConsistency(CommonVars& vars,
        const I& interpolator = I(),
        Real tolerance = 1.0e-9) {

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

        vars.termStructure = ext::shared_ptr<YieldTermStructure>(new
            PiecewiseYieldCurve<T, I, B>(vars.today, vars.bmaHelpers,
                Actual360(),
                interpolator));

        RelinkableHandle<YieldTermStructure> curveHandle;
        curveHandle.linkTo(vars.termStructure);

        // check BMA swaps
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

            Real expectedFraction = bmaData[i].rate / 100,
                estimatedFraction = swap.fairLiborFraction();
            Real error = std::fabs(expectedFraction - estimatedFraction);
            if (error > tolerance) {
                BOOST_ERROR(bmaData[i].n << " year(s) BMA swap:\n"
                    << std::setprecision(8)
                    << "\n estimated libor fraction: " << estimatedFraction
                    << "\n expected libor fraction:  " << expectedFraction
                    << "\n error:          " << error
                    << "\n tolerance:      " << tolerance);
            }
        }
    }

}

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
    else if (x == "Milliseconds"){
        return Milliseconds;
    }
    else {
        return Microseconds;
    }
    
}

LPOPER WINAPI testLinearZeroConsistency(XLOPER12 * depositData1,
                                XLOPER12 * fraData1,
                                XLOPER12 * immFutData1,
                                XLOPER12 *asxFutData1,
                                XLOPER12 *swapData1,
                                XLOPER12 *bondData1,
                                XLOPER12 * bmaData1) {
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
    CommonVars vars(depositData,fraData,immFutData,asxFutData,swapData,bondData,bmaData);
    
 //   testCurveConsistency<ZeroYield, Linear, IterativeBootstrap>(vars);
    RelinkableHandle<YieldTermStructure> curveHandle;
    curveHandle.linkTo(vars.termStructure);

    for (Size i = 0; i < vars.deposits; i++) {
        Euribor index(depositData[i].n * depositData[i].units, curveHandle);
        Rate expectedRate = depositData[i].rate / 100,
            estimatedRate = index.fixing(vars.today);
        if (std::fabs(expectedRate - estimatedRate) > tolerance) {
            BOOST_ERROR(
                depositData[i].n << " "
                << (depositData[i].units == Weeks ? "week(s)" : "month(s)")
                << " deposit:"
                << std::setprecision(8)
                << "\n    estimated rate: " << io::rate(estimatedRate)
                << "\n    expected rate:  " << io::rate(expectedRate));
        }
    }


    testBMACurveConsistency<ZeroYield, Linear, IterativeBootstrap>(vars);
    return &result;
}

// Calling convention *must* be WINAPI (aka __stdcall) for Excel.
/*
LPOPER WINAPI xll_function(_FP12 * arr)
{
	// Be sure to export your function.

#pragma XLLEXPORT
	static OPER result;

	try {
		double sub = 0;
		int size = arr->rows * arr->columns;
		for (int i = 0; i < size; i++) {
			sub = sub + arr->array[i];
		}
		result = sub;
		// OPER's act like Excel cells.
	}
	catch (const std::exception & ex) {
		XLL_ERROR(ex.what());

		result = OPER(xlerr::Num);
	}

	return &result;

}*/

