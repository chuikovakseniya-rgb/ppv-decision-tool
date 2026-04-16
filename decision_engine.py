def safe_divide(a, b):
    return a / b if b != 0 else 0


def diff(before, after):
    return ((after - before) / before * 100) if before != 0 else 0


def classify_change(value, growth_threshold=5, decline_threshold=-10):
    if value > growth_threshold:
        return "G"
    if value < decline_threshold:
        return "D"
    return "S"


GEO_THRESHOLDS = {
    "default": {
        "npl": {"growth": 5, "decline": -10},
        "sp": {"growth": 5, "decline": -10},
        "cr": {"growth": 5, "decline": -10},
    },
    "KG": {
        "npl": {"growth": 5, "decline": -10},
        "sp": {"growth": 5, "decline": -10},
        "cr": {"growth": 11, "decline": -10},
    },
    "AZ": {
        "npl": {"growth": 3, "decline": -8},
        "sp": {"growth": 3, "decline": -8},
        "cr": {"growth": 3.5, "decline": -10},
    },
    # RS thresholds are ambiguous in source files (e.g. CR 2.5 vs 3.9; NPL/SP as 9999).
    # Keep fallback to default until business rule is clarified.
}


def decode_status(code):
    return {
        "G": "Growth",
        "S": "Stable",
        "D": "Decrease",
    }.get(code, "Unknown")


def get_decision(code):
    matrix = {
        "DSG": {
            "decision": "No impact",
            "next_step": "Restart the experiment and look for a better price point. Attract Active listers.",
        },
        "DSD": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "DSS": {
            "decision": "No impact",
            "next_step": "Keep the prices and work on attracting Active listers.",
        },
        "DDG": {
            "decision": "Negative impact",
            "next_step": "Close current experiment as Negative impact. Start a new experiment and roll back prices by 50% of the implemented change. Work to increase Active listers.",
        },
        "DDD": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values since all metrics are falling.",
        },
        "DDS": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "DGG": {
            "decision": "Positive impact",
            "next_step": "Keep the prices or try to raise them by 10–20% in a new experiment. Attract Active listers.",
        },
        "DGD": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "DGS": {
            "decision": "Positive impact",
            "next_step": "Keep the prices and work on attracting Active listers.",
        },
        "SSG": {
            "decision": "No impact",
            "next_step": "Restart the experiment in search of the optimal price leading to Spendings growth. Attract Active listers.",
        },
        "SSD": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "SSS": {
            "decision": "No impact",
            "next_step": "Choose with your teamlead: raise prices by 10–20%, lower prices by 10–20%, or leave and watch.",
        },
        "SDG": {
            "decision": "No impact",
            "next_step": "Close current experiment as No impact. Start a new experiment and roll back prices by 50% of the implemented change. Attract Active listers.",
        },
        "SDD": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "SDS": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "SGG": {
            "decision": "Positive impact",
            "next_step": "Keep the prices. Repeated price increase up to 10% is allowed.",
        },
        "SGD": {
            "decision": "Positive impact",
            "next_step": "Keep the prices and work on increasing launched PPV amount. It is better not to raise prices anymore.",
        },
        "SGS": {
            "decision": "Positive impact",
            "next_step": "Keep the prices.",
        },
        "GGG": {
            "decision": "Positive impact",
            "next_step": "Keep the prices. Repeated price increase up to 10% is allowed.",
        },
        "GGD": {
            "decision": "Positive impact",
            "next_step": "Keep the prices, work on increasing launched PPV amount, and attract New Paid listers. It is better not to raise prices anymore.",
        },
        "GGS": {
            "decision": "Positive impact",
            "next_step": "Keep the prices.",
        },
        "GSG": {
            "decision": "No impact",
            "next_step": "Conversion and New Paid listers are growing, but Spendings are stable. Watch one more week and, if needed, adjust prices in a new experiment.",
        },
        "GSD": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "GSS": {
            "decision": "No impact",
            "next_step": "Close current experiment as No impact. Start a new experiment and roll back prices by 50% of the implemented change. Attract Active listers.",
        },
        "GDG": {
            "decision": "Negative impact",
            "next_step": "Close current experiment as Negative impact. Start a new experiment and roll back prices by 50% of the implemented change. Attract Active listers.",
        },
        "GDD": {
            "decision": "Negative impact",
            "next_step": "Return prices to their original values.",
        },
        "GDS": {
            "decision": "Negative impact",
            "next_step": "Close current experiment as Negative impact. Start a new experiment and roll back prices by 50% of the implemented change. Attract Active listers.",
        },
    }

    return matrix.get(
        code,
        {
            "decision": "Need review",
            "next_step": "No rule found yet. Add this combination to the matrix.",
        },
    )


def analyze_category(
    npl_before,
    npl_after,
    sp_before,
    sp_after,
    active_before,
    active_after,
    geo="default",
    force_low_npl=False,
    is_other_category=False,
):
    geo_key = str(geo).upper() if geo else "default"
    thresholds = GEO_THRESHOLDS.get(geo_key, GEO_THRESHOLDS["default"])
    # RS thresholds are ambiguous in source files, so fallback to default is used.
    conversion_reference = thresholds["cr"]["growth"] / 100

    npl_diff = diff(npl_before, npl_after)
    sp_diff = diff(sp_before, sp_after)

    cr_before = safe_divide(npl_before, active_before)
    cr_after = safe_divide(npl_after, active_after)
    cr_diff = diff(cr_before, cr_after)

    npl_code = classify_change(
        npl_diff,
        growth_threshold=thresholds["npl"]["growth"],
        decline_threshold=thresholds["npl"]["decline"],
    )
    sp_code = classify_change(
        sp_diff,
        growth_threshold=thresholds["sp"]["growth"],
        decline_threshold=thresholds["sp"]["decline"],
    )
    cr_code = classify_change(
        cr_diff,
        growth_threshold=thresholds["cr"]["growth"],
        decline_threshold=thresholds["cr"]["decline"],
    )

    decision_code = f"{npl_code}{sp_code}{cr_code}"
    decision_result = get_decision(decision_code)
    final_decision = decision_result["decision"]
    next_step = decision_result["next_step"]

    # TODO: Add separate decision matrix for Other category.
    if is_other_category:
        final_decision = "Other category - needs separate logic"
        next_step = "Use separate decision logic for Other category."
    elif force_low_npl or npl_after < 10:
        final_decision = "Insufficient data"
        next_step = (
            "The number of New Paid Listers is too small to draw a conclusion. "
            "Choose a longer period or decide with the teamlead."
        )

    low_conversion_warning = cr_after < conversion_reference
    low_conversion_message = ""
    if low_conversion_warning:
        low_conversion_message = (
            "Conversion after is below GEO reference. You may try reducing prices and working on "
            "attracting New Paid Listers, while not allowing Spendings to decrease."
        )

    return {
        "npl_diff": npl_diff,
        "sp_diff": sp_diff,
        "cr_diff": cr_diff,
        "npl_code": npl_code,
        "sp_code": sp_code,
        "cr_code": cr_code,
        "decision_code": decision_code,
        "final_decision": final_decision,
        "next_step": next_step,
        "low_conversion_warning": low_conversion_warning,
        "low_conversion_message": low_conversion_message,
    }