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

OTHER_SPENDING_DECREASE_THRESHOLD = -10.0
OTHER_SPENDING_GROWTH_THRESHOLD = 5.0
OTHER_ACTIVE_DECREASE_THRESHOLD = -10.0
OTHER_ACTIVE_GROWTH_THRESHOLD = 5.0
OTHER_CR_DECREASE_THRESHOLD = -10.0
OTHER_CR_GROWTH_THRESHOLD = 5.0


def classify_other_spending(sp_diff_pct: float) -> str:
    if sp_diff_pct < OTHER_SPENDING_DECREASE_THRESHOLD:
        return "DO"
    if sp_diff_pct > OTHER_SPENDING_GROWTH_THRESHOLD:
        return "GO"
    return "SO"


def classify_other_active(active_diff_pct: float) -> str:
    if active_diff_pct < OTHER_ACTIVE_DECREASE_THRESHOLD:
        return "DO"
    if active_diff_pct > OTHER_ACTIVE_GROWTH_THRESHOLD:
        return "GO"
    return "SO"


def classify_other_conversion(cr_diff_pct: float) -> str:
    if cr_diff_pct < OTHER_CR_DECREASE_THRESHOLD:
        return "D"
    if cr_diff_pct > OTHER_CR_GROWTH_THRESHOLD:
        return "G"
    return "S"


OTHER_CATEGORY_DECISIONS = {'O: DO: SO: G': {'decision': 'Negative impact', 'next_step': 'The increase in conversion to a New Paid lister, as well as stable Active listers, show that the price change failed to affect users in the desired way, they only lowered the average bill for launching Ads. In total, this is a negative impact.Next steps: return the previous prices or roll them back by 50% of the increase (decide with the teamlead). And in both cases, look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: DO: SO: D': {'decision': 'Negative impact', 'next_step': 'A stable number of Active Listers indicates that the listers were not influenced in the planned way, they only decreased ARPPU, which leads to a decrease in revenue. This result is considered negative. Next steps: It is necessary to return to the previous prices and look for other methods of influencing users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: DO: SO: S': {'decision': 'No impact', 'next_step': 'Stable both Active Listers and Сonversion rate show that the users were not influenced in the planned way, they only decreased Spendings.Therefore, the experiment cannot be considered successful and is closed with the status No impact. Next steps: return the previous prices or roll them back by 50% of the increase (decide with the teamlead). And in both cases, look for other methods of influencing users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: DO: DO: G': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: DO: DO: D': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: DO: DO: S': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users."}, 'O: DO: GO: G': {'decision': 'Negative impact', 'next_step': "The increase in both Active Listers and Conversions to New Paid lister indicates that the price change didn't have the desired effect on users behaviour, so the experiment cannot be considered a success. Decrease in Spendings against the background of stable Active listers indicates a negative impact: the listers were not influenced, and revenues decreased. The next step is to return prices to the previous values and look for other methods of influencing user behaviour.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: DO: GO: D': {'decision': 'Negative impact', 'next_step': 'The growth of Active Listers indicates that the price change failed to influence user behaviour in the desired way, so the experiment cannot be considered a success. Decrease in Spendings on the background of stable listers leads to a negative impact in general: listers were not influenced, revenues decreased. Next step: bring prices back to previous values and look for other methods of influencing users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: DO: GO: S': {'decision': 'Negative impact', 'next_step': 'The growth of Active Listers and the stable Сonversion indicate that the price change failed to influence the behaviour of users in the desired way, so the experiment cannot be considered a success. Decrease in Spendings against the background of stable listers leads to a negative impact in general: listers could not be influenced, and revenues decreased. Next step: bring prices back to previous values and look for other methods of influencing users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: SO: SO: G': {'decision': 'No impact', 'next_step': 'The stable number of Active Listers and the increase in Conversion to the New Paid lister indicate that the price change did not influence the behaviour of the users in the desired way, so the experiment is closed with the No impact status. Two options for the next step are possible (discuss and choose with your teamlead): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of impact on users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: SO: SO: D': {'decision': 'No impact', 'next_step': 'The stable number of Active Listers indicates that the price change did not influence the behaviour of the users in the desired way, so the experiment is closed with the No impact status. Two options for the next step are possible (discuss and choose with your teamlead): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of impact on users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: SO: SO: S': {'decision': 'No impact', 'next_step': 'Stable both Active listers and Conversion to a New Paid ister indicate that the price change failed to affect the behavior of users in the desired way, so the experiment is closed with the No impact status.There are two possible options for the next step (discuss and choose with the teamlead): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: SO: DO: G': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: SO: DO: D': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: SO: DO: S': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: SO: GO: G': {'decision': 'No impact', 'next_step': 'The increase in both Active Listers and Conversion to New Paid lister indicates that the price change did not have the desired impact on users behaviour. Since the Spendings remain stable and the impact on posting has not been achieved, the experiment is closed with the No impact status.There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Raise prices again by +10% in a new experiment;2. Leave the prices and look for other methods of impact on users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: SO: GO: D': {'decision': 'No impact', 'next_step': 'The growth of Active listers indicates that the price change failed to affect the behavior of users in the desired way. Since the Spendings remain stable, and it was not possible to influence the posting, the experiment is closed with the No impact status.There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: SO: GO: S': {'decision': 'No impact', 'next_step': 'The growth of Active listers with a stable Conversion to a New Paid lister indicate that the price change failed to affect the behavior of users in the desired way. Since the Spendings remain stable, and it was not possible to influence the posting, the experiment is closed with the No impact status.There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: GO: SO: G': {'decision': 'No impact', 'next_step': 'The stable number of Active listers and the increase in Сonversion to a New Paid lister indicate that the price change failed to affect the behavior of users in the desired way. The experiment is closed with the No impact status.There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: GO: SO: D': {'decision': 'No impact', 'next_step': 'The stable number of Active listers indicates that the price change failed to affect the behavior of users in the desired way. The experiment is closed with the No impact status.There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: GO: SO: S': {'decision': 'No impact', 'next_step': 'The stable number of both Active listers and Сonversion to a New Paid lister indicate that the price change failed to affect the behavior of users in the desired way. The experiment is closed with the No impact status.There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: GO: DO: G': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: GO: DO: D': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: GO: DO: S': {'decision': 'Positive impact', 'next_step': "The decrease in Active listers indirectly indicates that we have managed to influence users in the planned way. The preliminary status of the experiment is Positive impact, prices remain at the established level. To confirm this, you need to check the total category at +1 level from the Other one (for example, you changed prices in the category Other women's clothing, the category at +1 level from it: Women's clothing). The number of Active listers, launched campaigns (and preferably Spendigs) should remain stable (or grow) in total category. Check it out, and write the results into the experiment.If the indicators fall, it will mean that the listers have left not only the Other category, but also the overall total category. In this case, the experiment is closed with the Negative impact status. The next step is to return the prices to the previous values and look for another method of influencing the users.For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else."}, 'O: GO: GO: G': {'decision': 'No impact', 'next_step': 'The growth of both Active listers and Conversion to a New Paid lister indicates that the price change failed to affect the behavior of users in the desired way. The experiment is closed with the No impact status. There are two possible options for the next step (choose with a team leader if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: GO: GO: D': {'decision': 'No impact', 'next_step': 'The growth of Active listers indicates that the price change failed to affect the behavior of users in the desired way. The experiment is closed with the No impact status. There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}, 'O: GO: GO: S': {'decision': 'No impact', 'next_step': 'The growth of Active listers and the stable Conversion to a New Paid lister indicate that the price change failed to affect the behavior of users in the desired way. The experiment is closed with the No impact status.There are two possible options for the next step (choose with a teamlead if you are not sure yourself): 1. Re-increase prices by +10% in a new experiment;2. Leave the prices and look for other methods of influencing users. For example, ask moderation to transfer listings, launch an experiment with mailing lists, train listers to place ads in relevant categories, or something else.'}}

def get_other_category_decision(code: str):
    return OTHER_CATEGORY_DECISIONS.get(
        code,
        {
            "decision": "Need review",
            "next_step": "No Other category rule found for this combination yet.",
        },
    )



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
    active_diff = diff(active_before, active_after)

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

    if is_other_category:
        other_sp = classify_other_spending(sp_diff)
        other_active = classify_other_active(active_diff)
        other_cr = classify_other_conversion(cr_diff)
        decision_code = f"O: {other_sp}: {other_active}: {other_cr}"
        decision_result = get_other_category_decision(decision_code)
        final_decision = decision_result["decision"]
        next_step = decision_result["next_step"]
    else:
        decision_code = f"{npl_code}{sp_code}{cr_code}"
        decision_result = get_decision(decision_code)
        final_decision = decision_result["decision"]
        next_step = decision_result["next_step"]

        if force_low_npl or npl_after < 10:
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
        "active_diff": active_diff,
        "npl_code": npl_code,
        "sp_code": sp_code,
        "cr_code": cr_code,
        "decision_code": decision_code,
        "final_decision": final_decision,
        "next_step": next_step,
        "low_conversion_warning": low_conversion_warning,
        "low_conversion_message": low_conversion_message,
    }