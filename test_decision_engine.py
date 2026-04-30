"""Tests for Other category decision matrix (decision_engine.analyze_category)."""

from decision_engine import OTHER_CATEGORY_DECISIONS, analyze_category

STUB_FINAL = "Other category - needs separate logic"


def _assert_other_row(result, *, decision_code: str, final_decision: str):
    assert result["decision_code"] == decision_code
    assert result["final_decision"] == final_decision
    assert result["next_step"], "next_step must be non-empty"
    assert STUB_FINAL not in result["final_decision"]
    assert STUB_FINAL not in result["next_step"]


def test_other_do_so_g_negative_impact():
    """Spendings DO, Active SO, Conversion G → Negative impact."""
    r = analyze_category(
        npl_before=10,
        npl_after=15,
        sp_before=100,
        sp_after=85,
        active_before=100,
        active_after=102,
        is_other_category=True,
    )
    _assert_other_row(r, decision_code="O: DO: SO: G", final_decision="Negative impact")


def test_other_do_do_g_positive_impact():
    """Spendings DO, Active DO, Conversion G → Positive impact."""
    r = analyze_category(
        npl_before=10,
        npl_after=14,
        sp_before=100,
        sp_after=85,
        active_before=100,
        active_after=85,
        is_other_category=True,
    )
    _assert_other_row(r, decision_code="O: DO: DO: G", final_decision="Positive impact")


def test_other_so_so_s_no_impact():
    """Spendings SO, Active SO, Conversion S → No impact."""
    r = analyze_category(
        npl_before=10,
        npl_after=10,
        sp_before=100,
        sp_after=102,
        active_before=100,
        active_after=101,
        is_other_category=True,
    )
    _assert_other_row(r, decision_code="O: SO: SO: S", final_decision="No impact")


def test_other_go_go_g_no_impact():
    """Spendings GO, Active GO, Conversion G → No impact."""
    r = analyze_category(
        npl_before=10,
        npl_after=12,
        sp_before=100,
        sp_after=110,
        active_before=100,
        active_after=110,
        is_other_category=True,
    )
    _assert_other_row(r, decision_code="O: GO: GO: G", final_decision="No impact")


def test_other_go_do_d_positive_impact():
    """Spendings GO, Active DO, Conversion D → Positive impact."""
    r = analyze_category(
        npl_before=10,
        npl_after=7,
        sp_before=100,
        sp_after=106,
        active_before=100,
        active_after=85,
        is_other_category=True,
    )
    _assert_other_row(r, decision_code="O: GO: DO: D", final_decision="Positive impact")


def test_other_category_low_npl_not_insufficient():
    """Other category must not apply Low NPL → Insufficient data override."""
    r = analyze_category(
        npl_before=3,
        npl_after=4,
        sp_before=100,
        sp_after=85,
        active_before=100,
        active_after=102,
        geo="default",
        force_low_npl=True,
        is_other_category=True,
    )
    assert r["final_decision"] != "Insufficient data"
    assert r["decision_code"] == "O: DO: SO: G"
    assert r["next_step"]
    assert STUB_FINAL not in r["final_decision"]


def test_other_category_matrix_complete():
    """All PDF-derived keys exist (27 combinations)."""
    assert len(OTHER_CATEGORY_DECISIONS) == 27
    for code, row in OTHER_CATEGORY_DECISIONS.items():
        assert code.startswith("O: ")
        assert row.get("decision")
        assert row.get("next_step")
