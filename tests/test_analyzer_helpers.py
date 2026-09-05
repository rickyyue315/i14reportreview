import pandas as pd

from views.analyzer_helpers import (
    build_hybrid_params,
    canonical_site_codes,
    filter_by_site_code,
    hybrid_params_changed,
)


def test_hybrid_params_snapshot_and_stale_detection():
    params = build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False)
    assert not hybrid_params_changed(params, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False))
    assert hybrid_params_changed(params, build_hybrid_params(
        14, 100, {"Tourist": 8, "Mix": 9, "Non Tourist": 10}, False))
    assert hybrid_params_changed(params, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, True))
    assert hybrid_params_changed(params, build_hybrid_params(
        15, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False))


def test_hybrid_params_dimension_default_and_stale_detection():
    base = build_hybrid_params(14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False)
    assert base["dimension"] == "nature"  # 預設維度（舊呼叫未傳維度仍相容）
    assert not hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "nature"))
    assert hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "rp"))
    assert hybrid_params_changed(base, build_hybrid_params(
        14, 100, {"Tourist": 7, "Mix": 9, "Non Tourist": 10}, False, "both"))


def test_site_code_options_are_canonical_strings():
    assert canonical_site_codes([1002, "1001", " 1002 ", None, "", " B01 "]) == ["1001", "1002", "B01"]


def test_site_code_filter_handles_numeric_and_string_values():
    frame = pd.DataFrame({"Site Code": [1001, "1002", " 1001 ", None], "SKU": [1, 2, 3, 4]})
    filtered = filter_by_site_code(frame, "1001")
    assert list(filtered["SKU"]) == [1, 3]
    assert len(filter_by_site_code(frame, "全部")) == 4
    assert filter_by_site_code(frame, "missing").empty
