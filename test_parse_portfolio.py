"""parse_portfolio_candidate_line 的單元測試。

聚焦於「損益平衡價優先採用」邏輯，特別是報酬率（%）不可被誤抓成成本的回歸案例。
無需 pytest 也可執行：

    .venv/bin/python test_parse_portfolio.py

若已安裝 pytest，亦可直接 `pytest test_parse_portfolio.py`。
"""

from __future__ import annotations

import math

from app import parse_portfolio_candidate_line


def _parse(line: str, multiplier: int = 1):
    return parse_portfolio_candidate_line(
        line,
        default_category="test",
        share_unit_multiplier=multiplier,
        image_name="t.png",
    )


def test_profit_rate_not_taken_as_cost():
    """有小數報酬率但無損益平衡價時，成本應取成本價，而非報酬率本身（P0 回歸）。"""
    row = _parse("台積電(2330) 1000 580.0 5.23%")
    assert row is not None
    assert row["symbol"] == "2330.TW"
    assert math.isclose(row["cost"], 580.0)
    assert math.isclose(row["shares"], 1000)


def test_break_even_preferred_over_cost_price():
    """報酬率後出現損益平衡價時，成本優先採用損益平衡價。"""
    row = _parse("台積電(2330) 1000 580.0 5.23% 581.5")
    assert row is not None
    assert math.isclose(row["cost"], 581.5)


def test_first_decimal_after_rate_not_last():
    """報酬率後有多個小數時，應取第一個（損益平衡價）而非最後一個（現價）。"""
    row = _parse("台積電(2330) 1000 580.0 5.23% 581.5 590.0")
    assert row is not None
    assert math.isclose(row["cost"], 581.5)


def test_implausible_break_even_filtered():
    """報酬率後若僅有離譜大數（如市值），應過濾並退回成本價。"""
    row = _parse("台積電(2330) 1000 580.0 5.23% 580000.0")
    assert row is not None
    assert math.isclose(row["cost"], 580.0)


def test_negative_profit_rate_not_taken_as_cost():
    """負報酬率不應被當成數值/成本。"""
    row = _parse("鴻海(2317) 2000 105.5 -3.2%")
    assert row is not None
    assert math.isclose(row["cost"], 105.5)
    assert math.isclose(row["shares"], 2000)


def test_plain_line_without_rate_still_works():
    """不含報酬率的一般行不受改動影響（向後相容回歸）。"""
    row = _parse("鴻海(2317) 2000 105.5")
    assert row is not None
    assert math.isclose(row["cost"], 105.5)
    assert math.isclose(row["shares"], 2000)


def test_share_unit_multiplier_applied():
    """以「張」為單位時，股數需乘上倍率。"""
    row = _parse("鴻海(2317) 2 105.5", multiplier=1000)
    assert row is not None
    assert math.isclose(row["shares"], 2000)


def test_header_line_rejected():
    """欄位標題列應被忽略。"""
    assert _parse("代號 名稱 股數 成本") is None


if __name__ == "__main__":
    tests = [obj for name, obj in sorted(globals().items()) if name.startswith("test_") and callable(obj)]
    failures = 0
    for test in tests:
        try:
            test()
            print(f"PASS {test.__name__}")
        except AssertionError as exc:
            failures += 1
            print(f"FAIL {test.__name__}: {exc}")
        except Exception as exc:  # noqa: BLE001 - 測試執行器需回報所有錯誤
            failures += 1
            print(f"ERROR {test.__name__}: {type(exc).__name__}: {exc}")
    print(f"\n{len(tests) - failures}/{len(tests)} passed")
    raise SystemExit(1 if failures else 0)
