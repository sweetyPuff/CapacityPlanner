from captool.months import parse_month


def test_normal_formats():
    assert parse_month("2026'8") == ("2026-08", None)
    assert parse_month("2026-08") == ("2026-08", None)
    assert parse_month("2026/12") == ("2026-12", None)
    assert parse_month("202608") == ("2026-08", None)


def test_typo_header_recovered_with_warning():
    value, warning = parse_month("20267'")
    assert value == "2026-07"
    assert warning is not None and "20267'" in warning


def test_datetime_input():
    from datetime import datetime
    assert parse_month(datetime(2026, 9, 1)) == ("2026-09", None)


def test_unparseable():
    value, warning = parse_month("N/A")
    assert value is None
    assert warning is not None


def test_none():
    value, warning = parse_month(None)
    assert value is None
    assert warning is not None
