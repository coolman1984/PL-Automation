from src.validation import numbers_match


def test_numbers_match_within_tolerance():
    assert numbers_match(100.0, 100.009, 0.01)
    assert not numbers_match(100.0, 100.011, 0.01)


def test_blank_values_only_match_each_other():
    assert numbers_match(None, None, 0.01)
    assert not numbers_match(None, 0.0, 0.01)

