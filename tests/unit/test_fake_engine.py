from src.agent_contracts import TargetRef
from src.fake_engine import FakeEngine


def test_fake_engine_reads_and_writes_rectangular_ranges():
    engine = FakeEngine({"Data": {"1,1": "old"}})
    target = TargetRef("working-copy", sheet="Data", address="A1:B2")

    assert engine.read_values(target) == [["old", None], [None, None]]
    engine.write_values(target, [[1, 2], [3, 4]])

    assert engine.read_values(target) == [[1, 2], [3, 4]]
    assert engine.calls == [("read_values", "A1:B2"), ("write_values", "A1:B2"), ("read_values", "A1:B2")]


def test_fake_engine_rejects_shape_mismatch():
    engine = FakeEngine({"Data": {}})
    target = TargetRef("working-copy", sheet="Data", address="A1:B2")

    try:
        engine.write_values(target, [[1]])
    except ValueError as exc:
        assert "shape" in str(exc)
    else:
        raise AssertionError("shape mismatch was accepted")

