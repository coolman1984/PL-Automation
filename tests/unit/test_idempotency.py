from src.errors import ExistingActualColumnError


def test_idempotency_error_is_a_specific_domain_error():
    error = ExistingActualColumnError("A08 already exists")

    assert error.code == "existing_actual_column"
    assert "A08" in str(error)
