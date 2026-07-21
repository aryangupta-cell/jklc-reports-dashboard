class ReportProcessingError(Exception):
    """Expected, user-facing processing failure (bad/missing file, wrong format, missing column).

    Route handlers show this message directly to the user. Anything else that
    goes wrong is an unexpected bug and gets logged server-side with a generic
    message shown to the user instead.
    """


def describe_column_mismatch(actual_columns, expected_columns, filename: str):
    """Returns a message naming exactly which columns are missing and/or
    unexpectedly extra in `filename` versus `expected_columns` -- or None if
    there's no mismatch. Lets a report's own file-load step tell the user
    precisely what's wrong (which file, which columns) instead of a generic
    "something went wrong"."""
    actual = set(actual_columns)
    expected = set(expected_columns)
    missing = sorted(expected - actual)
    extra = sorted(actual - expected)
    if not missing and not extra:
        return None
    parts = []
    if missing:
        parts.append(f"missing columns: {missing}")
    if extra:
        parts.append(f"unexpected extra columns: {extra}")
    return f"'{filename}' has a column mismatch -- " + "; ".join(parts) + "."
