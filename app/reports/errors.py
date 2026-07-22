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


def describe_missing_columns(actual_columns, required_columns, filename: str):
    """Like describe_column_mismatch, but only flags MISSING columns --
    extra columns beyond `required_columns` are fine. Use this instead of
    describe_column_mismatch whenever a file's actual full column set
    isn't confirmed exhaustive/stable (e.g. only a handful of columns are
    read by name out of a much larger real export, or the export format
    has changed over time) -- describe_column_mismatch's stricter "no
    unexpected extras" check will reject perfectly valid real files in
    that case. (Confirmed the hard way: an API Vehicles export legitimately
    lacked 2 columns a docstring had assumed were always present but the
    code never actually used, which the stricter check rejected.)"""
    missing = sorted(set(required_columns) - set(actual_columns))
    if not missing:
        return None
    return f"'{filename}' is missing required columns: {missing}."
