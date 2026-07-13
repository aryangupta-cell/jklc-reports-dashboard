class ReportProcessingError(Exception):
    """Expected, user-facing processing failure (bad/missing file, wrong format, missing column).

    Route handlers show this message directly to the user. Anything else that
    goes wrong is an unexpected bug and gets logged server-side with a generic
    message shown to the user instead.
    """
