from pathlib import Path


def not_implemented_process(report_name: str):
    """Build a process() for a report whose business logic hasn't been provided yet.

    Keeps the upload/date/generate flow fully wired for every report from day
    one, without inventing any real processing logic for reports 2-9.
    """

    def process(input_files: dict, dates: dict, output_dir: Path) -> Path:
        raise NotImplementedError(
            f"'{report_name}' processing logic hasn't been implemented yet. "
            "The upload/date/download flow for this report works end-to-end — "
            "it just needs its real business logic wired in."
        )

    return process
