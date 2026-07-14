import os
import secrets

from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPBasic, HTTPBasicCredentials

security = HTTPBasic()


def require_auth(credentials: HTTPBasicCredentials = Depends(security)) -> str:
    """Fail closed: refuses to serve anything if the server isn't configured
    with credentials, rather than silently allowing unauthenticated access."""
    expected_user = os.environ.get("DASHBOARD_USERNAME")
    expected_pass = os.environ.get("DASHBOARD_PASSWORD")

    if not expected_user or not expected_pass:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="Server is missing DASHBOARD_USERNAME / DASHBOARD_PASSWORD configuration.",
        )

    user_ok = secrets.compare_digest(credentials.username, expected_user)
    pass_ok = secrets.compare_digest(credentials.password, expected_pass)
    if not (user_ok and pass_ok):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect username or password.",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username


# Scoped to just the /settings page — the rest of the app is intentionally
# open (per explicit instruction). Username is ignored; only the password
# is checked, since this is a single shared gate, not a per-user login.
SETTINGS_PASSWORD = "12345"


def require_settings_password(credentials: HTTPBasicCredentials = Depends(security)) -> str:
    if not secrets.compare_digest(credentials.password, SETTINGS_PASSWORD):
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Incorrect password.",
            headers={"WWW-Authenticate": "Basic"},
        )
    return credentials.username
