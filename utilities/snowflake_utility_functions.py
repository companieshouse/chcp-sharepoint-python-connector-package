"""Helper functions for managing snowflake sessions for the prevent model."""

import os
from snowflake.snowpark import Session

def get_snowpark_session(role: str | None = None) -> Session:
    """
    Returns an active Snowpark session.

    - If running inside Snowflake (worksheet, notebook, Streamlit),
      returns the active session.
    - If running locally, creates a session using connection details.
    """

    # ---- Case 1: Running inside Snowflake ----
    try:
        session = Session.get_active_session()
        if session is not None:
            return session
    except Exception:
        pass  # Not running inside Snowflake

    # ---- Case 2: Running locally ----

    if role is None:
        print("No role provided for local Snowflake session; using default connection name.")
        role = os.getenv("SNOWFLAKE_CONNECTION_NAME")

    session = Session.builder.config("connection_name", role).create()

    return session
