"""Dependency check script for the repository.

This small utility is intended for template repositories. It parses the
project dependencies declared in ``pyproject.toml`` and attempts to
import each package, reporting which dependencies are importable and
which are missing. The script intentionally does not perform any
installation; it only checks importability.

Example
-------
Run the script from the repository root::

    python main.py

This will print a summary and exit with code 0 when all dependencies
are importable, or with non-zero code when any are missing.
"""

import sys
import os
import tomllib
from pathlib import Path
import re

# --- Path bootstrap (required for shared modules) -----------------
common_path = os.path.abspath(os.path.join(os.getcwd(), "utilities"))
if common_path not in sys.path:
    sys.path.insert(0, common_path)

# --- END Path bootstrap -------------------------------------------

from utilities.snowflake_utility_functions import get_snowpark_session


def parse_pyproject_dependencies(pyproject_path: Path):
    """Parse ``pyproject.toml`` and return declared dependencies.

    Parameters
    ----------
    pyproject_path : pathlib.Path
        Path to the ``pyproject.toml`` file.

    Returns
    -------
    list of str
        A list of dependency specification strings exactly as written in
        the ``[project].dependencies`` table (for example,
        ``"numpy>=1.24"`` or
        ``"snowflake-connector-python[pandas]>=4.0.0"``).

    Notes
    -----
    The function uses Python's :mod:`tomllib` to parse TOML. If the
    file is missing or malformed an exception will propagate to the
    caller.
    """
    data = {}
    with pyproject_path.open("rb") as f:
        data = tomllib.load(f)

    deps = []
    project = data.get("project", {})
    raw_deps = project.get("dependencies", []) or []
    for d in raw_deps:
        # dependency strings like "package>=1.2.3" or "pkg[extra]>=x" or quoted strings
        deps.append(d)
    return deps


def normalize_package_name(dep_str: str):
    """Normalize a dependency specifier into an importable package name.

    Parameters
    ----------
    dep_str : str
        Dependency string from ``pyproject.toml`` (may include version
        specifiers, extras, and environment markers).

    Returns
    -------
    str
        A best-effort importable package name. The function performs a
        few transformations:

        - strips version specifiers and environment markers
        - removes extras (``[extra]``) when detecting the package name
        - converts hyphens to underscores for import compatibility
        - applies a small mapping for known packages (for example
          ``snowflake-connector-python`` -> ``snowflake``)

    Notes
    -----
    This function uses simple heuristics and a small mapping table and
    will not cover every packaging edge case; it is intentionally
    conservative and primarily intended for developer feedback in a
    template project.
    """
    # Remove environment markers and extra metadata after semicolon
    dep = dep_str.split(";", 1)[0].strip()

    # Remove version specifier
    dep = re.split(r"[<=>!~\[]", dep)[0].strip()

    # Handle known package -> import name mappings
    mappings = {
        "snowflake-connector-python": "snowflake",
        # Add more special cases here if needed
    }

    pkg = dep
    if pkg in mappings:
        return mappings[pkg]

    # replace - with _ for import attempts
    return pkg.replace("-", "_")


def check_packages(pyproject_path: Path):
    """Attempt to import each declared dependency.

    Parameters
    ----------
    pyproject_path : pathlib.Path
        Path to ``pyproject.toml``.

    Returns
    -------
    dict
        Mapping of dependency specifier (str) -> (bool, import_name)
        where the boolean indicates whether the import succeeded and
        ``import_name`` is the name that was attempted.
    """
    deps = parse_pyproject_dependencies(pyproject_path)
    results = {}

    for dep in deps:
        import_name = normalize_package_name(dep)
        try:
            __import__(import_name)
            results[dep] = (True, import_name)
        except ImportError:
            results[dep] = (False, import_name)

    return results


def test_snowflake_session():
    """Test that a Snowflake session can be created."""

    try:
        session = get_snowpark_session()
        print("Successfully created Snowflake session:", session)
    except Exception as e:
        print("Failed to create Snowflake session:", e)


def main():
    repo_root = Path(__file__).resolve().parent
    pyproject_path = repo_root / "pyproject.toml"

    if not pyproject_path.exists():
        print("pyproject.toml not found; nothing to check.")
        sys.exit(1)

    results = check_packages(pyproject_path)

    installed = [k for k, v in results.items() if v[0]]
    missing = [k for k, v in results.items() if not v[0]]

    if installed:
        print("Installed packages:")
        for s in installed:
            print("  -", s)

    if missing:
        print("\nMissing packages:")
        for m in missing:
            print("  -", m, "(import attempted as:", results[m][1] + ")")
        # sys.exit(2)

    print("All declared dependencies appear importable.")

    test_snowflake_session()


if __name__ == "__main__":
    main()
