import glob
import nbformat
from nbformat import NO_CONVERT

def strip_notebook(path):
    try:
        nb = nbformat.read(path, as_version=NO_CONVERT)
    except Exception:
        return False
    changed = False
    for cell in nb.get("cells", []):
        if cell.get("outputs"):
            cell["outputs"] = []
            changed = True
        if cell.get("execution_count") is not None:
            cell["execution_count"] = None
            changed = True
    # Optionally clear kernel/execution metadata that records output timestamps
    metadata = nb.get("metadata", {})
    if metadata.get("execution", None) is not None:
        metadata.pop("execution", None)
        nb["metadata"] = metadata
        changed = True
    if changed:
        nbformat.write(nb, path)
    return changed

def main():
    modified = False
    for path in glob.glob("**/*.ipynb", recursive=True):
        if strip_notebook(path):
            modified = True
    # Exit 0 either way; CI will detect git status to decide to commit
    return 0

if __name__ == "__main__":
    raise SystemExit(main())