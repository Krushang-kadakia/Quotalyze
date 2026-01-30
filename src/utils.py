import os
import tempfile
import uuid

DATA_ROOT = "data"

def get_tender_paths(tender_name: str, root_path: str = None) -> dict:
    """
    Returns a dictionary of paths for the given tender.
    If root_path is provided, uses that as base. Otherwise uses DATA_ROOT.
    """
    if root_path:
        base_dir = os.path.join(root_path, tender_name)
    else:
        base_dir = os.path.join(DATA_ROOT, tender_name)
        
    return {
        "base": base_dir,
        "reference": os.path.join(base_dir, "reference"),
        "vendors": os.path.join(base_dir, "vendors"),
        "output": os.path.join(base_dir, "output")
    }

def setup_tender_structure(tender_name: str, root_path: str = None) -> dict:
    """
    Creates the folder structure for a tender if it doesn't exist.
    Returns the paths.
    """
    paths = get_tender_paths(tender_name, root_path=root_path)
    
    for key, path in paths.items():
        if key == "base":
            os.makedirs(path, exist_ok=True)
        else:
            os.makedirs(path, exist_ok=True)
            
    return paths

def create_session_workspace() -> str:
    """
    Creates a unique temporary directory for the session.
    Returns absolute path to the workspace root.
    """
    # Create a persistent temp dir (not auto-cleanup on object deletion, 
    # dependent on system cleanup or explicit cleanup)
    # We use mkdtemp to ensure it persists for the session duration
    return tempfile.mkdtemp(prefix="quotalyze_session_")

def list_tenders() -> list:
    """Deprecated: Lists existing tenders in the data directory."""
    if not os.path.exists(DATA_ROOT):
        return []

    return [d for d in os.listdir(DATA_ROOT) if os.path.isdir(os.path.join(DATA_ROOT, d))]


def detect_boq_index(sheet_names: list) -> int:
    """Returns index of best guess for BOQ sheet."""
    keywords = ['boq', 'bill of quantities', 'price', 'pricing', 'quote', 'quotation', 'abstract', 'schedule']
    for i, name in enumerate(sheet_names):
        lname = name.lower()
        if any(k in lname for k in keywords):
            return i
    return 0 # Default to first sheet

