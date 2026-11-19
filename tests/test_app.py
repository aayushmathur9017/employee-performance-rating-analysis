"""
Test Suite: Application Infrastructure, Module Imports & UI Configuration
Employee Performance Rating Analysis
Student: Aayush Mathur | Enrolment: 2500320140001 | Section: MCA-B | ABES Engineering College
Target: Streamlit Configuration (.streamlit/config.toml), src/ modules, UI Components
"""

import os
import sys
import importlib
import pytest
import tomli

BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)


# -----------------------------------------------------------------------------
# TC-T-15.1: Module Imports & Syntax Validity
# -----------------------------------------------------------------------------
def test_module_imports():
    """Verify core application modules import cleanly without syntax or dependency errors."""
    modules_to_test = [
        "src.ui_components",
        "src.generate_data",
        "src.data_loader",
        "src.cleaner",
        "src.analysis",
    ]
    for mod_name in modules_to_test:
        mod = importlib.import_module(mod_name)
        assert mod is not None, f"Failed to import module {mod_name}"


# -----------------------------------------------------------------------------
# TC-T-15.2: Streamlit Configuration Integrity
# -----------------------------------------------------------------------------
def test_streamlit_config_integrity():
    """Verify .streamlit/config.toml exists, parses as valid TOML, and defines design system tokens."""
    config_path = os.path.join(BASE_DIR, ".streamlit", "config.toml")
    assert os.path.isfile(config_path), f"config.toml not found at {config_path}"
    
    with open(config_path, "rb") as f:
        config = tomli.load(f)
    
    assert "theme" in config, "Missing [theme] section in config.toml"
    assert "server" in config, "Missing [server] section in config.toml"
    
    theme = config["theme"]
    assert theme.get("primaryColor") == "#3B82F6", "Incorrect primaryColor in theme"
    assert theme.get("backgroundColor") == "#F8FAFC", "Incorrect backgroundColor in theme"
    assert theme.get("secondaryBackgroundColor") == "#FFFFFF", "Incorrect secondaryBackgroundColor in theme"
    assert theme.get("textColor") == "#1E293B", "Incorrect textColor in theme"
    
    server = config["server"]
    assert server.get("headless") is True, "Server should be configured headless"


# -----------------------------------------------------------------------------
# TC-T-15.3: UI Helper Functions & Badge Rendering
# -----------------------------------------------------------------------------
def test_ui_components_callable():
    """Verify all reusable UI helper functions are callable and properly exported."""
    from src import ui_components
    
    required_callables = [
        "inject_tailwind_css",
        "render_header",
        "render_kpi_card",
        "render_insight_box",
        "render_alert",
        "render_card",
        "render_badge",
    ]
    for func_name in required_callables:
        assert hasattr(ui_components, func_name), f"ui_components missing function '{func_name}'"
        assert callable(getattr(ui_components, func_name)), f"'{func_name}' is not callable"


def test_ui_badge_rendering():
    """Verify render_badge produces styled HTML with proper color mapping and content."""
    from src.ui_components import render_badge
    
    badge_html = render_badge("Active Performer", color="green")
    assert "<span" in badge_html
    assert "Active Performer" in badge_html
    assert "#DCFCE7" in badge_html  # green background
    assert "#166534" in badge_html  # green text
    
    blue_badge = render_badge("Review 2024", color="blue")
    assert "Review 2024" in blue_badge
    assert "#DBEAFE" in blue_badge


# -----------------------------------------------------------------------------
# TC-T-15.4: Workspace Project Directory & Asset Structure
# -----------------------------------------------------------------------------
def test_project_workspace_structure():
    """Verify directory structure and existence of core deliverables."""
    required_dirs = ["data", "docs", "src", "tests", ".streamlit"]
    for d in required_dirs:
        dir_path = os.path.join(BASE_DIR, d)
        assert os.path.isdir(dir_path), f"Required directory '{d}' is missing"
    
    required_files = [
        os.path.join(BASE_DIR, "data", "app.db"),
        os.path.join(BASE_DIR, "data", "departments.csv"),
        os.path.join(BASE_DIR, "data", "employees.csv"),
        os.path.join(BASE_DIR, "data", "performance_reviews.csv"),
        os.path.join(BASE_DIR, "docs", "04_Traceability_Matrix.xlsx"),
    ]
    for f in required_files:
        assert os.path.isfile(f), f"Required file '{f}' is missing"
        assert os.path.getsize(f) > 0, f"File '{f}' is empty (0 bytes)"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
