#!/usr/bin/env python3
"""
PPT Master MCP Server Entry Point
================================
Provides standard stdio MCP tools, resources, and prompts for creating, inspecting,
validating, and converting native PowerPoint presentations.
"""

import os
import sys
import subprocess
from typing import List, Optional
from mcp.server.fastmcp import FastMCP

# Initialize FastMCP Server
mcp = FastMCP("ppt-master")

# Root directory of ppt-master
REPO_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
VENV_PYTHON = os.path.join(REPO_ROOT, ".venv", "bin", "python")
if not os.path.exists(VENV_PYTHON):
    VENV_PYTHON = sys.executable

def run_script(script_relative_path: str, args: List[str]) -> dict:
    """Helper to run internal ppt-master python scripts safely."""
    script_path = os.path.join(REPO_ROOT, "skills", "ppt-master", "scripts", script_relative_path)
    cmd = [VENV_PYTHON, script_path] + args
    try:
        result = subprocess.run(
            cmd,
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            check=False
        )
        return {
            "success": result.returncode == 0,
            "exit_code": result.returncode,
            "stdout": result.stdout.strip(),
            "stderr": result.stderr.strip()
        }
    except Exception as e:
        return {
            "success": False,
            "exit_code": -1,
            "stdout": "",
            "stderr": str(e),
            "error": str(e)
        }

# -----------------------------------------------------------------------------
# MCP Tools
# -----------------------------------------------------------------------------

@mcp.tool()
def ppt_init_project(name: str, format: str = "ppt169") -> str:
    """
    Initialize a new PPT Master project workspace under projects/ directory.
    
    Args:
        name: Name slug of the project (e.g. 'quarterly_review')
        format: Canvas format, default 'ppt169' (16:9 1280x720)
    """
    res = run_script("project_manager.py", ["init", name, "--format", format])
    if res["success"]:
        return f"[SUCCESS] Project '{name}' initialized successfully.\n{res['stdout']}"
    else:
        return f"[ERROR] Failed to initialize project:\n{res['stderr'] or res['stdout']}"

@mcp.tool()
def ppt_import_sources(project_path: str, sources: List[str]) -> str:
    """
    Import document sources (PDF, Word, Excel, Web URLs) and convert to Markdown.
    
    Args:
        project_path: Absolute or relative path to project directory
        sources: List of file paths or web URLs to import
    """
    args = ["import-sources", project_path] + list(sources)
    res = run_script("project_manager.py", args)
    if res["success"]:
        return f"[SUCCESS] Imported sources into {project_path}.\n{res['stdout']}"
    else:
        return f"[ERROR] Failed to import sources:\n{res['stderr'] or res['stdout']}"

@mcp.tool()
def ppt_scaffold_spec(project_path: str) -> str:
    """
    Scaffold design_spec.md and spec_lock.md for a project.
    
    Args:
        project_path: Path to project directory
    """
    res1 = run_script("project_manager.py", ["scaffold-spec", project_path])
    res2 = run_script("project_manager.py", ["scaffold-lock", project_path])
    if res1["success"] and res2["success"]:
        return f"[SUCCESS] Scaffolding created for {project_path}.\n{res1['stdout']}\n{res2['stdout']}"
    else:
        return f"[ERROR] Scaffolding error:\n{res1['stderr'] or res1['stdout']}\n{res2['stderr'] or res2['stdout']}"

@mcp.tool()
def ppt_validate_project(project_path: str) -> str:
    """
    Validate project structure, design_spec.md placeholders, and spec_lock.md schema.
    
    Args:
        project_path: Path to project directory
    """
    res = run_script("project_manager.py", ["validate", project_path])
    if res["success"]:
        return f"[VALIDATED] Project structure is valid!\n{res['stdout']}"
    else:
        return f"[INVALID] Project validation failed:\n{res['stdout']}\n{res['stderr']}"

@mcp.tool()
def ppt_check_svg_quality(project_path: str) -> str:
    """
    Run automated SVG quality checker for bounds, typography, WCAG contrast, and layout rules.
    
    Args:
        project_path: Path to project directory
    """
    res = run_script("svg_quality_checker.py", [project_path])
    if res["success"]:
        return f"[QUALITY OK] All SVG slides passed check!\n{res['stdout']}"
    else:
        return f"[QUALITY ISSUES] SVG Quality Checker output:\n{res['stdout']}\n{res['stderr']}"

@mcp.tool()
def ppt_render_latex(project_path: str) -> str:
    """
    Render LaTeX formula manifests (images/formula_manifest.json) into 300 DPI vector PNGs.
    
    Args:
        project_path: Path to project directory
    """
    res = run_script("latex_render.py", [project_path])
    if res["success"]:
        return f"[SUCCESS] LaTeX formulas rendered.\n{res['stdout']}"
    else:
        return f"[ERROR] LaTeX rendering failed:\n{res['stderr'] or res['stdout']}"

@mcp.tool()
def ppt_sync_icons(project_path: str, icons: List[str]) -> str:
    """
    Copy icon vector files from bundled libraries (e.g. 'tabler-outline/cpu') into project icons.
    
    Args:
        project_path: Path to project directory
        icons: List of icon library paths (e.g. ['tabler-outline/cpu', 'tabler-outline/layers'])
    """
    args = [project_path] + list(icons)
    res = run_script("icon_sync.py", args)
    if res["success"]:
        return f"[SUCCESS] Synced icons.\n{res['stdout']}"
    else:
        return f"[ERROR] Failed to sync icons:\n{res['stderr'] or res['stdout']}"

@mcp.tool()
def ppt_convert_to_pptx(project_path: str, native_charts_tables: bool = False, transition: str = "fade") -> str:
    """
    Convert authored SVG slide pages into native DrawingML editable PowerPoint (.pptx) presentation.
    
    Args:
        project_path: Path to project directory
        native_charts_tables: Replace SVG fallbacks with Excel-backed native PowerPoint chart objects
        transition: Transition effect name (e.g. 'fade', 'push', 'wipe', 'morph')
    """
    args = [project_path, "-t", transition]
    res = run_script("svg_to_pptx.py", args)
    if res["success"]:
        return f"[EXPORT SUCCESS] PPTX presentation generated!\n{res['stdout']}"
    else:
        return f"[EXPORT ERROR] Failed to convert SVG to PPTX:\n{res['stderr'] or res['stdout']}"

@mcp.tool()
def ppt_enhance_native_pptx(pptx_path: str, project_name: str) -> str:
    """
    Initialize a native existing-PPTX enhancement project for applying OOXML transitions and narrations.
    
    Args:
        pptx_path: Path to target PPTX presentation file
        project_name: Name slug for enhancement workspace
    """
    out_dir = os.path.join("projects", project_name)
    res = run_script("pptx_to_svg.py", [pptx_path, "-o", out_dir, "--inheritance-mode", "both", "--roundtrip"])
    if res["success"]:
        return f"[ENHANCE INIT SUCCESS] Created roundtrip workspace at {out_dir}:\n{res['stdout']}"
    else:
        return f"[ENHANCE ERROR] Failed to initialize enhancement workspace:\n{res['stderr'] or res['stdout']}"

# -----------------------------------------------------------------------------
# MCP Resources
# -----------------------------------------------------------------------------

@mcp.resource("ppt://projects/{project_name}/design_spec")
def get_design_spec(project_name: str) -> str:
    """Read the human-readable design specification for a project."""
    path = os.path.join(REPO_ROOT, "projects", project_name, "design_spec.md")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return f"Error: design_spec.md not found in project '{project_name}'"

@mcp.resource("ppt://projects/{project_name}/spec_lock")
def get_spec_lock(project_name: str) -> str:
    """Read the machine-readable execution lock contract for a project."""
    path = os.path.join(REPO_ROOT, "projects", project_name, "spec_lock.md")
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return f.read()
    return f"Error: spec_lock.md not found in project '{project_name}'"

# -----------------------------------------------------------------------------
# MCP Prompts
# -----------------------------------------------------------------------------

@mcp.prompt()
def generate_presentation(topic_or_sources: str, style_preference: str = "Glassmorphism SaaS") -> str:
    """
    Prompt template for guiding the AI in converting source material into a native PowerPoint presentation.
    """
    return (
        f"You are driving the PPT Master MCP server to create a presentation deck.\n"
        f"Source Material/Topic: {topic_or_sources}\n"
        f"Design Style: {style_preference}\n\n"
        f"Steps to execute using MCP tools:\n"
        f"1. Call `ppt_init_project(name='...')` to create a project workspace.\n"
        f"2. Call `ppt_import_sources` if source files exist.\n"
        f"3. Call `ppt_scaffold_spec` and generate design_spec.md and spec_lock.md.\n"
        f"4. Author 16:9 SVG slides into projects/<name>/svg_output/.\n"
        f"5. Call `ppt_check_svg_quality` to verify layout bounds and contrast.\n"
        f"6. Call `ppt_convert_to_pptx` to compile into an editable .pptx deck.\n"
    )

# -----------------------------------------------------------------------------
# Main Runner
# -----------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
