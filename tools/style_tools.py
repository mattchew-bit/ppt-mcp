"""
Style analysis and profiling tools for PowerPoint MCP Server.

Provides MCP tools to analyze presentation styles, create reusable
style profiles, and apply them to new presentations.
"""

from typing import Any, Dict, Optional
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

# Style profiles stored in memory (keyed by profile name)
_style_profiles: Dict[str, Any] = {}


def register_style_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register style analysis and profiling tools with the FastMCP app."""

    @app.tool(
        annotations=ToolAnnotations(
            title="Analyze Presentation Style",
        ),
    )
    def analyze_presentation_style(file_path: str) -> Dict:
        """Analyze a PowerPoint file to extract fonts, colors, layouts, and text hierarchy.

        Returns a comprehensive style analysis including:
        - Primary font and font usage patterns
        - Color palette with usage contexts
        - Layout positioning patterns
        - Text hierarchy (title/subtitle/body styles)
        - Consistency score (0-1)

        Args:
            file_path: Path to the .pptx file to analyze
        """
        from utils.style_utils import analyze_presentation

        try:
            analysis = analyze_presentation(file_path)
            return {
                "message": f"Analyzed {analysis['slide_count']} slides from {file_path}",
                "primary_font": analysis["fonts"]["primary_font"],
                "font_count": len(analysis["fonts"]["font_usage"]),
                "color_count": analysis["colors"]["total_unique_colors"],
                "top_colors": analysis["colors"]["primary_palette"][:5],
                "common_sizes": analysis["fonts"]["common_sizes"][:5],
                "text_hierarchy": analysis.get("text_hierarchy", {}),
                "consistency_score": analysis["consistency_score"],
                "slide_dimensions": f"{analysis['slide_width']}x{analysis['slide_height']} inches",
                "total_shapes": analysis["shapes"]["total_shapes"],
                "full_analysis": analysis,
            }
        except FileNotFoundError as e:
            return {"error": str(e)}
        except Exception as e:
            return {"error": f"Analysis failed: {str(e)}"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Create Style Profile",
        ),
    )
    def create_style_profile(file_path: str, profile_name: str) -> Dict:
        """Analyze a presentation and create a named, reusable style profile.

        The profile captures fonts, colors, layout patterns, and text hierarchy
        that can later be applied to new presentations or referenced for
        consistent formatting.

        Args:
            file_path: Path to the .pptx file to profile
            profile_name: Name for the style profile (e.g., 'firm_standard', 'client_acme')
        """
        from utils.style_utils import analyze_presentation, create_profile

        try:
            analysis = analyze_presentation(file_path)
            profile = create_profile(analysis, profile_name)
            _style_profiles[profile_name] = profile

            return {
                "message": f"Created style profile '{profile_name}' from {file_path}",
                "profile_name": profile_name,
                "primary_font": profile.primary_font,
                "title_font_size": profile.title_font_size,
                "body_font_size": profile.body_font_size,
                "color_count": len(profile.color_palette),
                "consistency_score": profile.consistency_score,
            }
        except FileNotFoundError as e:
            return {"error": str(e)}
        except Exception as e:
            return {"error": f"Failed to create profile: {str(e)}"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Save Style Profile",
        ),
    )
    def save_style_profile(profile_name: str, output_path: str) -> Dict:
        """Save a style profile to a JSON file for reuse across sessions.

        Args:
            profile_name: Name of a previously created profile
            output_path: Path where the JSON profile will be saved
        """
        from utils.style_utils import save_profile

        if profile_name not in _style_profiles:
            available = list(_style_profiles.keys()) or ["(none)"]
            return {
                "error": f"Profile '{profile_name}' not found. Available: {', '.join(available)}"
            }

        try:
            profile = _style_profiles[profile_name]
            save_profile(profile, output_path)
            return {
                "message": f"Saved profile '{profile_name}' to {output_path}",
                "profile_name": profile_name,
                "output_path": output_path,
            }
        except Exception as e:
            return {"error": f"Failed to save profile: {str(e)}"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Load Style Profile",
        ),
    )
    def load_style_profile(file_path: str) -> Dict:
        """Load a previously saved style profile from a JSON file.

        The loaded profile is available for reference and can be used to
        guide formatting of new presentations.

        Args:
            file_path: Path to the JSON profile file
        """
        from utils.style_utils import load_profile

        try:
            profile = load_profile(file_path)
            _style_profiles[profile.name] = profile
            return {
                "message": f"Loaded profile '{profile.name}' from {file_path}",
                "profile_name": profile.name,
                "primary_font": profile.primary_font,
                "title_font_size": profile.title_font_size,
                "body_font_size": profile.body_font_size,
                "color_palette_count": len(profile.color_palette),
                "consistency_score": profile.consistency_score,
            }
        except FileNotFoundError:
            return {"error": f"Profile file not found: {file_path}"}
        except Exception as e:
            return {"error": f"Failed to load profile: {str(e)}"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Style Profile",
        ),
    )
    def get_style_profile(profile_name: str) -> Dict:
        """Get the full details of a loaded style profile.

        Returns all extracted style information including font hierarchy,
        color palette, and layout patterns.

        Args:
            profile_name: Name of a loaded profile
        """
        from dataclasses import asdict

        if profile_name not in _style_profiles:
            available = list(_style_profiles.keys()) or ["(none)"]
            return {
                "error": f"Profile '{profile_name}' not found. Available: {', '.join(available)}"
            }

        profile = _style_profiles[profile_name]
        return {
            "profile": asdict(profile) if hasattr(profile, "__dataclass_fields__") else profile,
            "profile_name": profile_name,
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="List Style Profiles",
        ),
    )
    def list_style_profiles() -> Dict:
        """List all style profiles currently loaded in memory."""
        profiles = []
        for name, profile in _style_profiles.items():
            if hasattr(profile, "primary_font"):
                profiles.append({
                    "name": name,
                    "primary_font": profile.primary_font,
                    "source": profile.source_file,
                    "consistency_score": profile.consistency_score,
                })
            else:
                profiles.append({"name": name, "type": "dict"})

        return {
            "profiles": profiles,
            "total": len(profiles),
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Apply Style to Presentation",
        ),
    )
    def apply_style_profile(
        profile_name: str,
        presentation_id: Optional[str] = None,
    ) -> Dict:
        """Apply a style profile's fonts and colors to the current presentation.

        Updates all text in the presentation to use the profile's primary font,
        title size, and body size. Does not change content — only formatting.

        Args:
            profile_name: Name of a loaded style profile
            presentation_id: ID of presentation to style (default: current)
        """
        from pptx.util import Pt

        if profile_name not in _style_profiles:
            return {"error": f"Profile '{profile_name}' not found"}

        profile = _style_profiles[profile_name]
        pres_id = presentation_id or get_current_presentation_id()
        if not pres_id or pres_id not in presentations:
            return {"error": "No presentation loaded. Create or open one first."}

        pres = presentations[pres_id]
        modified_runs = 0

        for slide in pres.slides:
            for shape in slide.shapes:
                if not (hasattr(shape, "text_frame") and shape.text_frame):
                    continue

                top = shape.top.inches if shape.top else 5
                is_title = top < 2 and len(shape.text_frame.text) < 100

                for para in shape.text_frame.paragraphs:
                    for run in para.runs:
                        run.font.name = profile.primary_font
                        if is_title:
                            run.font.size = Pt(profile.title_font_size)
                        else:
                            run.font.size = Pt(profile.body_font_size)
                        modified_runs += 1

        return {
            "message": f"Applied profile '{profile_name}' to presentation '{pres_id}'",
            "font_applied": profile.primary_font,
            "title_size": profile.title_font_size,
            "body_size": profile.body_font_size,
            "runs_modified": modified_runs,
        }
