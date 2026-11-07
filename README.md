<div align="center">

# 📐 PBA-AutoPack (autocad_master)
### High-Precision Python COM Automation Engine & Vector Drafting Toolkit for Autodesk AutoCAD

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python Version](https://img.shields.io/badge/Python-3.7+-3776AB?logo=python&logoColor=white)](https://python.org)
[![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20AutoCAD-blue)](https://autodesk.com/autocad)
[![Organization](https://img.shields.io/badge/Powered%20By-Akash-00d4ff)](https://github.com/Powered-By-Akash)
[![Founder](https://img.shields.io/badge/Founder-Akash%20Tanna-ff007f)](https://github.com/akashtanna)
[![Engine](https://img.shields.io/badge/Interface-comtypes%20%2F%20COM%20API-10b981)](https://github.com/Powered-By-Akash/PBA-AutoPack)

**The production-grade CAD automation core powering automated engineering drawing production, structural detailing schedules, and dynamic drawing generators at Powered By Akash (PBA).**

---

</div>

## 📌 Overview

**PBA-AutoPack** (internally designated `autocad_master`) is a robust Python COM/ActiveX interface engine engineered specifically for civil, structural, and mechanical engineers who need programmatic control over Autodesk AutoCAD:
- **Resilient COM Retry Wrappers**: Automatically intercepts and recovers from transient Windows COM errors (such as `-2147418111: Rejected by callee`) when AutoCAD is executing heavy background operations.
- **Full Vector Geometry Primitive Suite**: Direct creation and manipulation of lines, continuous polylines, 2-point arcs, circles, triangles, and complex boundaries.
- **Advanced Automated Hatching**: Native SAFEARRAY `IDispatch` outer-loop boundary binding supporting SOLID, ANSI, and custom hatching patterns with scale and RGB TrueColor assignment.
- **Code-Compliant Dimensioning & Styles**: Programmatic creation of aligned, linear, and rotated dimension entities with automatic arrow scaling, text placement, and dimension style overrides.
- **Dynamic Linetype & Layer Management**: On-the-fly layer creation with TrueColor RGB mapping, lineweights, and automatic runtime fallback loading of missing linetypes from `acadiso.lin` or `acad.lin`.
- **Interactive User Coordinate System (UCS)**: Screen-click point polling and dynamic UCS transformation alignment.

---

## 🧩 Core API & Method Reference

| Method | Return Type | Description |
|---|---|---|
| `ConnectAutoCAD()` | `Instance` | Connects to the active AutoCAD application session, initializes active document and model space with exponential retry logic. |
| `_safe_com_call(action)` | `Any` | Executes COM calls wrapped in a micro-delay retry mechanism to safeguard against busy COM server crashes. |
| `true_color(rgb)` | `AcCmColor` | Dynamically loads and instantiates the version-matched AutoCAD TrueColor interface for precise RGB palettes. |
| `ensure_layer(name, rgb, weight, type)` | `Layer` | Validates layer existence, creates missing layers, assigns RGB tones, lineweights, and loads linetypes. |
| `draw_line(xy1, xy2, layer, rgb, ...)` | `Line` | Draws 2D/3D vector line with specific layer, color, linetype, and lineweight. |
| `draw_polyline(points, layer, rgb, ...)` | `Polyline` | Generates continuous 2D lightweight polylines from coordinate tuples with optional closure. |
| `draw_arc_2p(xy1, xy2, radius, ...)` | `Arc` | Mathematically derives center point and sweep angles to draw precise 2-point circular arcs. |
| `apply_hatch(boundary_obj, pattern, ...)` | `Hatch` | Binds pre-defined hatch patterns to closed boundaries via COM SAFEARRAY interfaces. |
| `join_objects(objects_list, width)` | `Polyline` | Joins contiguous lines and arcs into unified polylines and sets global widths. |
| `fillet_objects(obj1, obj2, radius)` | `None` | Automates corner filleting between adjacent drawing entities. |
| `draw_dimension(xy1, xy2, text_xy, ...)` | `Dimension` | Generates aligned or linear engineering dimensions with custom text and formatting. |
| `draw_text(string, xy, height, align, ...)` | `Text` | Injects single-line text entities with 9-point alignment constants (`TL`, `MC`, `BR`, etc.) and rotation angles. |
| `draw_leader(points, text, head_type, ...)` | `Leader` | Draws multi-vertex annotation leaders with customizable arrowheads. |
| `draw_circle(xy, radius, layer, rgb, ...)` | `Circle` | Draws circles with center coordinates and radius specification. |
| `set_ucs_by_click(app_instance)` | `None` | Interactive prompt requesting the user to click a drawing point to set a new UCS origin. |

---

## 🚀 Installation

### Option 1: Install via Pip from GitHub
```bash
pip install git+https://github.com/Powered-By-Akash/PBA-AutoPack.git
```

### Option 2: Clone and Install in Editable Mode
```bash
git clone https://github.com/Powered-By-Akash/PBA-AutoPack.git
cd PBA-AutoPack
pip install -e .
```

---

## 💻 Minimal Code Example

```python
from pba_autopack import ConnectAutoCAD

# 1. Connect to running AutoCAD instance
acad = ConnectAutoCAD()

# 2. Setup standard engineering layer
acad.ensure_layer(
    layer_name="PBA_BEAM_OUTLINE",
    rgb=(0, 212, 255),
    line_weight=35
)

# 3. Draw closed structural beam section
beam_profile = [
    (0, 0),
    (300, 0),
    (300, 600),
    (0, 600)
]
poly = acad.draw_polyline(
    points=beam_profile,
    layer_name="PBA_BEAM_OUTLINE",
    closed=True
)

# 4. Apply solid hatching
acad.apply_hatch(
    boundary_obj=poly,
    hatch_pattern="SOLID",
    layer_name="PBA_BEAM_OUTLINE",
    rgb=(20, 30, 45)
)

# 5. Add dimensional annotations
acad.draw_dimension(
    xy1=(0, 600),
    xy2=(300, 600),
    text_xy=(150, 650),
    dim_type="ALIGNED",
    layer_name="PBA_BEAM_OUTLINE"
)

# 6. Add label
acad.draw_text(
    text_string="BEAM B-101 (300x600)",
    xy=(150, -50),
    height=30.0,
    alignment="MC",
    layer_name="PBA_BEAM_OUTLINE"
)
```

---

## 📁 Repository Structure

```
PBA-AutoPack/
├── __init__.py           # Legacy root entry point for autocad_master
├── pba_autopack.py       # Primary framework module with ConnectAutoCAD engine
├── autocad_master.py     # Backward-compatibility bridge
├── setup.py              # Setuptools distribution config
├── pyproject.toml        # Modern PEP 517/518 build specification
├── LICENSE               # MIT Open Source License
├── .gitignore            # Git exclusions
├── README.md             # Documentation and usage guide
└── examples/
    └── demo.py           # Verification and usage demonstration
```

---

## 🏢 Organization & Founder

- **Organization:** [Powered By Akash (PBA)](https://github.com/Powered-By-Akash)
- **Founder & Lead Architect:** [Akash Tanna](https://github.com/akashtanna)
- **Location:** Gujarat, India
- **Official Website:** [poweredbyakash.in](https://www.poweredbyakash.in/)

---

## 📜 License

This project is licensed under the **MIT License** — see the [LICENSE](LICENSE) file for details.
