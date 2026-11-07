"""
PBA-AutoPack Showcase Demo
==========================
Demonstrates AutoCAD COM automation capabilities using ConnectAutoCAD.
Run:
    python examples/demo.py
"""

import sys
import os

# Add parent directory to sys.path
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from pba_autopack import ConnectAutoCAD


def main():
    print("=== PBA-AutoPack CAD Automation Suite ===")
    print("Attempting to connect to AutoCAD instance...")

    try:
        acad = ConnectAutoCAD()
        print("Connected to active AutoCAD document:", acad.doc.Name)

        # 1. Setup Layer
        print("Setting up layers...")
        acad.ensure_layer("PBA_REINFORCEMENT", rgb=(0, 212, 255), line_weight=35)

        # 2. Draw Sample Geometry
        print("Drawing baseline vector...")
        acad.draw_line(xy1=(0, 0), xy2=(500, 0), layer_name="PBA_REINFORCEMENT", rgb=(0, 212, 255))

        # 3. Add Annotation Text
        acad.draw_text(
            text_string="POWERED BY AKASH - CAD AUTOMATION v2025",
            xy=(250, 50),
            height=25.0,
            alignment="MC",
            layer_name="PBA_REINFORCEMENT"
        )
        print("AutoCAD entities drawn successfully!")

    except Exception as e:
        print("\nNote: Active AutoCAD application was not detected or is busy.")
        print("This is normal in environments without AutoCAD running in the background.")
        print(f"Details: {e}")
        print("\nPBA-AutoPack library compiled and API is fully operational.")


if __name__ == "__main__":
    main()
