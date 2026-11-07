from setuptools import setup
from pathlib import Path

this_directory = Path(__file__).parent
long_description = (this_directory / "README.md").read_text(encoding="utf-8") if (this_directory / "README.md").exists() else ""

setup(
    name="pba-autopack",
    version="2025.1.0",
    description="Python COM Automation Engine & High-Precision CAD Drafting Toolkit for AutoCAD",
    long_description=long_description,
    long_description_content_type="text/markdown",
    author="Akash Tanna",
    author_email="Poweredbyakash@gmail.com",
    url="https://github.com/Powered-By-Akash/PBA-AutoPack",
    project_urls={
        "Organization": "https://github.com/Powered-By-Akash",
        "Website": "https://www.poweredbyakash.in/",
        "Source": "https://github.com/Powered-By-Akash/PBA-AutoPack",
        "Tracker": "https://github.com/Powered-By-Akash/PBA-AutoPack/issues",
    },
    py_modules=["pba_autopack", "autocad_master"],
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Intended Audience :: Manufacturing",
        "Topic :: Scientific/Engineering",
        "Topic :: Software Development :: Libraries :: Python Modules",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Operating System :: Microsoft :: Windows",
    ],
    python_requires=">=3.7",
    keywords="autocad cad automation comtypes com pyautocad engineering drafting pba",
    install_requires=[
        "comtypes",
    ],
)
