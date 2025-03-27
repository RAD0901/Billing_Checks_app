from setuptools import setup, find_packages
import os

setup(
    name="BillingCheckTool",
    version="1.0.0",
    description="A tool for processing billing CSV files.",
    author="Your Name",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    include_package_data=True,
    install_requires=[
        "tkinter",  # Ensure tkinter is installed
    ],
    entry_points={
        "gui_scripts": [
            "billing-check-tool=gui:main",
        ],
    },
    data_files=[
        ("ico", [os.path.join("ico", "Amecor_Logo-01-small.ico")]),
    ],
)
