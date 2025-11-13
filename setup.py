"""
Setup script for ExcelBot Pro
"""

from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = [line.strip() for line in fh if line.strip() and not line.startswith("#")]

setup(
    name="excelbot-pro",
    version="1.0.0",
    author="Mandar Bahadarpurkar",
    author_email="",
    description="A chatbot interface for automating Excel tasks using VBA with GitHub integration",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/mandarab76/ExcelBotPro",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
    ],
    python_requires=">=3.8",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "excelbot=excelbot_chat:main",
        ],
    },
)
