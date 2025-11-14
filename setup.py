# setup.py
from setuptools import setup, find_packages

setup(
    name="automation_batch",
    version="1.0",
    description="Internal Affinity Brand Partners Automated Order Batching Generation",
    author="David Kaufman",
    author_email="david.kaufman222@gmail.com",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    python_requires=">=3.11",
    install_requires=[],
)