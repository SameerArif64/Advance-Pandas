from setuptools import setup, find_packages

setup(
    name="advance_pandas",
    version="0.1.1",
    packages=find_packages(),
    install_requires = [line.strip() for line in open("requirements.txt") if line.strip()],
    author="Sameer Arif",
    author_email="supersameer64@gmail.com",
    description="Enhanced pandas DataFrame with async save, format retention, and backups.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/SameerArif64/Advance-Pandas",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.6",
    license="MIT",
)
