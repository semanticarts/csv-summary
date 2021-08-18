import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="csv_summary",
    version="1.1.1",
    author="Boris Pelakh",
    author_email="boris.pelakh@semanticarts.com",
    description="CSV Summary Tool",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/semanticarts/csv-summary",
    packages=setuptools.find_packages(),
    install_requires=[
        'openpyxl>=3.0.6'
    ],
    classifiers=[
        "Development Status :: 4 - Beta",
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: Apache Software License",
        "Operating System :: OS Independent",
    ],
    entry_points={
        "console_scripts": [
            "csv_summary=csv_summary.main:summarize_csv"
        ]
    },
    python_requires='>=3.7',
)
