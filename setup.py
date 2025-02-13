# setup.py

from setuptools import setup

setup(
    name="file2txt_tool",
    version="0.1.0",
    py_modules=["file2txt"],  # We have just one Python file: file2txt.py
    install_requires=[
        # minimal "must have" libs. 
        # If you want optional extras for docx, PDF, etc., you can leave them out or do extras_require below.
    ],
    extras_require={
        "parsers": [
            "python-docx",
            "PyMuPDF",
            "PyPDF2",
            "openpyxl",
            "odfpy",
        ],
    },
    entry_points={
        "console_scripts": [
            # 'console_command_name = module_name:function_name'
            "file2txt=file2txt:main",
        ]
    },
    author="Your Name",
    description="CLI tool to convert various document formats to sanitized text files.",
    long_description=open("README.md", encoding="utf-8").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/<yourusername>/file2txt_tool",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",
)
