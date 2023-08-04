from setuptools import setup

setup(
    name="excel_text_match_highlighter",
    version="0.1",
    author="Cheffey",
    author_email="not@reachable.com",
    description="compare two excel columns and highlight the identical matching parts",
    packages=["match_highlighter"],
    install_requires=['pandas', 'xlsxwriter'],
)
