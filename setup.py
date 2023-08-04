from setuptools import setup

setup(
    name="match_highlighter",
    version="0.1",
    author="Cheffey",
    author_email="not@reachable.com",
    description="compare two excel columns and highlight the identical matching parts",
    packages=["excel_text_match_highlighter"],
    install_requires=['pandas', 'xlsxwriter'],
)
