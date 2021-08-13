from setuptools import setup
import os.path

# from Export Python
def get_version(version_tuple):
    return ".".join(map(str, version_tuple))

init = os.path.join(os.path.dirname(__file__), "docx2md", "__init__.py")
version_line = list(filter(lambda l: l.startswith("VERSION"), open(init)))[0]
VERSION = get_version(eval(version_line.split("=")[-1]))

readme = os.path.join(os.path.dirname(__file__), "README.rst")
README = open(readme).read()

setup(
    name="docx2md",
    version=VERSION,
    author="Toshihiko Ichida",
    author_email="dogatana@gmail.com",
    url="https://github.com/dogatana/docx2md",
    packages=["docx2md"],
    description="convert .docx to .md",
    long_description=README,
    install_requires=[
        "lxml",
        "Pillow",
    ],
    classifiers = [
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: Developers",
        "Environment :: Other Environment",
        "Intended Audience :: Other Audience",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Topic :: Text Editors",
        "Topic :: Text Editors :: Documentation",
        "Topic :: Text Editors :: Word Processors",
        "Topic :: Text Processing",
        "Topic :: Text Processing :: Markup",
        "Topic :: Text Processing :: Markup :: Markdown",
        "Topic :: Utilities",
        ],

)