# -*- coding: utf-8 -*-
import re
from setuptools import find_packages, setup

# Utilities
with open("README.md") as readme_file:
    readme = readme_file.read()

with open("NEWS.rst") as history_file:
    history = history_file.read()
history = history.replace("\n-------------------", "\n^^^^^^^^^^^^^^^^^^^").replace("\n=====", "\n-----")


def find_version():
    result = re.search(r'{}\s*=\s*[\'"]([^\'"]*)[\'"]'.format("__version__"), open("autocalendar/__init__.py").read())
    return result.group(1)


# Dependencies
requirements = ["numpy", "pandas", "pickle", "os", "dateutil", "datetime", "apiclient", "google", "google_auth_oauthlib"]


# Setup
setup(

    # Info
    name="autocalendar",
    keywords="automation, calendar events, google calendar api, automatic scheduling, Python",
    url="https://github.com/zen-juen/AutoCalendar",
    download_url = 'https://github.com/zen-juen/AutoCalendar/zipball/master',
    version=find_version(),
    description="A Python automation scheduling system based on the Google Calendar API.",
    long_description=readme + "\n\n" + history,
    long_description_content_type="text/markdown",
    license="MIT license",

    # The name and contact of a maintainer
    author="Zen Juen Lau",
    author_email="lauzenjuen@gmail.com",

    # Dependencies
    install_requires=requirements,
#    setup_requires=setup_requirements,
#    extras_require={"test": test_requirements},
#    test_suite="pytest",
#    tests_require=test_requirements,

    # Misc
    packages=find_packages(),
    include_package_data=True,
    zip_safe=False,
    classifiers=[
        "Development Status :: 2 - Pre-Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Natural Language :: English",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
    ]
)
