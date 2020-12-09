# Dependencies
import DateTime
import datetime
import platform
import dateutil
import pandas as pd
import numpy as np
import pickle
import google_auth_oauthlib
import google.auth
import pkg_resources

# Info
__version__ = "0.0.6"


# Maintainer info
__author__ = "Zen Juen Lau"
__email__ = "lauzenjuen@gmail.com"


# Citation
__bibtex__ = r"""
@misc{AutoCalendar,
  doi = {10.5281/ZENODO.4113733},
  url = {https://github.com/zen-juen/AutoCalendar},
  author = {Lau, Zen J},
  title = {AutoCalendar: A Python automation scheduling system based on the Google Calendar API},
  publisher = {Zenodo},
  month={Oct},
  year = {2020},
}
"""

__cite__ = (
    """
You can cite AutoCalendar as follows:

- Lau, Z. J. (2020). AutoCalendar: A Python automation scheduling system based on the Google Calendar API. Retrieved """
    + datetime.date.today().strftime("%B %d, %Y")
    + """, from https://github.com/zen-juen/AutoCalendar


Full bibtex reference:
"""
    + __bibtex__
)
# Aliases for citation
__citation__ = __cite__


# =============================================================================
# Helper functions to retrieve info
# =============================================================================
def cite(silent=False):
    """Cite AutoCalendar (prints bibtex and APA reference).

    Examples
    ---------
    >>> import autocalendar as autocalendar
    >>>
    >>> autocalendar.cite()

    """
    if silent is False:
        print(__cite__)
    else:
        return __bibtex__


def version(silent=False):
    """AutoCalendar's version.
    This function retrieves the version of the package.

    Examples
    ---------
    >>> import autocalendar as autocalendar
    >>>
    >>> autocalendar.version()


    """
    if silent is False:
        print(
            "- OS: " + platform.system(),
            "(" + platform.architecture()[1] + " " + platform.architecture()[0] + ")",
            "\n- Python: " + platform.python_version(),
            "\n\n- NumPy: " + np.__version__,
            "\n- Pandas: " + pd.__version__,
            "\n- datetime: " + pkg_resources.get_distribution("DateTime").version,
            "\n- dateutil: " + dateutil.__version__,
            "\n- pickle: " + pickle.format_version,
            "\n- google-auth-oauthlib: " + pkg_resources.get_distribution("google_auth_oauthlib").version,
            "\n- google-auth: " + pkg_resources.get_distribution("google_auth_oauthlib").version,
            "\n- apiclient: " + pkg_resources.get_distribution("google-api-python-client").version,
        )
    else:
        return __version__
