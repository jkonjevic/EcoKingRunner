"""Entry point kept for muscle memory: ``python ecoking_web_launcher.py``.

The UI itself lives in :mod:`ecoking.webapp` so the same server can run locally
and inside the hosted container.
"""

from ecoking.webapp import main

if __name__ == "__main__":
    main()
