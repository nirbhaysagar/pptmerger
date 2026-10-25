"""
api/index.py - Serverless entrypoint for Vercel functions
"""
import os
import sys

# Ensure project root directory is in sys.path
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from app import handler, app, application, PPTUnifyHandler

__all__ = ["handler", "app", "application", "PPTUnifyHandler"]
