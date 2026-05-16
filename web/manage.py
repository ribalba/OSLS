#!/usr/bin/env python3
"""Django management entry point for the OSLS web application."""

import os
import sys


def main():
    os.environ.setdefault("DJANGO_SETTINGS_MODULE", "osls_web.settings")
    from django.core.management import execute_from_command_line

    execute_from_command_line(sys.argv)


if __name__ == "__main__":
    main()
