""" Fetches usage data CSV files from your PGE account so you can calculate your own bill under different scenarios

Use Environment variables for your pge.com credentials
PGE_USER = your username
"""
import requests
from openpyxl import load_workbook
from bs4 import BeautifulSoup
from datetime import date, datetime
import re
import os
from collections import defaultdict


def login():
    ...