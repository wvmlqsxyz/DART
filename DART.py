
import os, json, threading, time
from tkinter import Tk, Frame, Label, Entry, Button, Listbox, Scrollbar, StringVar, END, SINGLE, messagebox, filedialog, ttk, VERTICAL
import requests, zipfile
from io import BytesIO
import xml.etree.ElementTree as ET
import pandas as pd
import numpy as np 

# ---------- 설정 ----------
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CORP_LIST_FILE = os.path.join(SCRIPT_DIR, "dart_corplist.json")
CONFIG_FILE = os.path.join(SCRIPT_DIR, "dart_config.json")
MAPPING_FILE = os.path.join(SCRIPT_DIR, "dart_mapping.json")
KEY_XLSX_FILE = os.path.join(SCRIPT_DIR, "key.xlsx") 
KEY_CSV_FILE = os.path.join(SCRIPT_DIR, "key.csv")   

REPORT_TYPES = {
    "사업보고서": "11001",
    "분기보고서": "11013",
    "반기보고서": "11012",
}
FS_DIVS = {
    "연결": "CFS",
    "개별/별도": "OFS",
}

# ---------- 유틸 (안정화) ----------
def normalize_key(s):
    """모든 공백 제거 및 소문자 변환을 통해 키를 표준화합니다."""
    if s is None: return ""
    # 모든 공백(띄어쓰기, 탭 등)을 제거하고, 앞뒤 공백도 제거 후 소문자 변환
    return str(s).replace(" ", "").strip().lower() 
    
def parse_num(s):
    """문자열에서 숫자만 추출하여 float으로 변환. 실패 시 None 반환."""
    try:
        if s is None: return None
        s_str = str(s).replace(",", "").strip()
        if s_str in ["", "-", "N/A", "ERR"]: return None
        return float(s_str)
    except: 
        return None

def sdiv(a,b):
    """안전한 나누기 헬퍼: 분자/분모 중 None이거나 분모가 0이면 None 반환"""
    return None if a is None or b in [None,0] else a/b
    
def safe_ratio(numerator, denominator, multiplier=100):
    """
    안전하게 나누기를 수행하고 곱셈을 처리합니다. 
    결과가 None이면 None을 반환하여 TypeError를 방지합니다.
    """
    val = sdiv(numerator, denominator)
    return val * multiplier if val is not None else None

def get_safe_account_name(item):
    """항목(item)에서 계정명을 가져와 None, float 등에 상관없이 항상 문자열로 반환"""
    name = item.get("account_nm") or item.get("account_nm_kor")
    if name is not None:
        return str(name).strip() 
    return ""
