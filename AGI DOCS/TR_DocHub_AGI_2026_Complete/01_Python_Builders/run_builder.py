#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""임시 실행 스크립트"""
import sys
import os

# 현재 스크립트의 디렉토리로 이동
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

# 통합빌더.py 실행
with open('통합빌더.py', 'r', encoding='utf-8') as f:
    code = compile(f.read(), '통합빌더.py', 'exec')
    exec(code)
