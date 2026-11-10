# -*- coding: utf-8 -*-
import sys
sys.path.append(r'D:\TOAN\05_CONG_CU_VA_ENGINE\Cham_Diem_Hoc_Sinh')
import cham_diem_core

students = cham_diem_core.STUDENT_RESULTS
for s in students:
    stt = s['stt']
    name = s['name']
    code = s['code']
    print(f"{stt:2d} | {name:22s} | {code}")
