# -*- coding: utf-8 -*-
"""
구글플레이 영수증 CSV 파서.

Gmail API 연동 없이, 매달 Claude한테 "이번 달 구글플레이 영수증 찾아줘"라고
요청하면 이 형식의 CSV 파일을 만들어서 input/ 폴더에 넣어준다.
naver_pay.py / coupang.py 와 마찬가지로 별도 거래 행을 만들지 않고,
(날짜, 시각, 금액) -> 상품명 조회용 목록만 만든다.

CSV 형식 (헤더 포함):
  결제일시,상품명,금액
  2026-08-29 22:06:32,ChatGPT Plus (ChatGPT),29000
"""
import csv
import datetime


def extract_entries(path):
    """[(date, hour, minute, amount, product_name), ...] 목록을 돌려준다."""
    entries = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            pay_dt = (row.get("결제일시") or "").strip()
            product = (row.get("상품명") or "").strip()
            amount_str = (row.get("금액") or "").strip()
            if not pay_dt or not amount_str:
                continue
            dt = _parse_datetime(pay_dt)
            if dt is None:
                continue
            try:
                amount = int(amount_str.replace(",", ""))
            except ValueError:
                continue
            entries.append((dt.date(), dt.hour, dt.minute, amount, product))
    return entries


def _parse_datetime(value):
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M"):
        try:
            return datetime.datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None
