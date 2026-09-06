# -*- coding: utf-8 -*-
"""
쿠팡 결제 영수증 모음 PDF 파서 (주문 1건당 1페이지).

이 파일도 naver_pay.py와 마찬가지로 별도 거래 행을 만들지 않고,
(날짜, 시각, 금액) -> 상품명 조회용 목록만 만든다.
"""
import datetime
import re

import pdfplumber

_DT_RE = re.compile(r"거래일시\s*(\d{4}/\d{2}/\d{2})\s+(\d{2}):(\d{2}):\d{2}")
_AMOUNT_RE = re.compile(r"합계금액\s*([\d,]+)\s*원")
_PRODUCT_RE = re.compile(r"상품명\s*(.+?)\s*과세금액", re.DOTALL)
_APPROVAL_RE = re.compile(r"승인번호\s*(\S+)")


def extract_entries(path):
    """[(date, hour, minute, amount, product_name), ...] 목록을 돌려준다.

    한 주문에 서로 다른 판매자(마켓플레이스 셀러)의 상품이 섞여 있으면, 쿠팡이
    영수증을 판매자별로 여러 페이지로 나눠서 발행한다. 이때 거래일시/승인번호는
    페이지마다 똑같이 찍히지만, 합계금액은 판매자(페이지)별 결제 금액이고
    실제 카드/계좌에서는 이 금액들이 합산되어 한 번에 청구된다. 상품명도
    페이지마다 다르므로, 같은 승인번호끼리는 상품명은 합치고 금액은 더해서
    하나의 항목으로 만든다.

    병합 기준은 주문번호가 아니라 승인번호다. 주문번호는 쿠팡이 정의하는
    "주문" 단위일 뿐이고, 승인번호가 실제로 카드사에 찍히는 "결제 승인" 단위와
    1:1로 대응된다. 같은 주문이라도 배송/판매자가 갈려서 카드사에는 승인이
    각각 따로(금액도 다르게) 찍히는 경우가 있을 수 있는데, 주문번호로 묶으면
    이런 서로 다른 실거래를 하나로 뭉개버리는 문제가 생긴다.
    """
    by_approval = {}  # 승인번호(또는 없으면 페이지별 고유값) -> 항목 dict
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if "결제정보" not in text or "거래일시" not in text:
                continue  # 안내문만 있는 페이지는 건너뜀

            m_dt = _DT_RE.search(text)
            m_amt = _AMOUNT_RE.search(text)
            m_prod = _PRODUCT_RE.search(text)
            m_appr = _APPROVAL_RE.search(text)
            if not (m_dt and m_amt and m_prod):
                continue

            date_str, hh, mm = m_dt.group(1), int(m_dt.group(2)), int(m_dt.group(3))
            d = datetime.datetime.strptime(date_str, "%Y/%m/%d").date()
            amount = int(m_amt.group(1).replace(",", ""))
            product = re.sub(r"\s+", " ", m_prod.group(1)).strip()

            key = m_appr.group(1) if m_appr else object()  # 승인번호 없으면 단독 항목 취급
            if key in by_approval:
                by_approval[key]["products"].append(product)
                by_approval[key]["amount"] += amount  # 판매자별 결제 금액을 합산
            else:
                by_approval[key] = {"date": d, "hour": hh, "minute": mm, "amount": amount, "products": [product]}

    entries = []
    for entry in by_approval.values():
        product = " + ".join(dict.fromkeys(entry["products"]))  # 중복 제거하며 합침
        entries.append((entry["date"], entry["hour"], entry["minute"], entry["amount"], product))
    return entries
