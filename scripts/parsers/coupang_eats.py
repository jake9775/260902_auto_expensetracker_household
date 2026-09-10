# -*- coding: utf-8 -*-
"""
쿠팡이츠 결제내역 엑셀 파일 파서.

쿠팡(주문내역 PDF, coupang.py)과는 별개 파일이다. 우리은행/현대카드/토스뱅크
내역에 "쿠팡이츠" 등으로 찍히는 결제 건의 실제 메뉴명을 알려주는 "보조 자료"로만
쓴다. 그래서 거래 행을 만들지 않고, (날짜, 시각, 금액) -> 상품명 조회용 목록만
만든다.

시트: 결제내역
헤더: 결제일시, PG사, 상점명, 상품명, 결제금액, 구분(결제/취소), 주문번호, 승인번호
"""
import datetime


def extract_entries(path):
    """[(date, hour, minute, amount, product_name), ...] 목록을 돌려준다.

    '구분'이 '취소'인 행은 실제로 청구되지 않은(또는 환불된) 건이라 카드
    내역과 맞춰볼 필요가 없으므로 건너뛴다.

    쿠팡/네이버페이와 마찬가지로, 한 결제(승인번호)에 여러 메뉴가 섞여 있으면
    같은 승인번호로 여러 줄에 나눠서 찍힐 수 있다. 그대로 두면 (날짜, 금액)
    키가 겹쳐서 상품명 하나만 남고 나머지가 유실되므로, 같은 승인번호끼리는
    상품명을 합치고 금액은 더해서 하나의 항목으로 만든다.
    """
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["결제내역"] if "결제내역" in wb.sheetnames else wb[wb.sheetnames[0]]

    by_approval = {}  # 승인번호(또는 없으면 페이지별 고유값) -> 항목 dict
    for raw in ws.iter_rows(min_row=2, values_only=True):
        if raw is None or len(raw) < 8:
            continue
        pay_dt, _pg, _shop, product, amount, kind, _order_no, approval_no = raw[:8]
        if not pay_dt or amount in (None, ""):
            continue
        if (kind or "").strip() != "결제":
            continue
        dt = _parse_datetime(pay_dt)
        if dt is None:
            continue
        product = str(product or "").strip()

        key = approval_no if approval_no is not None else object()
        if key in by_approval:
            by_approval[key]["products"].append(product)
            by_approval[key]["amount"] += int(amount)
        else:
            by_approval[key] = {
                "date": dt.date(), "hour": dt.hour, "minute": dt.minute,
                "amount": int(amount), "products": [product],
            }

    entries = []
    for entry in by_approval.values():
        product = " + ".join(dict.fromkeys(entry["products"]))  # 중복 제거하며 합침
        entries.append((entry["date"], entry["hour"], entry["minute"], entry["amount"], product))
    return entries


def _parse_datetime(value):
    if isinstance(value, datetime.datetime):
        return value
    if isinstance(value, str):
        try:
            return datetime.datetime.strptime(value.strip(), "%Y-%m-%d %H:%M:%S")
        except ValueError:
            return None
    return None
