# -*- coding: utf-8 -*-
"""
우리은행 '거래내역조회' 엑셀 파일 파서.

기대하는 형식 (인터넷뱅킹에서 내려받은 그대로):
  1~3행: 제목/계좌번호/조회기간 안내문
  4행  : 헤더 (No. 거래일시 적요 기재내용 찾으신금액 맡기신금액 거래후잔액 취급기관 메모)
  5행~ : 실제 거래 데이터
"""
import datetime
import openpyxl


def parse(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[wb.sheetnames[0]]

    rows = []
    for raw in ws.iter_rows(min_row=5, values_only=True):
        if raw is None or len(raw) < 6:
            continue
        no = raw[0]
        if not isinstance(no, (int, float)):
            continue  # 데이터가 아닌 행(빈 줄, 안내문 등)은 건너뜀

        dt_str = raw[1]
        channel = (raw[2] or "").strip()      # 적요: 체크하나/모바일/F-B출금/우리카드 등 거래 채널
        content = (raw[3] or "").strip()      # 기재내용: 실제 거래처/설명
        out_amt = raw[4] or 0
        in_amt = raw[5] or 0

        if not dt_str:
            continue
        try:
            dt = datetime.datetime.strptime(str(dt_str).strip(), "%Y.%m.%d %H:%M")
        except ValueError:
            continue

        desc = content or channel

        # 채널 자체로 성격이 명확한 거래는 카테고리를 미리 정해둔다.
        # (같은 문구가 입금/환급에도 쓰일 수 있어 키워드 규칙만으로는 오분류될 수 있음)
        preset_main = preset_sub = None
        if channel == "F/B 출금":
            preset_main, preset_sub = "카드대금", content or "카드대금"
        elif channel == "우리카드":
            preset_main, preset_sub = "카드대금", "우리카드"
        elif channel == "수수료":
            preset_main, preset_sub = "수수료", "은행수수료"
        elif channel == "대체취소":
            preset_main, preset_sub = "환불", "결제취소"

        base = {
            "은행": "우리은행",
            "거래일자": dt.date(),
            "거래시간": dt.time(),
            "적요": desc,
            "내용": desc,
        }
        if preset_main:
            base["대분류"] = preset_main
            base["소분류"] = preset_sub

        if out_amt:
            row = dict(base)
            row["출금"], row["입금"] = out_amt, 0
            rows.append(row)
        if in_amt:
            row = dict(base)
            row["출금"], row["입금"] = 0, in_amt
            rows.append(row)

    return rows
