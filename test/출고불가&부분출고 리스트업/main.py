import concurrent.futures
import requests as req
from bs4 import BeautifulSoup as bs
import pandas as pd
import re
from datetime import datetime, timedelta
import os

def process_orders():
    # 오늘 날짜 구하기
    end_date = datetime.now()
    # 30일 전 날짜 계산
    start_date = end_date - timedelta(days=30)

    # 날짜를 YYYY-MM-DD 형식의 문자열로 변환
    end_date_str = end_date.strftime("%Y-%m-%d")
    start_date_str = start_date.strftime("%Y-%m-%d")

    # URL 생성
    url1 = (
        f"{os.getenv('BASE_URL')}/orders/order_list.asp"
        "?first_in_check=no&detail_search=no&sub_manager=&show_order_store_opt_chk=off"
        "&service_category_idx=0&invoice_category_flag=no&invoice_category_idx=0"
        "&special_area_site_code=0&product_tags=&shop_category_idx=all&date_option=order"
        f"&date_type=custom&start_date={start_date_str}&end_date={end_date_str}"
        "&search_how=all&search_key=&is_collapsed=false&search_how_add=all&search_key_add="
        "&stat_type=custom&matched=dontcare&invoiced=off&delivered=off&cs_type=nocs"
        "&sum_scale=all&amount_scale=all&inSchDay=all&soldDate=all&pdCat_flag=all"
        "&optClass_flag=all&sales_flag=all&sold_out_flag=all&output_psb_flag=all"
        "&post_category=all&delivery_method=all&brand=all&product_year=all&product_season=all"
        "&bySM_gubun=all&hold_flag=noHold&able_stk_out_flag=divide_out&opt_match_flag=all"
        "&hapo_order=all&uId=all&apartFlag=no&separate_category_idx=all"
        "&delivery_location_flag=0&dangol_gubun=all&sep_no=all&ord_list_page_count=500"
    )
    url2 = (
        f"{os.getenv('BASE_URL')}/orders/order_list.asp"
        "?first_in_check=no&detail_search=no&sub_manager=&show_order_store_opt_chk=off"
        "&service_category_idx=0&invoice_category_flag=no&invoice_category_idx=0"
        "&special_area_site_code=0&product_tags=&shop_category_idx=all&date_option=order"
        f"&date_type=custom&start_date={start_date_str}&end_date={end_date_str}"
        "&search_how=all&search_key=&is_collapsed=false&search_how_add=all&search_key_add="
        "&stat_type=custom&matched=dontcare&invoiced=off&delivered=off&cs_type=nocs"
        "&sum_scale=all&amount_scale=all&inSchDay=all&soldDate=all&pdCat_flag=all"
        "&optClass_flag=all&sales_flag=all&sold_out_flag=all&output_psb_flag=all"
        "&post_category=all&delivery_method=all&brand=all&product_year=all&product_season=all"
        "&bySM_gubun=all&hold_flag=noHold&able_stk_out_flag=no_out&opt_match_flag=all"
        "&hapo_order=all&uId=all&apartFlag=no&separate_category_idx=all"
        "&delivery_location_flag=0&dangol_gubun=all&sep_no=all&ord_list_page_count=500"
    )

    # 헤더 설정
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,"
                  "image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.7",
        "Accept-Encoding": "gzip, deflate, br, zstd",
        "Accept-Language": "ko,en-US;q=0.9,en;q=0.8",
        "Cache-Control": "max-age=0",
        "Connection": "keep-alive",
        "Host": "lcnine.sellmate.co.kr",
        "Referer": (
            f"https://lcnine.sellmate.co.kr/orders/order_list.asp"
            "?first_in_check=no&gotopage=1&search_how=all&search_key=&date_type=custom"
            f"&start_date={start_date_str}&end_date={end_date_str}&shop_code=all&stat_type=c"
        ),
    }

    # 쿠키 설정
    cookies = {
        "domain": "lcnine",
        "id_save": "OK",
        "smi": "pVEXIdA76H%2FPQj3yL%2BtwsI13HFqlbtsuTIU8c2asQbk%3D",
        "show_lack_chk": "on",
        "show_lack_chk2": "on",
        "eventLayer3": "done",
        "show_lack_chk3": "on",
        "SELLMATESESSIONID": "5EF1C80C12954C75AF5B4B84C79AF3C6",
        "useLeftMenu": "Y",
        "cslog_init_check": f"{os.getenv('SELLMATE_ID')}=True",
        "SMSS.CCC173549B54904C34C4ED9BB6F20A71": str(os.getenv("SMSS")),
    }

    # 주소 요청용 헤더
    address_headers = {
        "Content-Type": "application/x-www-form-urlencoded",
        **headers
    }

    # POST 요청용 데이터 (주소 조회할 때 사용하는 폼 데이터 기본값)
    x_www_data = {
        "mode": "list",
        "request_from": "",
        "orderListDiv_height": "",
        "orderDetailDiv_height": "",
        "excel_make_type": "4",
        "gotopage": "",
        "side_idx": "",
        "side_csidx": "",
        "side_selectidx": "",
        "newcs_page_count": "20",
        "search_type": "",
        "gf_use_flag": "1",
        "check_epost": "0",
        "cjLogistics": "0",
        "logen": "1",
        "hanjin": "0",
        "hanjin_version": "1",
        "quickfinder_use": "0",
        "returneeds_use_flag": "0",
        "pantos_use_flag": "0",
        "get_csuser_filter_flag": "0",
        "barcodeNo2_check": "1",
        "barcodeNo3_check": "1",
        "search_page_data": "",
        "search_page_data_switch": "0",
        "first_in_check": "no",
        "shpandout_version": "1",
        "date_option": "order",
        "date_type": "week",
        "search_how": "order_code",
        "hidden_search_text": "",
        "search_text": "",
        "stat_type": "all",
        "matched": "off",
        "optmatched": "dontcare",
        "invoiced": "off",
        "delivered": "off",
        "cs_kind": "",
        "cs_type": "all",
        "cs_status": "",
        "cs_type2": "0",
        "cs_channel": "0",
        "ord_status": "all",
        "dangol_gubun": "all",
        "output_psb_flag": "all",
        "sold_out_flag": "all",
        "dontcare_soldout_idx": ["0", "1", "2", "3", "4"],
        "sep_no": "all",
        "able_stk_out_flag": "all",
        "first_stk_out_flag": "all",
        "apartFlag": "all",
        "separate_category_idx": "all",
        "hold_flag": "all",
        "chasu": "all",
        "hapo_order": "all",
        "shop_category_idx": "all",
        "sales_flag": "all",
        "prevent_remerge_flag": "all",
        "dealer_risk_flag": "all",
        "invoice_risk_flag": "all",
        "post_category": "all",
        "delivery_location_flag": "all",
        "optClass_flag": "all",
        "preventDivide_flag": "all",
        "brand": "all",
        "product_year": "all",
        "product_season": "all",
        "sum_scale": "all",
        "MinCost": "0",
        "MaxCost": "0",
        "amount_scale": "all",
        "minAmount": "0",
        "maxAmount": "0",
        "inSchDay": "all",
        "back_cost_type": "all",
        "soldDate": "all",
        "back_scale": "all",
        "search_how_add": "all",
        "search_key_add": "",
        "packagingVideoColumnExists": "true",
    }

    def fetch_order_page(url: str) -> pd.DataFrame:
        """
        주어진 URL에서 주문 테이블을 파싱해 DataFrame으로 반환한다.
        """
        response = req.get(url=url, headers=headers, cookies=cookies)
        response.encoding = "utf-8"
        soup = bs(response.text, "html.parser")

        data = []
        order_table = soup.select_one("table.second_table")
        if order_table:
            order_table_rows = order_table.select("tbody tr")
            for tr in order_table_rows:
                tds = tr.select("td")
                # 10개 미만이면 제대로 된 행이 아니라는 의미
                if len(tds) < 10:
                    continue

                # 판매처
                sale_place = tds[2].find("a").get_text(strip=True)

                # 주문번호, 연락처
                order_td = tds[3]
                order_number = order_td.select_one("a").get_text(strip=True)
                contact = next(
                    (
                        link.get_text(strip=True)
                        for link in order_td.select("a[onclick*='checkSMSsend_power']")
                    ),
                    "",
                )

                # 상품정보 (상품이름, 옵션, 주문수량, 출고불가능여부)
                product_td = tds[4]
                product_name = (
                    product_td.select_one("strong").get_text(strip=True)
                    if product_td.select_one("strong")
                    else ""
                )
                option = product_td.select_one(".product_name").get_text(strip=True)

                quantity_font = product_td.find("font", string=lambda t: t and "주문수량" in t)
                order_quantity = 0
                if quantity_font:
                    match = re.search(r"x(\d+)", quantity_font.get_text(strip=True))
                    order_quantity = int(match.group(1)) if match else 0

                status_tag = product_td.select_one("span.label-danger")
                is_unavailable = (
                    "O" if (status_tag and status_tag.get_text(strip=True) == "출고불가능") else "X"
                )

                # 주문자명
                name_td = tds[9]
                orderer_name = (
                    name_td.select_one('font[color="#333333"]').get_text(strip=True)
                    if name_td.select_one('font[color="#333333"]')
                    else ""
                )

                data.append(
                    {
                        "주문자명": orderer_name,
                        "연락처": contact,
                        "주소": None,  # 주소는 나중에 별도 요청으로 채움
                        "주문번호": order_number,
                        "상품이름": product_name,
                        "옵션": option,
                        "수량": order_quantity,
                        "출고불가능": is_unavailable,
                        "판매처": sale_place,
                    }
                )

        return pd.DataFrame(
            data,
            columns=[
                "주문번호",
                "주문자명",
                "출고불가능",
                "연락처",
                "상품이름",
                "옵션",
                "수량",
                "주소",
                "판매처",
            ],
        )

    def fetch_address(order_number: str) -> str:
        """
        주문번호( order_number )를 이용해 주소를 조회한다.
        """
        url = "https://lcnine.sellmate.co.kr/cs/new_cs_search.asp"
        form_data = x_www_data.copy()
        form_data["hidden_search_text"] = order_number
        form_data["search_text"] = order_number

        response = req.post(url, data=form_data, headers=address_headers, cookies=cookies)
        response.encoding = "utf-8"
        soup = bs(response.text, "html.parser")

        cs_tr = soup.select_one("#mainTableTr1")
        if cs_tr:
            address_td = cs_tr.select_one(".address-col")
            if address_td:
                return (
                    address_td.get_text()
                    .split("\n")[0]
                    .strip()
                    .replace("[배송지수정]", "")
                    .strip()
                )
        return None

    # ------------------- 주문 목록 병렬 요청 -------------------
    with concurrent.futures.ThreadPoolExecutor(max_workers=2) as executor:
        results = list(executor.map(fetch_order_page, [url1, url2]))

    # results[0] -> 부분출고 대상, results[1] -> 출고불가능 대상(이라고 가정)
    # 그러나 실제 로직에서는 같은 구조의 테이블을 두 번 받아온 뒤
    # '출고불가능' 컬럼을 보고 O/X로 구분한다.

    final_dfs = []
    for df in results:
        # 동일 연락처 묶음에 대해 주문자명만 가져오는 로직: 첫 행에 있는 이름으로 통일
        df["주문자명"] = df.groupby("연락처")["주문자명"].transform(lambda x: x.iloc[0])

        # 출고불가능(O)만 선별
        unavailable_orders = df[df["출고불가능"] == "O"].copy()
        unavailable_orders = unavailable_orders.drop(columns=["출고불가능"])

        # ------------------- 주소 정보 채우기 (중복 최소화) -------------------
        unique_order_numbers = unavailable_orders["주문번호"].unique()
        with concurrent.futures.ThreadPoolExecutor(max_workers=5) as executor:
            address_list = list(executor.map(fetch_address, unique_order_numbers))
        # 주문번호:주소 매핑 딕셔너리
        address_dict = dict(zip(unique_order_numbers, address_list))

        # unavailable_orders에 map으로 주소 할당
        unavailable_orders["주소"] = unavailable_orders["주문번호"].map(address_dict)

        # cs링크 컬럼 추가
        unavailable_orders["cs링크"] = unavailable_orders["주문번호"].apply(
            lambda x: f'=HYPERLINK("{os.getenv("BASE_URL")}/cs/new_cs_frame.asp'
                      f'?search_how=order_code&search_text={x.strip()}&first_in_check=no", "이동하기")'
        )

        final_dfs.append(unavailable_orders)

    # ------------------- Excel로 저장 -------------------
    timestamp = end_date.strftime("%Y-%m-%d_%H-%M-%S")
    excel_filename = f"부분출고&출고불가{timestamp}.xlsx"
    with pd.ExcelWriter(excel_filename) as writer:
        # final_dfs[0] -> 첫 번째 URL의 '출고불가능' 데이터 = "부분출고" 시트
        final_dfs[0].to_excel(writer, sheet_name="부분출고", index=False)
        # final_dfs[1] -> 두 번째 URL의 '출고불가능' 데이터 = "출고불가" 시트
        final_dfs[1].to_excel(writer, sheet_name="출고불가", index=False)

    print(f"Excel 파일이 저장되었습니다: {excel_filename}")


if __name__ == "__main__":
    process_orders()
