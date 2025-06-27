import os
import sys
import pandas as pd
import unicodedata
from datetime import datetime
import traceback

def get_base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.abspath(os.path.join(os.path.dirname(sys.executable), "../../../"))
    else:
        return os.path.dirname(os.path.abspath(__file__))

def find_excel_file(base_dir, prefix, ext, log):
    for fname in os.listdir(base_dir):
        normalized = unicodedata.normalize("NFC", fname)
        log.append(f"▶ 검사 중: {normalized}")
        if normalized.startswith(prefix) and normalized.endswith(ext):
            log.append("✅ 조건 만족!")
            return os.path.join(base_dir, fname)
    return None

def main():
    log = []
    error_log = ""
    try:
        base_dir = get_base_dir()
        log.append(f"현재 base_dir: {base_dir}")
        today_prefix = datetime.now().strftime("%Y%m%d")

        # 1. 이지어드민 파일
        file_to_read1 = find_excel_file(base_dir, f"판매처별송장등록_()_{today_prefix}", ".xls", log)
        if not file_to_read1:
            raise FileNotFoundError(f"❌ '{today_prefix}'일자 기준 송장 기준 엑셀(판매처별송장등록)을 현재 폴더에서 찾을 수 없습니다.")
        log.append(f"✅ 파일 읽기 완료: {os.path.basename(file_to_read1)}")
        easy_songjang_df = pd.read_excel(file_to_read1)

        # 2. 플레이오토 파일
        file_to_read2 = find_excel_file(base_dir, f"토글형식_{today_prefix}", ".xlsx", log)
        if not file_to_read2:
            raise FileNotFoundError(f"❌ '{today_prefix}'일자 기준 플레이오토 파일(토글형식)을 현재 폴더에서 찾을 수 없습니다.")
        log.append(f"✅ 파일 읽기 완료: {os.path.basename(file_to_read2)}")
        playauto_df = pd.read_excel(file_to_read2)

        log.append("😎 송장 파일로 변환 중입니다...")

        # 3. 매칭 처리
        easy_songjang_df.columns = easy_songjang_df.columns.str.strip()
        playauto_df.columns = playauto_df.columns.str.strip()

        merge_keys_a = ["주문자", "수령자", "수령자전화", "수령자핸드폰"]
        merge_keys_b = ["주문자명", "수령자명", "수령자휴대폰번호", "수령자전화번호"]

        df_invoice_map = easy_songjang_df[merge_keys_a + ["송장번호"]]

        def find_invoice(row):
            condition = pd.Series([True] * len(df_invoice_map))
            for col_a, col_b in zip(merge_keys_a, merge_keys_b):
                val = row.get(col_b)
                if pd.notna(val) and val != "":
                    condition &= (df_invoice_map[col_a] == val)
            match = df_invoice_map[condition]
            if not match.empty:
                return str(match.iloc[0]["송장번호"])
            return str(row.get("운송장번호", ""))

        playauto_df["운송장번호"] = playauto_df.apply(find_invoice, axis=1)

        # 4. 저장 (텍스트 포맷으로 송장번호 저장)
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"송장_{timestamp}.xlsx"
        save_path = os.path.join(base_dir, filename)

        with pd.ExcelWriter(save_path, engine='xlsxwriter') as writer:
            playauto_df.to_excel(writer, index=False, sheet_name="토글_송장출력")
            workbook = writer.book
            worksheet = writer.sheets["토글_송장출력"]

            # 송장번호 열을 텍스트 형식으로 지정 (포맷: '@')
            text_format = workbook.add_format({'num_format': '@'})
            송장_col_idx = playauto_df.columns.get_loc("운송장번호")
            worksheet.set_column(송장_col_idx, 송장_col_idx, 20, text_format)

        log.append(f"✅ 송장 파일 저장 완료: {filename}")

    except Exception as e:
        error_log += f"\n❌ 오류 발생: {str(e)}\n"
        error_log += traceback.format_exc()

    # 5. 로그 저장 (홈 디렉토리)
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    home_dir = os.path.expanduser("~")

    # 일반 로그
    log_path = os.path.join(home_dir, f"송장_{timestamp}_log.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write("\n".join(log))

    # 에러 로그
    if error_log:
        error_log_path = os.path.join(home_dir, "invoice_app_error_log.txt")
        with open(error_log_path, "a", encoding="utf-8") as f:
            f.write(f"\n[{timestamp}] 에러 발생 로그\n")
            f.write(error_log)

    # macOS 자동 로그 열기
    if sys.platform == "darwin":
        os.system(f"open '{log_path}'")

if __name__ == "__main__":
    main()
