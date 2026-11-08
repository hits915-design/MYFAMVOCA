#!/usr/bin/env python3
"""우리가족 VOCA 빌더 (build_app.py)
family_voca.xlsx 엑셀 파일을 읽어 데이터 유효성을 검증하고,
PWA 배포용 파일(dist/words.json, dist/index.html, dist/sw.js 등)을 생성합니다.
"""
import hashlib
import json
import os
import re
import shutil
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(ROOT, "family_voca.xlsx")
SRC_DIR = os.path.join(ROOT, "src")
DIST_DIR = os.path.join(ROOT, "dist")

SHEETS = {
    "아빠": "dad",
    "엄마": "mom",
    "은봉이": "eun",
    "지봉이": "jib",
}

VALID_LEVELS = {"Lv1", "Lv2", "Lv3", "Lv4", "Lv5"}


def validate_and_extract_words():
    if not os.path.exists(XLSX_PATH):
        print(f"❌ [오류] 엑셀 파일을 찾을 수 없습니다: {XLSX_PATH}")
        sys.exit(1)

    try:
        from openpyxl import load_workbook
    except ImportError:
        print("❌ [오류] openpyxl 라이브러리가 필요합니다: pip install openpyxl")
        sys.exit(1)

    try:
        wb = load_workbook(XLSX_PATH, data_only=True, read_only=True)
    except Exception as e:
        print(f"❌ [오류] 엑셀 파일을 열 수 없습니다: {e}")
        sys.exit(1)

    errors = []
    profiles_data = {}

    for sheet_name, pid in SHEETS.items():
        if sheet_name not in wb.sheetnames:
            errors.append(f"시트 누락: '{sheet_name}' 시트가 엑셀에 없습니다. (필수 시트: {', '.join(SHEETS.keys())})")
            continue

        ws = wb[sheet_name]
        words = []
        seen_words = {}  # {lower_w: row_num}

        row_num = 3  # min_row=4 시작 전 헤더 위치
        for row in ws.iter_rows(min_row=4, values_only=True):
            row_num += 1
            # 행이 비어있는지 확인
            if not row or all(v is None or str(v).strip() == "" for v in row):
                continue

            r = list(row) + [None] * 8
            # 열 구조: 0:번호, 1:단어, 2:뜻, 3:발음, 4:예문, 5:카테고리, 6:그림, 7:난이도
            w = str(r[1]).strip() if r[1] is not None else ""
            m = str(r[2]).strip() if r[2] is not None else ""
            p = str(r[3]).strip() if r[3] is not None else ""
            e = str(r[4]).strip() if r[4] is not None else ""
            c = str(r[5]).strip() if r[5] is not None else ""
            x = str(r[6]).strip() if r[6] is not None else ""
            lv = str(r[7]).strip() if r[7] is not None else ""

            row_has_error = False

            # 1. 단어/뜻/예문 필수 항목 누락 검사
            missing = []
            if not w:
                missing.append("단어")
            if not m:
                missing.append("뜻")
            if not e:
                missing.append("예문")
            if missing:
                errors.append(f"[{sheet_name}] {row_num}행: 필수 항목 누락 ({', '.join(missing)})")
                row_has_error = True

            # 단어가 없으면 다음 세부 검사는 건너뜀
            if not w:
                continue

            # 2. 단어 칸에 작은따옴표(') 검사
            if any(q in w for q in ("'", "’", "‘", "`")):
                errors.append(
                    f"[{sheet_name}] {row_num}행: 단어 '{w}'에 작은따옴표(')가 포함되어 있습니다. 작은따옴표를 제거해 주세요."
                )
                row_has_error = True

            # 3. 같은 시트 안 단어 중복 검사
            w_lower = w.lower()
            if w_lower in seen_words:
                prev_row = seen_words[w_lower]
                errors.append(
                    f"[{sheet_name}] {row_num}행: 단어 '{w}'가 시트 내에서 중복되었습니다. (이전 {prev_row}행에 이미 존재)"
                )
                row_has_error = True
            else:
                seen_words[w_lower] = row_num

            # 4. 예문 안에 단어 포함 여부 검사 (빈칸 퀴즈용)
            if e and not re.search(re.escape(w), e, re.IGNORECASE):
                errors.append(
                    f"[{sheet_name}] {row_num}행: 단어 '{w}'가 예문 '{e}'에 포함되어 있지 않습니다. (빈칸 퀴즈 출제를 위해 예문에 단어가 포함되어야 합니다)"
                )
                row_has_error = True

            # 5. 난이도 Lv1~Lv5 검사
            # Lv 표기 정규화 (대소문자 지원: lv1 -> Lv1)
            norm_lv = "Lv" + lv[2:] if lv.lower().startswith("lv") else lv
            if norm_lv not in VALID_LEVELS:
                errors.append(
                    f"[{sheet_name}] {row_num}행: 난이도 '{lv}'가 올바르지 않습니다. (Lv1, Lv2, Lv3, Lv4, Lv5 중 하나여야 합니다)"
                )
                row_has_error = True

            if not row_has_error:
                words.append({
                    "w": w,
                    "m": m,
                    "p": p,
                    "e": e,
                    "c": c or "General",
                    "x": x,
                    "lv": norm_lv
                })

        if not words and sheet_name not in errors:
            errors.append(f"[{sheet_name}] 단어가 하나도 없습니다. 최소 1개 이상의 단어를 입력해 주세요.")

        profiles_data[pid] = words

    if errors:
        print("=" * 60)
        print("❌ [엑셀 검증 실패] 다음 오류를 수정한 후 다시 시도해 주세요:")
        print("=" * 60)
        for err in errors:
            print(f" • {err}")
        print("=" * 60)
        sys.exit(1)

    return profiles_data


def build():
    print("📦 [1/4] 엑셀 데이터 검증 중...")
    profiles = validate_and_extract_words()
    
    # 통계 출력
    for sheet_name, pid in SHEETS.items():
        count = len(profiles.get(pid, []))
        print(f"   ✓ {sheet_name}: {count}개 단어 확인")

    # words.json 내용 구성 및 해시 생성
    content_raw = json.dumps(profiles, sort_keys=True, ensure_ascii=False)
    version_hash = hashlib.sha256(content_raw.encode("utf-8")).hexdigest()[:10]
    words_payload = {
        "version": version_hash,
        "profiles": profiles
    }

    print(f"📦 [2/4] 빌드 버전 생성: {version_hash}")

    # dist 디렉터리 준비
    os.makedirs(DIST_DIR, exist_ok=True)

    # dist/words.json 저장
    words_json_path = os.path.join(DIST_DIR, "words.json")
    with open(words_json_path, "w", encoding="utf-8") as f:
        json.dump(words_payload, f, ensure_ascii=False, indent=2)
    print("   ✓ dist/words.json 생성 완료")

    # dist/index.html 복사 (src/app.html -> dist/index.html)
    app_html_src = os.path.join(SRC_DIR, "app.html")
    if not os.path.exists(app_html_src):
        print(f"❌ [오류] {app_html_src} 파일이 없습니다.")
        sys.exit(1)
    
    shutil.copy2(app_html_src, os.path.join(DIST_DIR, "index.html"))
    print("   ✓ dist/index.html 생성 완료 (src/app.html 복사)")

    # dist/manifest.json 복사
    manifest_src = os.path.join(SRC_DIR, "manifest.json")
    if os.path.exists(manifest_src):
        shutil.copy2(manifest_src, os.path.join(DIST_DIR, "manifest.json"))
        print("   ✓ dist/manifest.json 복사 완료")

    # 아이콘 복사 (src/assets/icon-*.png -> dist/icon-*.png)
    assets_dir = os.path.join(SRC_DIR, "assets")
    if os.path.exists(assets_dir):
        for item in os.listdir(assets_dir):
            if item.endswith(".png"):
                shutil.copy2(os.path.join(assets_dir, item), os.path.join(DIST_DIR, item))
        print("   ✓ dist/ 아이콘 파일 복사 완료")

    # dist/sw.js 생성 (캐시 버전 치환)
    sw_src = os.path.join(SRC_DIR, "sw.js")
    if os.path.exists(sw_src):
        with open(sw_src, "r", encoding="utf-8") as f:
            sw_code = f.read()
        sw_code = sw_code.replace("{{CACHE_VERSION}}", version_hash)
        with open(os.path.join(DIST_DIR, "sw.js"), "w", encoding="utf-8") as f:
            f.write(sw_code)
        print(f"   ✓ dist/sw.js 생성 완료 (캐시 태그: family-voca-{version_hash})")

    print("=" * 60)
    print(f"🎉 [성공] PWA 빌드가 성공적으로 완료되었습니다! (버전: {version_hash})")
    print("=" * 60)


if __name__ == "__main__":
    build()
