#!/usr/bin/env python3
"""우리가족 VOCA 로컬 실행기 (run_app.py)
- 로컬 PC에서 웹 서버를 띄워 마이크 음성 인식(STT)과 발음 합성(TTS)이 정상 동작하도록 합니다.
- family_voca.xlsx 엑셀 파일을 읽어 /words.json 에 실시간 제공하므로,
  엑셀을 수정한 뒤 저장하고 웹 브라우저에서 새로고침(F5)하면 즉시 최신 단어가 반영됩니다.
- 종료: 콘솔에서 Ctrl + C
"""
import functools
import hashlib
import http.server
import json
import os
import re
import socketserver
import subprocess
import sys
import threading
import webbrowser

PORT = 8000
ROOT = os.path.dirname(os.path.abspath(__file__))
XLSX_PATH = os.path.join(ROOT, "family_voca.xlsx")
DIST_DIR = os.path.join(ROOT, "dist")

SHEETS = {"아빠": "dad", "엄마": "mom", "은봉이": "eun", "지봉이": "jib"}


def load_words_from_excel():
    """엑셀을 읽어 {version, profiles} 형태로 반환"""
    if not os.path.exists(XLSX_PATH):
        return None
    try:
        from openpyxl import load_workbook
    except ImportError:
        return None

    try:
        wb = load_workbook(XLSX_PATH, data_only=True, read_only=True)
    except Exception as e:
        print(f"! 엑셀 읽기 오류: {e}")
        return None

    profiles = {}
    for sheet_name, pid in SHEETS.items():
        if sheet_name not in wb.sheetnames:
            continue
        words = []
        for r in wb[sheet_name].iter_rows(min_row=4, values_only=True):
            if not r or all(v is None or str(v).strip() == "" for v in r):
                continue
            r = list(r) + [None] * 8
            w = str(r[1]).strip() if r[1] is not None else ""
            m = str(r[2]).strip() if r[2] is not None else ""
            p = str(r[3]).strip() if r[3] is not None else ""
            e = str(r[4]).strip() if r[4] is not None else ""
            c = str(r[5]).strip() if r[5] is not None else ""
            x = str(r[6]).strip() if r[6] is not None else ""
            lv = str(r[7]).strip() if r[7] is not None else "Lv1"
            if not (w and m and e):
                continue
            words.append({"w": w, "m": m, "p": p, "e": e, "c": c or "General", "x": x, "lv": lv})
        if words:
            profiles[pid] = words

    if not profiles:
        return None

    raw = json.dumps(profiles, sort_keys=True, ensure_ascii=False)
    ver = hashlib.sha256(raw.encode("utf-8")).hexdigest()[:10]
    return {"version": ver, "profiles": profiles}


class VocaHandler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        pass

    def end_headers(self):
        # 로컬 개발 시에는 캐시를 비활성화하여 항상 최신 코드/데이터 반영
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

    def do_GET(self):
        path_only = self.path.split("?")[0]
        if path_only == "/words.json":
            data = load_words_from_excel()
            if data is None:
                # 엑셀을 못 읽는 경우 dist/words.json 파일 폴백
                return super().do_GET()
            body = json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()


def main():
    # 1. build_app.py 실행하여 dist 생성
    build_script = os.path.join(ROOT, "build_app.py")
    if os.path.exists(build_script):
        print("🔨 build_app.py 실행 중...")
        res = subprocess.run([sys.executable, build_script])
        if res.returncode != 0:
            print("❌ 빌드에 실패했습니다. 위의 오류를 확인하세요.")
            sys.exit(1)

    serve_dir = DIST_DIR if os.path.exists(DIST_DIR) else ROOT
    handler = functools.partial(VocaHandler, directory=serve_dir)
    socketserver.TCPServer.allow_reuse_address = True

    port = PORT
    server = None
    while port < PORT + 50:
        try:
            server = socketserver.TCPServer(("127.0.0.1", port), handler)
            break
        except OSError:
            port += 1

    if not server:
        print("❌ 사용 가능한 포트를 찾지 못했습니다.")
        sys.exit(1)

    url = f"http://localhost:{port}/"
    print("=" * 60)
    print(f"🚀 우리가족 VOCA 로컬 서버가 시작되었습니다: {url}")
    print("   • 엑셀(family_voca.xlsx)을 수정 후 브라우저 새로고침(F5)하면 즉시 반영됩니다.")
    print("   • 서버 종료: Ctrl + C")
    print("=" * 60)

    threading.Timer(0.7, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n서버를 종료합니다.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
