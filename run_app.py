"""우리가족 VOCA 실행기 (엑셀 연동)
- family-voca.html 을 localhost 로 열어 마이크 권한이 정상 동작하게 합니다.
- 같은 폴더의 family_voca.xlsx 에서 단어를 읽어 앱에 자동 반영합니다.
  (시트 이름: 아빠 / 엄마 / 은봉이 / 지봉이, 4행부터, 열: 번호·단어·뜻·발음·예문·카테고리·그림)
- 엑셀을 고쳐 저장한 뒤 브라우저를 새로고침(F5)하면 바로 적용됩니다.
- 필요: Python 3.8+, 엑셀 연동 시 `pip install openpyxl` (없어도 기본 단어로 실행)
종료: Ctrl+C
"""
import functools
import http.server
import json
import os
import re
import socketserver
import threading
import webbrowser

HTML = "family-voca.html"
XLSX = "family_voca.xlsx"
PORT = 8000
ROOT = os.path.dirname(os.path.abspath(__file__))
SHEETS = {"아빠": "dad", "엄마": "mom", "은봉이": "eun", "지봉이": "jib"}


def load_words():
    """엑셀 → {프로필id: [[단어, 뜻, 발음, 예문, 카테고리, 그림], ...]} (매 요청마다 다시 읽음)"""
    path = os.path.join(ROOT, XLSX)
    if not os.path.exists(path):
        return {}
    try:
        from openpyxl import load_workbook
    except ImportError:
        print("! openpyxl 이 없어 엑셀을 읽지 못했습니다:  pip install openpyxl")
        return {}
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
    except Exception as e:  # 엑셀이 열려 있거나 손상된 경우
        print(f"! 엑셀을 읽지 못했습니다({e}). 기본 단어로 실행합니다.")
        return {}
    out = {}
    for sheet, pid in SHEETS.items():
        if sheet not in wb.sheetnames:
            continue
        words = []
        for r in wb[sheet].iter_rows(min_row=4, values_only=True):
            r = list(r) + [None] * 7
            w, m, p, e, c, x = (str(v).strip() if v is not None else "" for v in r[1:7])
            if not (w and m and e):
                continue
            if not re.search(re.escape(w), e, re.I):
                print(f"! [{sheet}] '{w}' 예문에 단어가 들어 있지 않아 빈칸 퀴즈가 어색할 수 있어요.")
            words.append({"w": w, "m": m, "p": p, "e": e, "c": c or "General", "x": x})
        if words:
            out[pid] = words
    return out


class Handler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, *args):
        pass

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def do_GET(self):
        if self.path.split("?")[0] == "/words.json":
            body = json.dumps(load_words(), ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        super().do_GET()


def main():
    if not os.path.exists(os.path.join(ROOT, HTML)):
        raise SystemExit(f"{HTML} 파일이 {ROOT} 에 없습니다. 같은 폴더에 두세요.")

    handler = functools.partial(Handler, directory=ROOT)
    socketserver.TCPServer.allow_reuse_address = True
    port = PORT
    while True:
        try:
            server = socketserver.TCPServer(("127.0.0.1", port), handler)
            break
        except OSError:
            port += 1

    url = f"http://localhost:{port}/{HTML}"
    found = load_words()
    print(f"실행 중: {url}")
    print("엑셀 연동:", ", ".join(f"{k}={len(v)}단어" for k, v in found.items()) if found else "사용 안 함(기본 단어)")
    print("종료하려면 Ctrl+C")
    threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n종료합니다.")
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
