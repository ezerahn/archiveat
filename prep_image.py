# -*- coding: utf-8 -*-
"""
메뉴 사진 정규화 — 자르지 않고 정방형으로 맞추고 참고용 크기로 줄인다.
결과물은 gen_menu.py가 HTML 안에 base64로 박아 넣으므로,
서버(github)에는 이미지 파일을 올리지 않는다. menu-images/는 원본 보관용.

사용법: python prep_image.py <원본파일> <메뉴ID> [순번]
  예:   python prep_image.py photo.png HG-021        → HG-021.webp
        python prep_image.py photo2.png HG-021 2     → HG-021-2.webp

규격: 280x280 · WebP 품질72 · 여백색 #FFFDF9(카드 배경과 동일)
      화면 표시 112px 기준 2.5배 해상도. 100장 내장 시 HTML 약 1.7MB.
"""
import sys, os
from PIL import Image

SIDE, QUALITY = 280, 72
BG = (255, 253, 249)
OUTDIR = "menu-images"

def prep(src, name):
    im = Image.open(src).convert("RGB")
    w, h = im.size
    s = max(w, h)                                  # 잘라내지 않음
    canvas = Image.new("RGB", (s, s), BG)
    canvas.paste(im, ((s - w) // 2, (s - h) // 2))
    canvas = canvas.resize((SIDE, SIDE), Image.LANCZOS)
    os.makedirs(OUTDIR, exist_ok=True)
    path = os.path.join(OUTDIR, name)
    canvas.save(path, "WEBP", quality=QUALITY, method=6)
    print(f"{name}  ←  {w}x{h}  →  {SIDE}x{SIDE} · {os.path.getsize(path)/1024:.0f}KB (자름 없음)")
    return path

if __name__ == "__main__":
    src, mid = sys.argv[1], sys.argv[2]
    seq = sys.argv[3] if len(sys.argv) > 3 else ""
    prep(src, f"{mid}-{seq}.webp" if seq else f"{mid}.webp")
