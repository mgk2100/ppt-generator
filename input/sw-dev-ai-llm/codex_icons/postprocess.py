"""Codex imagegen 결과 후처리: 체커보드/흰 배경 → 투명, 여백 트리밍 후 정사각 720px"""
import glob, os, sys
from PIL import Image
import numpy as np
S = os.path.dirname(os.path.abspath(__file__))
src_dir, dst_dir = f"{S}/out", f"{S}/png"
os.makedirs(dst_dir, exist_ok=True)
for f in sorted(glob.glob(f"{src_dir}/*.png")):
    im = Image.open(f).convert("RGBA")
    a = np.array(im).astype(np.int16)
    r, g, b, al = a[..., 0], a[..., 1], a[..., 2], a[..., 3]
    mx = np.maximum(np.maximum(r, g), b); mn = np.minimum(np.minimum(r, g), b)
    sat = mx - mn
    # 채도 낮고(회색·흰색) 밝은 픽셀 = 배경(체커보드 포함) → 투명
    bg = (sat < 18) & (mn > 170)
    # 기존 알파가 이미 0 인 곳도 배경
    bg |= (al < 10)
    # 가장자리 연결 성분만 배경으로 삼기(아이콘 내부 흰 면 보존): flood fill from border
    from collections import deque
    h, w = bg.shape
    seen = np.zeros_like(bg, dtype=bool)
    dq = deque()
    for x in range(w):
        for y in (0, h - 1):
            if bg[y, x] and not seen[y, x]: seen[y, x] = True; dq.append((y, x))
    for y in range(h):
        for x in (0, w - 1):
            if bg[y, x] and not seen[y, x]: seen[y, x] = True; dq.append((y, x))
    while dq:
        y, x = dq.popleft()
        for ny, nx in ((y-1,x),(y+1,x),(y,x-1),(y,x+1)):
            if 0 <= ny < h and 0 <= nx < w and bg[ny, nx] and not seen[ny, nx]:
                seen[ny, nx] = True; dq.append((ny, nx))
    al2 = np.where(seen, 0, 255).astype(np.uint8)
    out = np.dstack([a[..., 0].astype(np.uint8), a[..., 1].astype(np.uint8), a[..., 2].astype(np.uint8), al2])
    im2 = Image.fromarray(out, "RGBA")
    # 트리밍 + 정사각 패딩
    bbox = im2.getbbox()
    if bbox:
        im2 = im2.crop(bbox)
    side = max(im2.size); pad = int(side * 0.06)
    canvas = Image.new("RGBA", (side + 2 * pad, side + 2 * pad), (0, 0, 0, 0))
    canvas.paste(im2, ((canvas.width - im2.width) // 2, (canvas.height - im2.height) // 2))
    canvas = canvas.resize((720, 720), Image.LANCZOS)
    canvas.save(f"{dst_dir}/{os.path.basename(f)}")
    print(os.path.basename(f), "bg%:", round(seen.mean() * 100, 1))
