# 產生 earth.png：把 NASA Blue Marble 衛星影像（three.js 範例的 earth_day_4096.jpg，公有領域）轉成插畫風的平塗底圖
# 用法：python3 make_earth.py earth_day_4096.jpg earth.png 2048   （沙漠／乾草原／淺綠／深綠／雪，海透明；座標與 world.js 相同）
import sys, numpy as np
from PIL import Image, ImageFilter
SRC = sys.argv[1]; OUT = sys.argv[2]; W = int(sys.argv[3]) if len(sys.argv) > 3 else 2048
im = Image.open(SRC).convert("RGB").resize((W, W // 2), Image.LANCZOS).filter(ImageFilter.GaussianBlur(1.2))
a = np.asarray(im).astype(np.float32); r, g, b = a[..., 0], a[..., 1], a[..., 2]
v = a.max(-1); mn = a.min(-1); s = np.where(v > 0, (v - mn) / np.maximum(v, 1), 0)
gi = g - (r + b) / 2                      # 綠度
water = (b > g + 4) & (b > r + 10) & (v < 200) | ((b >= g) & (r < 0.75 * b) & (v < 160))
lat = 90 - (np.arange(v.shape[0], dtype=np.float32) + 0.5) / v.shape[0] * 180   # 每列的緯度
lat = np.repeat(lat[:, None], v.shape[1], axis=1)
# 雪：很白才算（亮沙漠的飽和度較高，不算）；高緯度放寬
snow = (~water) & (((v > 215) & (s < 0.14)) | ((np.abs(lat) > 58) & (v > 170) & (s < 0.22)))
rock = (~water) & (~snow) & (s < 0.16) & (v < 185)          # 高原、裸岩（灰）
desert = (~water) & (~snow) & (~rock) & (r > g + 6) & (gi < -4)
dense = (~water) & (~snow) & (gi > 22)
light = (~water) & (~snow) & (~dense) & (gi > 6)
# 類別：0 水 1 沙漠 2 乾草原 3 淺綠 4 深綠 5 雪 6 裸岩
cls = np.full(v.shape, 2, np.uint8)
cls[desert] = 1; cls[light] = 3; cls[dense] = 4; cls[rock] = 6; cls[snow] = 5; cls[water] = 0
# 讓色塊圓潤：眾數濾波兩次（水域另外處理，避免海岸線跑掉）
ci = Image.fromarray(cls, "L")
for _ in range(2): ci = ci.filter(ImageFilter.ModeFilter(7))
cls2 = np.asarray(ci).copy(); cls2[water] = 0; cls2[(~water) & (cls2 == 0)] = 2
PAL = {1: (233, 213, 165), 2: (214, 208, 154), 3: (184, 211, 154), 4: (134, 184, 110), 5: (243, 240, 232), 6: (217, 201, 168)}
rgb = np.zeros(a.shape, np.uint8)
for k, c in PAL.items(): rgb[cls2 == k] = c
alpha = Image.fromarray(((cls2 != 0) * 255).astype(np.uint8), "L").filter(ImageFilter.GaussianBlur(0.8))
out = Image.fromarray(rgb, "RGB").filter(ImageFilter.GaussianBlur(0.6))
out.putalpha(alpha)
out.save(OUT, optimize=True)
tot = (cls2 != 0).sum()
print("saved", OUT, out.size, "| land share by class:", {k: round(float((cls2 == k).sum()) / tot, 2) for k in PAL})
