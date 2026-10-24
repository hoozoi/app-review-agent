"""데모 페이지의 자동 재생(?autoplay)을 녹화해 demo/demo.gif로 저장

  pip install playwright imageio-ffmpeg
  python -m playwright install chromium
  python demo/record_gif.py

브라우저 화면을 동영상으로 녹화한 뒤 ffmpeg로 GIF를 만듦 (스크린샷을 이어 붙이는 것보다 스크롤이 부드러움)
"""
import shutil
import subprocess
import tempfile
import time
from pathlib import Path

import imageio_ffmpeg
from playwright.sync_api import sync_playwright

HERE = Path(__file__).resolve().parent
VIEWPORT = {"width": 1200, "height": 760}
OUT_WIDTH = 720      # GIF 가로 크기 (README에서 보기 좋고 파일이 너무 커지지 않게)
FPS = 10
COLORS = 96
TAIL_HOLD = 0.3      # 마지막 화면을 멈춰 보여줄 시간(초)


def record(video_dir):
    """녹화한 동영상 경로와, 자동 재생이 시작된 시점(동영상 시작 기준 초)을 돌려줌"""
    with sync_playwright() as p:
        browser = p.chromium.launch()
        context = browser.new_context(viewport=VIEWPORT, record_video_dir=str(video_dir), record_video_size=VIEWPORT)
        t0 = time.time()
        page = context.new_page()
        page.goto((HERE / "index.html").as_uri() + "?autoplay")
        page.wait_for_function("document.body.dataset.started === '1'", timeout=30000)
        ready = max(0, time.time() - t0 - 0.3)
        page.wait_for_function("document.body.dataset.done === '1'", timeout=60000)
        page.wait_for_timeout(TAIL_HOLD * 1000)
        video = page.video.path()
        context.close()
        browser.close()
    return Path(video), ready


def to_gif(video, start, out):
    ffmpeg = imageio_ffmpeg.get_ffmpeg_exe()
    graph = (f"fps={FPS},scale={OUT_WIDTH}:-1:flags=lanczos,split[a][b];"
             f"[a]palettegen=max_colors={COLORS}:stats_mode=diff[p];[b][p]paletteuse=dither=none:diff_mode=rectangle")
    subprocess.run([ffmpeg, "-y", "-loglevel", "error", "-ss", f"{start:.2f}", "-i", str(video),
                    "-filter_complex", graph, "-loop", "0", str(out)], check=True)


if __name__ == "__main__":
    tmp = Path(tempfile.mkdtemp())
    try:
        video, ready = record(tmp)
        out = HERE / "demo.gif"
        to_gif(video, ready, out)
        print(f"저장: {out.name} ({out.stat().st_size / 1e6:.1f}MB)")
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
