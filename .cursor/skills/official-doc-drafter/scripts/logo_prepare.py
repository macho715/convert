#!/usr/bin/env python3
"""
공식 로고 다운로드 및 SVG/PNG 준비
"""

import argparse
import os
import subprocess
from io import BytesIO

import requests
from PIL import Image


def download_logo(url: str, out_base: str) -> str:
    r = requests.get(url, timeout=15)
    r.raise_for_status()
    content = r.content
    text_start = content[:200].decode(errors="ignore")
    if "<svg" in text_start.lower():
        svg_path = out_base + ".svg"
        with open(svg_path, "wb") as f:
            f.write(content)
        return svg_path
    img = Image.open(BytesIO(content))
    png_path = out_base + ".png"
    img.save(png_path)
    return png_path


def vectorize_raster(png_path: str, svg_path: str) -> str:
    """
    potrace를 통해 벡터화 (외부 명령: ffmpeg, potrace 필요)
    """
    tmp = png_path + ".pbm"
    try:
        subprocess.run(
            [
                "ffmpeg",
                "-y",
                "-i",
                png_path,
                "-f",
                "image2",
                "-pix_fmt",
                "gray",
                tmp,
            ],
            check=True,
            capture_output=True,
        )
        subprocess.run(
            ["potrace", "-b", "svg", tmp, "-o", svg_path],
            check=True,
            capture_output=True,
        )
        if os.path.exists(tmp):
            os.remove(tmp)
        return svg_path
    except (subprocess.CalledProcessError, FileNotFoundError) as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise RuntimeError(
            "vectorize requires ffmpeg and potrace; install or skip --vectorize"
        ) from e


if __name__ == "__main__":
    p = argparse.ArgumentParser()
    p.add_argument("--url", "-u", required=True)
    p.add_argument("--out-dir", "-d", default=".")
    p.add_argument("--vectorize", "-v", action="store_true")
    args = p.parse_args()
    base = os.path.join(
        args.out_dir,
        os.path.splitext(os.path.basename(args.url.split("?")[0]))[0],
    )
    downloaded = download_logo(args.url, base)
    print("Downloaded:", downloaded)
    if args.vectorize and downloaded.endswith(".png"):
        try:
            svg = vectorize_raster(downloaded, base + ".svg")
            print("Vectorized:", svg)
        except RuntimeError as err:
            print("Vectorize skipped:", err)
