import os
import re
from dataclasses import dataclass
from typing import List, Optional

import requests
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt


INPUT_FILE = "input.txt"
OUTPUT_PPTX = "output.pptx"
META_FILE = "meta.txt"
SUMMARY_FILE = "summary.txt"
IMAGES_DIR = "images"


@dataclass
class ImageInfo:
    path: Optional[str]
    url: str
    orientation: str  # "horizontal", "vertical", or "square"
    status: str  # "downloaded" or "placeholder"


@dataclass
class SlideContent:
    keyword: str
    summary: str
    images: List[ImageInfo]


def read_keywords(path: str) -> List[str]:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    with open(path, "r", encoding="utf-8") as f:
        text = f.read()
    return re.findall(r"\[(.+?)\]", text)


def generate_summary(keyword: str) -> str:
    templates = [
        "{kw}의 핵심 개념과 최신 동향을 간단히 정리합니다.",
        "{kw} 관련 주요 활용 사례와 기대 효과를 소개합니다.",
        "{kw}를 도입할 때 고려해야 할 포인트를 요약합니다.",
        "{kw}의 배경과 향후 전망을 한눈에 살펴봅니다.",
    ]
    idx = sum(ord(c) for c in keyword) % len(templates)
    return templates[idx].format(kw=keyword)


def ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)


def fetch_image(keyword: str, idx: int) -> ImageInfo:
    ensure_dir(IMAGES_DIR)
    safe_keyword = re.sub(r"[^0-9A-Za-z\uAC00-\uD7A3]+", "_", keyword).strip("_") or "keyword"
    url = f"https://source.unsplash.com/featured/?{requests.utils.quote(keyword)}&sig={idx}"
    filename = f"{safe_keyword}_{idx}.jpg"
    local_path = os.path.join(IMAGES_DIR, filename)

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        with open(local_path, "wb") as f:
            f.write(response.content)
        orientation = detect_orientation(local_path)
        return ImageInfo(path=local_path, url=url, orientation=orientation, status="downloaded")
    except Exception:
        # Remove partial file if any
        if os.path.exists(local_path):
            try:
                os.remove(local_path)
            except OSError:
                pass
        return ImageInfo(path=None, url=url, orientation="unknown", status="placeholder")


def detect_orientation(image_path: str) -> str:
    with Image.open(image_path) as img:
        width, height = img.size
    if height > width * 1.1:
        return "vertical"
    if width > height * 1.1:
        return "horizontal"
    return "square"


def collect_images(keyword: str, count: int = 2) -> List[ImageInfo]:
    images = []
    for idx in range(count):
        images.append(fetch_image(keyword, idx))
    return images


def create_presentation(contents: List[SlideContent]) -> Presentation:
    prs = Presentation()
    for content in contents:
        slide = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
        add_text_block(slide, content.keyword, content.summary, prs)
        place_images(slide, content.images, prs)
    return prs


def add_text_block(slide, keyword: str, summary: str, prs: Presentation) -> None:
    left_margin = Inches(0.5)
    top_margin = Inches(0.5)
    width = prs.slide_width * 0.4
    height = prs.slide_height - top_margin * 2

    textbox = slide.shapes.add_textbox(left_margin, top_margin, width, height)
    tf = textbox.text_frame
    tf.clear()

    title_run = tf.paragraphs[0].add_run()
    title_run.text = keyword
    title_run.font.size = Pt(32)
    title_run.font.bold = True

    para = tf.add_paragraph()
    para.text = summary
    para.font.size = Pt(18)
    para.space_after = Pt(12)


def place_images(slide, images: List[ImageInfo], prs: Presentation) -> None:
    available_width = prs.slide_width * 0.55
    right_x = prs.slide_width * 0.45 + Inches(0.2)
    top_y = Inches(0.5)
    bottom_margin = Inches(0.5)
    available_height = prs.slide_height - top_y - bottom_margin

    valid_images = [img for img in images if img.path and os.path.exists(img.path)]
    if not valid_images:
        placeholder = slide.shapes.add_textbox(right_x, top_y, available_width, available_height)
        placeholder.text_frame.text = "이미지를 불러올 수 없습니다."
        placeholder.text_frame.paragraphs[0].font.size = Pt(18)
        return

    orientations = {img.orientation for img in valid_images}
    dominant_orientation = "horizontal" if "horizontal" in orientations or "square" in orientations else "vertical"

    if dominant_orientation in {"horizontal", "square"}:
        slot_height = available_height / max(1, len(valid_images))
        for i, img in enumerate(valid_images):
            top = top_y + slot_height * i
            height = slot_height - Inches(0.2)
            slide.shapes.add_picture(img.path, right_x, top, width=available_width, height=height)
    else:
        # Vertical: place diagonally in a 2x2 grid on the right side
        cell_width = available_width / 2
        cell_height = available_height / 2
        positions = [
            (right_x, top_y),
            (right_x + cell_width, top_y + cell_height),
        ]
        for img, (x, y) in zip(valid_images, positions):
            slide.shapes.add_picture(img.path, x, y, width=cell_width - Inches(0.1), height=cell_height - Inches(0.1))


def write_outputs(contents: List[SlideContent]) -> None:
    with open(META_FILE, "w", encoding="utf-8") as meta, open(SUMMARY_FILE, "w", encoding="utf-8") as summary:
        for item in contents:
            summary_line = f"[{item.keyword}] {item.summary}\n"
            summary.write(summary_line)

            meta.write(f"[{item.keyword}]\n")
            meta.write(f"요약: {item.summary}\n")
            for idx, img in enumerate(item.images, start=1):
                status = img.status
                local_path = img.path if img.path else "placeholder"
                meta.write(f"이미지{idx}: {local_path} | 원본: {img.url} | 상태: {status}\n")
            meta.write("\n")


def main() -> None:
    keywords = read_keywords(INPUT_FILE)
    contents: List[SlideContent] = []

    for keyword in keywords:
        summary = generate_summary(keyword)
        images = collect_images(keyword, count=2)
        contents.append(SlideContent(keyword=keyword, summary=summary, images=images))

    presentation = create_presentation(contents)
    presentation.save(OUTPUT_PPTX)
    write_outputs(contents)
    print(f"Slides generated: {len(contents)}")
    print(f"- {OUTPUT_PPTX}\n- {META_FILE}\n- {SUMMARY_FILE}")


if __name__ == "__main__":
    main()
