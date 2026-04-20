#!/usr/bin/env python3
import contextlib
import difflib
import io
import json
import math
import os
import sys
import traceback
from pathlib import Path

import numpy as np
from PIL import Image, ImageDraw, ImageFilter, ImageFont

os.environ.setdefault("PYTORCH_ENABLE_MPS_FALLBACK", "1")

try:
    import cv2  # type: ignore
except Exception:
    cv2 = None


_ANYTEXT_MANAGER = None
_ANYTEXT_DEVICE = None
_ANYTEXT_IMPORT_ERROR = None


def emit(payload):
    sys.stdout.write(json.dumps(payload, ensure_ascii=False))
    sys.stdout.flush()


def read_payload():
    raw = sys.stdin.read()
    if not raw.strip():
        return {}
    return json.loads(raw)


def hex_color(rgb):
    return "#{:02x}{:02x}{:02x}".format(*[int(max(0, min(255, v))) for v in rgb])


def hex_to_rgb(value, fallback=(255, 255, 255)):
    try:
        value = str(value or "").strip()
        if len(value) != 7 or not value.startswith("#"):
            return fallback
        return tuple(int(value[i : i + 2], 16) for i in (1, 3, 5))
    except Exception:
        return fallback


def polygon_to_bounds(points):
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    left, right = min(xs), max(xs)
    top, bottom = min(ys), max(ys)
    return {
        "left": float(left),
        "top": float(top),
        "width": float(max(1.0, right - left)),
        "height": float(max(1.0, bottom - top)),
    }


def polygon_angle(points):
    if len(points) < 2:
        return 0.0
    dx = points[1][0] - points[0][0]
    dy = points[1][1] - points[0][1]
    return math.degrees(math.atan2(dy, dx))


def get_ocr():
    from paddleocr import PaddleOCR  # type: ignore

    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        return PaddleOCR(use_angle_cls=True, lang="ch", show_log=False)


def build_polygon_mask(size, polygon):
    mask = Image.new("L", size, 0)
    draw = ImageDraw.Draw(mask)
    draw.polygon([(p[0], p[1]) for p in polygon], fill=255)
    return np.array(mask)


def dilate_mask(mask, radius):
    radius = max(1, int(radius))
    if cv2 is not None:
        kernel = np.ones((radius * 2 + 1, radius * 2 + 1), np.uint8)
        return cv2.dilate(mask.astype(np.uint8), kernel, iterations=1)

    pil_mask = Image.fromarray(mask.astype(np.uint8), mode="L")
    for _ in range(radius):
        pil_mask = pil_mask.filter(ImageFilter.MaxFilter(3))
    return np.array(pil_mask)


def soften_mask(mask, radius=3):
    pil_mask = Image.fromarray(mask.astype(np.uint8), mode="L")
    if radius > 0:
        pil_mask = pil_mask.filter(ImageFilter.GaussianBlur(radius=radius))
    return np.array(pil_mask).astype(np.float32) / 255.0


def estimate_text_region(image_np, polygon):
    h, w = image_np.shape[:2]
    polygon_mask = build_polygon_mask((w, h), polygon)
    ys, xs = np.where(polygon_mask > 0)
    if len(xs) == 0:
        bounds = polygon_to_bounds(polygon)
        return {
            "mask": polygon_mask,
            "background": np.array([255, 255, 255], dtype=np.float32),
            "foreground": np.array([0, 0, 0], dtype=np.float32),
            "text_bounds": bounds,
        }

    left, right = xs.min(), xs.max()
    top, bottom = ys.min(), ys.max()
    crop = image_np[top : bottom + 1, left : right + 1]
    crop_mask = polygon_mask[top : bottom + 1, left : right + 1] > 0

    border_pixels = []
    if crop.shape[0] > 0 and crop.shape[1] > 0:
        border_pixels.extend(crop[0, :, :].tolist())
        border_pixels.extend(crop[-1, :, :].tolist())
        border_pixels.extend(crop[:, 0, :].tolist())
        border_pixels.extend(crop[:, -1, :].tolist())

    background = np.median(np.array(border_pixels or crop.reshape(-1, 3)), axis=0)
    diff = np.linalg.norm(crop.astype(np.float32) - background.astype(np.float32), axis=2)
    active = diff[crop_mask]
    threshold = max(18.0, float(np.percentile(active, 78))) if active.size else 18.0
    text_crop_mask = (diff >= threshold) & crop_mask

    if cv2 is not None:
        kernel_close = np.ones((3, 3), np.uint8)
        kernel_open = np.ones((2, 2), np.uint8)
        text_crop_mask = cv2.morphologyEx(text_crop_mask.astype(np.uint8), cv2.MORPH_CLOSE, kernel_close) > 0
        text_crop_mask = cv2.morphologyEx(text_crop_mask.astype(np.uint8), cv2.MORPH_OPEN, kernel_open) > 0

    full_mask = np.zeros((h, w), dtype=np.uint8)
    full_mask[top : bottom + 1, left : right + 1] = text_crop_mask.astype(np.uint8) * 255

    if np.any(text_crop_mask):
        rel_ys, rel_xs = np.where(text_crop_mask)
        text_bounds = {
            "left": float(left + rel_xs.min()),
            "top": float(top + rel_ys.min()),
            "width": float(max(1, rel_xs.max() - rel_xs.min() + 1)),
            "height": float(max(1, rel_ys.max() - rel_ys.min() + 1)),
        }
        foreground = np.median(crop[text_crop_mask], axis=0)
    else:
        text_bounds = polygon_to_bounds(polygon)
        foreground = np.array([0, 0, 0], dtype=np.float32)

    return {
        "mask": full_mask,
        "background": background.astype(np.float32),
        "foreground": foreground.astype(np.float32),
        "text_bounds": text_bounds,
    }


def sample_text_style(image_np, polygon, text):
    region = estimate_text_region(image_np, polygon)
    bounds = region["text_bounds"]
    if bounds["width"] <= 0 or bounds["height"] <= 0:
        return {
            "fontSize": 16,
            "textColor": "#000000",
            "backgroundColor": "#ffffff",
            "rotation": 0,
            "lineCount": 1,
            "align": "left",
            "familyHint": "sans",
            "shadowColor": "#000000",
            "shadowOpacity": 0.0,
            "shadowOffsetX": 0,
            "shadowOffsetY": 0,
            "textBounds": polygon_to_bounds(polygon),
        }

    bg_rgb = tuple(int(round(v)) for v in region["background"].tolist())
    fg_rgb = tuple(int(round(v)) for v in region["foreground"].tolist())
    lines = max(1, str(text).count("\n") + 1)
    font_size = max(12, int(bounds["height"] * 0.88 / lines))
    align = "center" if bounds["width"] > 300 and len(str(text)) <= 20 else "left"
    family_hint = "serif" if bounds["height"] >= 42 or "目录" in str(text) or len(str(text)) <= 14 else "sans"
    shadow_opacity = 0.18 if bounds["height"] >= 42 else 0.08

    return {
        "fontSize": font_size,
        "textColor": hex_color(fg_rgb),
        "backgroundColor": hex_color(bg_rgb),
        "rotation": polygon_angle(polygon),
        "lineCount": lines,
        "align": align,
        "familyHint": family_hint,
        "shadowColor": "#000000",
        "shadowOpacity": shadow_opacity,
        "shadowOffsetX": max(1, int(bounds["height"] * 0.03)),
        "shadowOffsetY": max(2, int(bounds["height"] * 0.08)),
        "shadowBlur": max(1, int(bounds["height"] * 0.06)),
        "strokeColor": hex_color(tuple(max(0, min(255, int(v * 0.25))) for v in fg_rgb)),
        "strokeWidth": max(0, round(bounds["height"] * 0.02, 2)),
        "letterSpacing": 0.0,
        "lineHeight": 1.0,
        "opacity": 1.0,
        "blendMode": "normal",
        "textBounds": bounds,
    }


def estimate_background_complexity(image_np, polygon):
    region = estimate_text_region(image_np, polygon)
    mask = region["mask"] > 0
    if not np.any(mask):
        return "medium"

    bounds = region["text_bounds"]
    left = int(bounds["left"])
    top = int(bounds["top"])
    right = int(bounds["left"] + bounds["width"])
    bottom = int(bounds["top"] + bounds["height"])
    crop = image_np[max(0, top - 8): min(image_np.shape[0], bottom + 8), max(0, left - 8): min(image_np.shape[1], right + 8)]
    if crop.size == 0:
        return "medium"

    std = float(np.std(crop.astype(np.float32)))
    if cv2 is not None:
        gray = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
        edge_density = float((cv2.Canny(gray, 64, 128) > 0).mean())
    else:
        edge_density = 0.0

    score = std / 64.0 + edge_density * 2.2
    if score < 0.55:
        return "simple"
    if score < 1.05:
        return "medium"
    return "complex"


def estimate_style_complexity(style):
    font_size = float(style.get("fontSize") or 0.0)
    shadow = float(style.get("shadowOpacity") or 0.0)
    stroke = float(style.get("strokeWidth") or 0.0)
    rotation = abs(float(style.get("rotation") or 0.0))
    if shadow > 0.2 or stroke > 1.8 or rotation > 8:
        return "styled"
    if font_size >= 58:
        return "styled"
    return "plain"


def split_text_char_boxes(text, bounds):
    chars = [ch for ch in str(text or "") if not ch.isspace()]
    if not chars:
        return []
    count = max(1, len(chars))
    width = float(bounds["width"])
    height = float(bounds["height"])
    cell_w = max(1.0, width / count)
    return [
        {
            "char": ch,
            "index": idx,
            "bounds": {
                "left": float(bounds["left"] + idx * cell_w),
                "top": float(bounds["top"]),
                "width": float(cell_w),
                "height": float(height),
            },
        }
        for idx, ch in enumerate(chars)
    ]


def build_font_candidates(repo_root, style, text):
    fonts_dir = Path(repo_root) / "Fonts"
    family_hint = str(style.get("familyHint") or "").lower()
    is_cjk = any("\u4e00" <= ch <= "\u9fff" for ch in str(text or ""))
    candidates = []

    def push(family, confidence, source="system", font_path=None):
        candidate_id = f"{source}:{family}".replace(" ", "_")
        if any(item["family"] == family for item in candidates):
            return
        item = {
            "candidateId": candidate_id,
            "family": family,
            "confidence": round(float(confidence), 4),
            "source": source,
        }
        if font_path:
            item["fontPath"] = str(font_path)
        item["previewText"] = str(text or "")[:24]
        candidates.append(item)

    if is_cjk:
        if "serif" in family_hint or "song" in family_hint:
            push("Songti SC", 0.92)
            push("SimSun", 0.9)
            push("Source Han Serif SC", 0.84)
        elif "kai" in family_hint:
            push("Kaiti SC", 0.92)
            push("STKaiti", 0.88)
            push("Songti SC", 0.76)
        else:
            push("PingFang SC", 0.95)
            push("Microsoft YaHei", 0.9)
            push("Source Han Sans SC", 0.86)
    else:
        push("Arial", 0.9)
        push("Helvetica", 0.85)
        push("Times New Roman", 0.72)

    if fonts_dir.exists():
        for entry in sorted(fonts_dir.iterdir()):
            if entry.suffix.lower() not in {".ttf", ".otf", ".ttc", ".woff", ".woff2"}:
                continue
            push(entry.stem, 0.68, source="workspace", font_path=entry)
            if len(candidates) >= 6:
                break

    return candidates[:6]


def infer_text_direction(text, bounds):
    if "\n" not in str(text or "") and float(bounds["height"]) > float(bounds["width"]) * 1.5:
        return "ttb"
    return "ltr"


def detect_text_boxes(payload):
    image_path = payload["image_path"]
    image = Image.open(image_path).convert("RGB")
    image_np = np.array(image)
    ocr = get_ocr()
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        result = ocr.ocr(image_path, cls=True)
    boxes = []
    reading_order = 0

    for page in result or []:
        for line in page or []:
            polygon = [[float(p[0]), float(p[1])] for p in line[0]]
            text = str(line[1][0] if line[1] else "")
            confidence = float(line[1][1] if line[1] and len(line[1]) > 1 else 0.0)
            bounds = polygon_to_bounds(polygon)
            style = sample_text_style(image_np, polygon, text)
            char_boxes = split_text_char_boxes(text, style.get("textBounds") or bounds)
            background_complexity = estimate_background_complexity(image_np, polygon)
            style_complexity = estimate_style_complexity(style)
            text_direction = infer_text_direction(text, bounds)
            reading_order += 1
            box_id = f"box_{reading_order:03d}"
            boxes.append(
                {
                    "boxId": box_id,
                    "text": text,
                    "confidence": confidence,
                    "polygon": polygon,
                    "bounds": bounds,
                    "readingOrder": reading_order,
                    "styleHint": style,
                    "charBoxes": char_boxes,
                    "rotation": float(style.get("rotation") or 0.0),
                    "skew": 0.0,
                    "textDirection": text_direction,
                    "backgroundComplexity": background_complexity,
                    "styleComplexity": style_complexity,
                    "fontCandidates": build_font_candidates(payload.get("repo_root", ""), style, text),
                    "styleEstimate": {
                        **style,
                        "textDirection": text_direction,
                        "skewX": 0.0,
                        "skewY": 0.0,
                    },
                }
            )

    boxes.sort(key=lambda item: (round(item["bounds"]["top"] / 6), item["bounds"]["left"]))
    for index, item in enumerate(boxes, start=1):
        item["readingOrder"] = index
        item["boxId"] = f"box_{index:03d}"

    return {
        "success": True,
        "engine": "paddleocr",
        "canvasWidth": image.width,
        "canvasHeight": image.height,
        "boxes": boxes,
    }


def find_font_path(repo_root, family_hint):
    candidates = []
    fonts_dir = Path(repo_root) / "Fonts"
    lower = str(family_hint or "").lower()
    if "song" in lower or "宋" in lower:
        candidates += [fonts_dir / "simsun.ttc", Path("/System/Library/Fonts/Supplemental/Songti.ttc")]
    elif "fang" in lower or "仿" in lower:
        candidates += [fonts_dir / "STFANGSO.TTF", Path("/System/Library/Fonts/Supplemental/Songti.ttc")]
    elif "kai" in lower or "楷" in lower:
        candidates += [fonts_dir / "STKAITI.TTF", Path("/System/Library/Fonts/Supplemental/Songti.ttc")]
    elif "hei" in lower or "黑" in lower:
        candidates += [Path("/System/Library/Fonts/PingFang.ttc")]
    else:
        candidates += [Path("/System/Library/Fonts/PingFang.ttc"), Path("/System/Library/Fonts/Supplemental/Songti.ttc")]

    for candidate in candidates:
        if candidate.exists():
            return str(candidate)
    return None


def clear_region(image_np, mask, background_color):
    if cv2 is None:
        pil_img = Image.fromarray(image_np)
        draw = ImageDraw.Draw(pil_img)
        ys, xs = np.where(mask > 0)
        if len(xs) > 0:
            left, right = xs.min(), xs.max()
            top, bottom = ys.min(), ys.max()
            draw.rectangle((left, top, right, bottom), fill=background_color)
        return np.array(pil_img)
    return cv2.inpaint(image_np, mask.astype(np.uint8), 3, cv2.INPAINT_TELEA)


def fit_font(draw, text, font_path, target_w, target_h, base_size):
    size = max(8, int(base_size))
    while size >= 8:
        font = ImageFont.truetype(font_path, size=size) if font_path else ImageFont.load_default()
        bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=max(2, int(size * 0.25)))
        width = bbox[2] - bbox[0]
        height = bbox[3] - bbox[1]
        if width <= target_w and height <= target_h:
            return font, width, height, size
        size -= 1
    font = ImageFont.truetype(font_path, size=8) if font_path else ImageFont.load_default()
    bbox = draw.multiline_textbbox((0, 0), text, font=font, spacing=2)
    return font, bbox[2] - bbox[0], bbox[3] - bbox[1], 8


def draw_text_with_effects(base_image, text, style, bounds, repo_root):
    font_path = find_font_path(repo_root, style.get("familyHint") or "")
    draw = ImageDraw.Draw(base_image)
    target_w = max(8, int(bounds["width"]))
    target_h = max(8, int(bounds["height"]))
    font, text_w, text_h, _ = fit_font(
        draw,
        text,
        font_path,
        target_w,
        target_h,
        style.get("fontSize") or target_h * 0.8,
    )

    left = float(bounds["left"])
    top = float(bounds["top"])
    x = left
    if style.get("align") == "center":
        x = left + max(0, (target_w - text_w) / 2)
    elif style.get("align") == "right":
        x = left + max(0, target_w - text_w)
    y = top + max(0, (target_h - text_h) / 2)

    spacing = max(2, int((style.get("fontSize") or 14) * 0.22))
    shadow_color = style.get("shadowColor", "#000000")
    shadow_opacity = float(style.get("shadowOpacity") or 0.0)
    shadow_offset_x = int(style.get("shadowOffsetX") or 0)
    shadow_offset_y = int(style.get("shadowOffsetY") or 0)

    if shadow_opacity > 0 and text:
        shadow_layer = Image.new("RGBA", base_image.size, (0, 0, 0, 0))
        shadow_draw = ImageDraw.Draw(shadow_layer)
        shadow_rgb = hex_to_rgb(shadow_color, fallback=(0, 0, 0))
        shadow_draw.multiline_text(
            (x + shadow_offset_x, y + shadow_offset_y),
            text,
            fill=(*shadow_rgb, int(255 * shadow_opacity)),
            font=font,
            spacing=spacing,
            align=style.get("align") or "left",
        )
        blur_radius = max(1, int((style.get("fontSize") or 14) * 0.08))
        shadow_layer = shadow_layer.filter(ImageFilter.GaussianBlur(radius=blur_radius))
        base_image = Image.alpha_composite(base_image.convert("RGBA"), shadow_layer).convert("RGBA")

    draw = ImageDraw.Draw(base_image)
    if text:
        draw.multiline_text(
            (x, y),
            text,
            fill=style.get("textColor", "#000000"),
            font=font,
            spacing=spacing,
            align=style.get("align") or "left",
        )
    return base_image


def safe_target_text(value):
    text = str(value or "").replace('"', " ").strip()
    return text


def normalize_compare_text(value):
    return "".join(ch for ch in str(value or "") if not ch.isspace()).replace("“", "").replace("”", "").replace('"', "")


def build_anytext_prompt(box, edit):
    style = box.get("styleHint") or {}
    bounds = box.get("bounds") or style.get("textBounds") or {}
    to_text = safe_target_text(edit.get("toText"))
    family_hint = str(style.get("familyHint") or "").lower()
    align = style.get("align") or "left"
    font_size = float(style.get("fontSize") or bounds.get("height") or 24)
    brightness = np.mean(hex_to_rgb(style.get("textColor"), fallback=(20, 20, 20)))

    role = "large Chinese title text" if font_size >= 72 else "Chinese subtitle text" if font_size >= 36 else "Chinese body text"
    family = "serif Chinese font" if "serif" in family_hint or "song" in family_hint else "sans Chinese font"
    tone = "light text" if brightness >= 180 else "dark text"
    shadow = "subtle shadow" if float(style.get("shadowOpacity") or 0.0) > 0.05 else "clean edges"

    return (
        f"{role}, {family}, {tone}, {align} aligned, {shadow}, "
        f"text reads \"{to_text}\", clear Chinese strokes, best quality"
    )


def build_anytext_negative_prompt():
    return (
        "extra text, wrong Chinese characters, malformed letters, deformed text, broken text, "
        "blurred text, messy strokes, duplicated text, extra icon, extra symbol, watermark, logo, layout change, "
        "extra decoration, background distortion, low quality"
    )


def choose_anytext_device():
    import torch  # type: ignore

    if getattr(torch.backends, "mps", None) and torch.backends.mps.is_available():
        return torch.device("mps")
    return torch.device("cpu")


def validate_rendered_text(image_np, target_text):
    normalized_target = normalize_compare_text(target_text)
    if not normalized_target:
        return True, {"score": 1.0, "detectedText": ""}

    ocr = get_ocr()
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        result = ocr.ocr(image_np, cls=True)

    candidates = []
    for page in result or []:
        for line in page or []:
            text = str(line[1][0] if line[1] else "")
            normalized = normalize_compare_text(text)
            if normalized:
                candidates.append(normalized)

    if not candidates:
        return False, {"score": 0.0, "detectedText": ""}

    best_text = ""
    best_score = 0.0
    for candidate in candidates:
        score = difflib.SequenceMatcher(None, candidate, normalized_target).ratio()
        if normalized_target in candidate or candidate in normalized_target:
            score = max(score, 0.92)
        if score > best_score:
            best_score = score
            best_text = candidate

    return best_score >= 0.78, {"score": best_score, "detectedText": best_text}


def probe_anytext_import():
    try:
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            import torch  # type: ignore  # noqa: F401
            from iopaint.model_manager import ModelManager  # type: ignore  # noqa: F401
            from iopaint.schema import HDStrategy, InpaintRequest  # type: ignore  # noqa: F401
        device = choose_anytext_device()
        return True, {"device": str(device)}
    except Exception as exc:
        return False, {"error": str(exc)}


def get_anytext_manager():
    global _ANYTEXT_MANAGER, _ANYTEXT_DEVICE, _ANYTEXT_IMPORT_ERROR

    if _ANYTEXT_MANAGER is not None:
        return _ANYTEXT_MANAGER, _ANYTEXT_DEVICE
    if _ANYTEXT_IMPORT_ERROR is not None:
        raise RuntimeError(_ANYTEXT_IMPORT_ERROR)

    try:
        import torch  # type: ignore
        from iopaint.const import ANYTEXT_NAME  # type: ignore
        from iopaint.model import models  # type: ignore
        from iopaint.model_manager import ModelManager  # type: ignore

        device = choose_anytext_device()
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            if ANYTEXT_NAME in models:
                models[ANYTEXT_NAME].download(local_files_only=False)
            manager = ModelManager(
                name=ANYTEXT_NAME,
                device=device,
                disable_nsfw=True,
                sd_cpu_textencoder=str(device) == "cpu",
            )
        _ANYTEXT_MANAGER = manager
        _ANYTEXT_DEVICE = str(device)
        return manager, _ANYTEXT_DEVICE
    except Exception as exc:
        _ANYTEXT_IMPORT_ERROR = str(exc)
        raise RuntimeError(_ANYTEXT_IMPORT_ERROR)


def crop_rect_from_bounds(image_shape, bounds, padding):
    height, width = image_shape[:2]
    left = max(0, int(math.floor(float(bounds.get("left", 0)) - padding)))
    top = max(0, int(math.floor(float(bounds.get("top", 0)) - padding)))
    right = min(width, int(math.ceil(float(bounds.get("left", 0)) + float(bounds.get("width", 1)) + padding)))
    bottom = min(height, int(math.ceil(float(bounds.get("top", 0)) + float(bounds.get("height", 1)) + padding)))
    return {
        "left": left,
        "top": top,
        "right": max(left + 1, right),
        "bottom": max(top + 1, bottom),
    }


def build_rect_mask(image_shape, bounds, pad_x, pad_y):
    height, width = image_shape[:2]
    left = max(0, int(math.floor(float(bounds.get("left", 0)) - pad_x)))
    top = max(0, int(math.floor(float(bounds.get("top", 0)) - pad_y)))
    right = min(width, int(math.ceil(float(bounds.get("left", 0)) + float(bounds.get("width", 1)) + pad_x)))
    bottom = min(height, int(math.ceil(float(bounds.get("top", 0)) + float(bounds.get("height", 1)) + pad_y)))
    mask = np.zeros((height, width), dtype=np.uint8)
    mask[top:bottom, left:right] = 255
    return mask, {
        "left": left,
        "top": top,
        "right": max(left + 1, right),
        "bottom": max(top + 1, bottom),
    }


def sample_background_around_bounds(image_np, mask, rect):
    height, width = image_np.shape[:2]
    outer_pad = 10
    left = max(0, rect["left"] - outer_pad)
    top = max(0, rect["top"] - outer_pad)
    right = min(width, rect["right"] + outer_pad)
    bottom = min(height, rect["bottom"] + outer_pad)
    crop = image_np[top:bottom, left:right]
    crop_mask = mask[top:bottom, left:right] > 0
    ring_pixels = crop[~crop_mask]
    if ring_pixels.size == 0:
        return tuple(int(round(v)) for v in np.median(crop.reshape(-1, 3), axis=0))
    return tuple(int(round(v)) for v in np.median(ring_pixels, axis=0))


def build_cleanup_mask(image_np, box, strategy="local_inpaint"):
    style = (box.get("styleEstimate") or box.get("styleHint") or {})
    bounds = style.get("textBounds") or box.get("bounds") or polygon_to_bounds(box["polygon"])
    height = max(1.0, float(bounds.get("height", 1)))
    stroke_width = float(style.get("strokeWidth") or 0.0)
    shadow_blur = float(style.get("shadowBlur") or 0.0)
    shadow_offset_x = abs(float(style.get("shadowOffsetX") or 0.0))
    shadow_offset_y = abs(float(style.get("shadowOffsetY") or 0.0))
    letter_spacing = max(0.0, float(style.get("letterSpacing") or 0.0))

    pad_x = max(8, int(round(height * 0.22 + stroke_width * 3 + shadow_offset_x + shadow_blur + letter_spacing * 0.5)))
    pad_y = max(6, int(round(height * 0.24 + stroke_width * 3 + shadow_offset_y + shadow_blur)))

    region = estimate_text_region(image_np, box["polygon"])
    text_mask = dilate_mask(region["mask"], max(3, int(round(height * 0.08))))
    rect_mask, rect = build_rect_mask(image_np.shape, bounds, pad_x, pad_y)
    if strategy == "analytic_fill":
        cleanup_mask = np.maximum(rect_mask, text_mask)
    else:
        char_mask = np.zeros(image_np.shape[:2], dtype=np.uint8)
        char_boxes = box.get("charBoxes") or []
        if char_boxes:
            per_char_pad_x = max(3, int(round(height * 0.12 + stroke_width * 2 + shadow_offset_x * 0.5 + shadow_blur * 0.4)))
            per_char_pad_y = max(3, int(round(height * 0.16 + stroke_width * 2 + shadow_offset_y + shadow_blur * 0.5)))
            for char_box in char_boxes:
                char_bounds = char_box.get("bounds") or {}
                char_rect_mask, _ = build_rect_mask(image_np.shape, char_bounds, per_char_pad_x, per_char_pad_y)
                char_mask = np.maximum(char_mask, char_rect_mask)
            cleanup_mask = np.maximum(text_mask, char_mask)
        else:
            cleanup_mask = text_mask
        cleanup_mask = dilate_mask(cleanup_mask, max(1, int(round(height * 0.03))))
    background_rgb = sample_background_around_bounds(image_np, cleanup_mask, rect)
    return cleanup_mask, background_rgb, rect


def erode_mask(mask, radius):
    radius = max(1, int(radius))
    if cv2 is not None:
        kernel = np.ones((radius * 2 + 1, radius * 2 + 1), np.uint8)
        return cv2.erode(mask.astype(np.uint8), kernel, iterations=1)

    pil_mask = Image.fromarray(mask.astype(np.uint8), mode="L")
    for _ in range(radius):
        pil_mask = pil_mask.filter(ImageFilter.MinFilter(3))
    return np.array(pil_mask)


def simple_fill_region(image_np, mask, background_rgb):
    filled = image_np.copy()
    active = mask > 0
    if not np.any(active):
        return filled
    filled[active] = np.array(background_rgb, dtype=np.uint8)
    if cv2 is not None:
        blurred = cv2.GaussianBlur(filled, (0, 0), sigmaX=3.0, sigmaY=3.0)
        alpha = (soften_mask(mask, radius=4))[..., np.newaxis]
        return (blurred.astype(np.float32) * alpha + image_np.astype(np.float32) * (1.0 - alpha)).clip(0, 255).astype(np.uint8)
    return filled


def cleanup_text_boxes(payload):
    image_path = payload["image_path"]
    output_path = payload["output_path"]
    boxes = {box["boxId"]: box for box in payload.get("boxes", [])}
    target_ids = payload.get("box_ids") or [edit.get("boxId") for edit in payload.get("edits", [])]
    image = Image.open(image_path).convert("RGB")
    image_np = np.array(image)
    logs = []

    force_strategy = str(payload.get("force_strategy") or "").strip()

    for box_id in target_ids:
        box = boxes.get(box_id)
        if not box:
            logs.append({"boxId": box_id, "success": False, "error": "box not found"})
            continue
        style = box.get("styleHint") or {}
        polygon = box["polygon"]
        original_image_np = image_np.copy()
        complexity = box.get("backgroundComplexity") or estimate_background_complexity(image_np, polygon)
        strategy = force_strategy if force_strategy in {"analytic_fill", "local_inpaint"} else ("analytic_fill" if complexity == "simple" else "local_inpaint")
        cleanup_mask, background_rgb, cleanup_rect = build_cleanup_mask(image_np, box, strategy=strategy)
        if strategy == "analytic_fill":
            cleaned_np = simple_fill_region(image_np, cleanup_mask, background_rgb)
        else:
            cleaned_np = clear_region(image_np, cleanup_mask, background_rgb)

        blend_radius = max(4, int(round(max(cleanup_rect["bottom"] - cleanup_rect["top"], cleanup_rect["right"] - cleanup_rect["left"]) * 0.03)))
        core_mask = erode_mask(cleanup_mask, max(1, blend_radius // 2))
        alpha_soft = soften_mask(cleanup_mask, radius=blend_radius)
        alpha = np.maximum(alpha_soft, (core_mask > 0).astype(np.float32))[..., np.newaxis]
        image_np = (
            cleaned_np.astype(np.float32) * alpha +
            original_image_np.astype(np.float32) * (1.0 - alpha)
        ).clip(0, 255).astype(np.uint8)
        logs.append(
            {
                "boxId": box_id,
                "success": True,
                "cleanupStrategy": strategy,
                "backgroundComplexity": complexity,
                "backgroundColor": style.get("backgroundColor") or hex_color(background_rgb),
                "cleanupRect": cleanup_rect,
                "blendRadius": blend_radius,
            }
        )

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(image_np).save(output_path, format="PNG")
    return {
        "success": True,
        "outputPath": output_path,
        "logs": logs,
    }


def recognize_text(payload):
    image_path = payload["image_path"]
    image = Image.open(image_path).convert("RGB")
    ocr = get_ocr()
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        result = ocr.ocr(image_path, cls=True)

    lines = []
    for page in result or []:
        for line in page or []:
            text = str(line[1][0] if line[1] else "")
            confidence = float(line[1][1] if line[1] and len(line[1]) > 1 else 0.0)
            polygon = [[float(p[0]), float(p[1])] for p in line[0]]
            lines.append(
                {
                    "text": text,
                    "confidence": confidence,
                    "polygon": polygon,
                    "bounds": polygon_to_bounds(polygon),
                }
            )

    combined = "\n".join(item["text"] for item in lines if item["text"])
    return {
        "success": True,
        "text": combined,
        "lines": lines,
        "canvasWidth": image.width,
        "canvasHeight": image.height,
    }


def recognize_texts_batch(payload):
    image_paths = payload.get("image_paths") or []
    if not isinstance(image_paths, list) or not image_paths:
        return {
            "success": False,
            "error": "missing image_paths",
            "items": [],
        }

    ocr = get_ocr()
    items = []
    for image_path in image_paths:
        image = Image.open(image_path).convert("RGB")
        with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
            result = ocr.ocr(image_path, cls=True)

        lines = []
        for page in result or []:
            for line in page or []:
                text = str(line[1][0] if line[1] else "")
                confidence = float(line[1][1] if line[1] and len(line[1]) > 1 else 0.0)
                polygon = [[float(p[0]), float(p[1])] for p in line[0]]
                lines.append(
                    {
                        "text": text,
                        "confidence": confidence,
                        "polygon": polygon,
                        "bounds": polygon_to_bounds(polygon),
                    }
                )

        combined = "\n".join(item["text"] for item in lines if item["text"])
        items.append(
            {
                "success": True,
                "text": combined,
                "lines": lines,
                "canvasWidth": image.width,
                "canvasHeight": image.height,
            }
        )

    return {
        "success": True,
        "items": items,
    }


def apply_simple_redraw_edit(image_np, edit, box, repo_root):
    style = box.get("styleHint") or {}
    polygon = box["polygon"]
    bounds = style.get("textBounds") or edit.get("bounds") or box["bounds"]
    background_color = style.get("backgroundColor", "#ffffff")
    to_text = str(edit.get("toText") or "").strip()

    region = estimate_text_region(image_np, polygon)
    bg_rgb = hex_to_rgb(background_color)
    next_image_np = clear_region(image_np, region["mask"], bg_rgb)
    next_image = Image.fromarray(next_image_np).convert("RGBA")
    next_image = draw_text_with_effects(next_image, to_text, style, bounds, repo_root)
    return np.array(next_image.convert("RGB"))


def apply_anytext_edit(image_np, edit, box, anytext_manager):
    style = box.get("styleHint") or {}
    rotation = float(style.get("rotation") or 0)
    if abs(rotation) > 8:
        raise RuntimeError("rotation not supported")

    to_text = safe_target_text(edit.get("toText"))
    if not to_text:
        raise RuntimeError("empty target text")

    polygon = box["polygon"]
    region = estimate_text_region(image_np, polygon)
    mask = build_polygon_mask((image_np.shape[1], image_np.shape[0]), polygon)

    source_bounds = box["bounds"] or region["text_bounds"] or style.get("textBounds")
    padding = max(24, min(72, int(max(source_bounds["height"], 24) * 0.65)))
    dilate_radius = max(3, min(10, int(max(source_bounds["height"], 24) * 0.06)))
    expanded_mask = dilate_mask(mask, dilate_radius)
    crop_rect = crop_rect_from_bounds(image_np.shape, source_bounds, padding)

    crop = image_np[crop_rect["top"] : crop_rect["bottom"], crop_rect["left"] : crop_rect["right"]].copy()
    crop_mask = expanded_mask[crop_rect["top"] : crop_rect["bottom"], crop_rect["left"] : crop_rect["right"]]
    if not np.any(crop_mask):
        raise RuntimeError("empty crop mask")
    original_crop = crop
    original_mask = crop_mask.astype(np.uint8)

    max_side = 640
    resize_scale = 1.0
    if max(crop.shape[0], crop.shape[1]) > max_side:
        resize_scale = max_side / float(max(crop.shape[0], crop.shape[1]))
        target_w = max(64, int(round(crop.shape[1] * resize_scale / 64.0)) * 64)
        target_h = max(64, int(round(crop.shape[0] * resize_scale / 64.0)) * 64)
        if cv2 is not None:
            crop = cv2.resize(crop, (target_w, target_h), interpolation=cv2.INTER_AREA)
            crop_mask = cv2.resize(crop_mask, (target_w, target_h), interpolation=cv2.INTER_NEAREST)
        else:
            crop = np.array(Image.fromarray(crop).resize((target_w, target_h), Image.Resampling.LANCZOS))
            crop_mask = np.array(Image.fromarray(crop_mask, mode="L").resize((target_w, target_h), Image.Resampling.NEAREST))

    crop_mask = crop_mask[..., np.newaxis].astype(np.uint8)
    prepared_crop = clear_region(
        crop,
        crop_mask[:, :, 0],
        hex_to_rgb(style.get("backgroundColor"), fallback=(245, 245, 245)),
    )

    from iopaint.schema import HDStrategy, InpaintRequest  # type: ignore

    config = InpaintRequest(
        hd_strategy=HDStrategy.ORIGINAL,
        prompt=build_anytext_prompt(box, edit),
        negative_prompt=build_anytext_negative_prompt(),
        sd_steps=18,
        sd_guidance_scale=7.5,
        sd_seed=-1,
        sd_strength=0.92,
        sd_match_histograms=True,
        sd_mask_blur=9,
    )

    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        result_bgr = anytext_manager(prepared_crop, crop_mask, config)
    result_rgb = result_bgr[:, :, ::-1]
    if resize_scale != 1.0:
        if cv2 is not None:
            result_rgb = cv2.resize(
                result_rgb,
                (original_crop.shape[1], original_crop.shape[0]),
                interpolation=cv2.INTER_CUBIC,
            )
        else:
            result_rgb = np.array(
                Image.fromarray(result_rgb).resize(
                    (original_crop.shape[1], original_crop.shape[0]),
                    Image.Resampling.LANCZOS,
                )
            )
    valid, validation_meta = validate_rendered_text(result_rgb, to_text)
    if not valid:
        raise RuntimeError(
            f"rendered text mismatch: got {validation_meta.get('detectedText') or '<none>'} "
            f"(score={validation_meta.get('score', 0):.2f})"
        )
    alpha = soften_mask(original_mask, radius=3)[..., np.newaxis]
    merged_crop = (result_rgb.astype(np.float32) * alpha + original_crop.astype(np.float32) * (1.0 - alpha)).clip(0, 255).astype(np.uint8)

    next_image_np = image_np.copy()
    next_image_np[crop_rect["top"] : crop_rect["bottom"], crop_rect["left"] : crop_rect["right"]] = merged_crop
    return next_image_np


def apply_text_edits(payload):
    image_path = payload["image_path"]
    output_path = payload["output_path"]
    boxes = {box["boxId"]: box for box in payload.get("boxes", [])}
    edits = payload.get("edits", [])
    repo_root = payload.get("repo_root", "")
    prefer_high_fidelity = bool(payload.get("prefer_high_fidelity", True))

    image = Image.open(image_path).convert("RGB")
    image_np = np.array(image)
    logs = []
    applied_count = 0
    anytext_manager = None
    anytext_device = None
    anytext_available = False
    anytext_error = None
    anytext_used = False
    fallback_used = False

    if prefer_high_fidelity:
        try:
            anytext_manager, anytext_device = get_anytext_manager()
            anytext_available = True
        except Exception as exc:
            anytext_error = str(exc)

    for edit in edits:
        box_id = edit.get("boxId")
        box = boxes.get(box_id)
        if not box:
            logs.append({"boxId": box_id, "success": False, "error": "box not found"})
            continue

        style = box.get("styleHint") or {}
        rotation = float(style.get("rotation") or 0)
        if abs(rotation) > 8:
            logs.append({"boxId": box_id, "success": False, "error": "rotation not supported"})
            continue

        to_text = str(edit.get("toText") or "").strip()
        if prefer_high_fidelity and anytext_available and to_text:
            try:
                image_np = apply_anytext_edit(image_np, edit, box, anytext_manager)
                anytext_used = True
                applied_count += 1
                logs.append(
                    {
                        "boxId": box_id,
                        "success": True,
                        "toText": to_text,
                        "engine": "iopaint_anytext",
                        "device": anytext_device,
                    }
                )
                continue
            except Exception as exc:
                fallback_used = True
                image_np = apply_simple_redraw_edit(image_np, edit, box, repo_root)
                applied_count += 1
                logs.append(
                    {
                        "boxId": box_id,
                        "success": True,
                        "toText": to_text,
                        "engine": "simple_redraw",
                        "warning": f"AnyText failed: {exc}",
                    }
                )
                continue

        image_np = apply_simple_redraw_edit(image_np, edit, box, repo_root)
        applied_count += 1
        log_item = {
            "boxId": box_id,
            "success": True,
            "toText": to_text,
            "engine": "simple_redraw",
        }
        if prefer_high_fidelity and anytext_error and to_text:
            log_item["warning"] = f"AnyText unavailable: {anytext_error}"
            fallback_used = True
        logs.append(log_item)

    if applied_count == 0:
        return {
            "success": False,
            "error": "没有成功应用任何文字修改",
            "logs": logs,
            "highFidelityAvailable": anytext_available,
            "highFidelityError": anytext_error,
        }

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    Image.fromarray(image_np).save(output_path, format="PNG")

    if anytext_used and fallback_used:
        engine = "hybrid_anytext_redraw"
    elif anytext_used:
        engine = "iopaint_anytext"
    else:
        engine = "simple_redraw"

    return {
        "success": True,
        "engine": engine,
        "outputPath": output_path,
        "logs": logs,
        "highFidelityAvailable": anytext_available,
        "highFidelityUsed": anytext_used,
        "highFidelityDevice": anytext_device,
        "highFidelityError": anytext_error,
    }


def health(payload):
    try:
        import importlib.util

        paddle_ok = importlib.util.find_spec("paddleocr") is not None
        paddle_error = None if paddle_ok else "paddleocr not installed"
        anytext_ok = (
            importlib.util.find_spec("iopaint") is not None
            and importlib.util.find_spec("torch") is not None
        )
        anytext_meta = {"device": None}
        if anytext_ok:
            try:
                anytext_meta["device"] = str(choose_anytext_device())
            except Exception:
                anytext_meta["device"] = None
        else:
            anytext_meta["error"] = "iopaint or torch not installed"

        return {
            "ready": paddle_ok,
            "python": sys.executable,
            "ocrEngine": "paddleocr",
            "editEngine": "deterministic_pipeline",
            "highFidelityAvailable": anytext_ok,
            "highFidelityDevice": anytext_meta.get("device") if anytext_ok else None,
            "highFidelityError": None if anytext_ok else anytext_meta.get("error"),
            "paddleAvailable": paddle_ok,
            "paddleError": paddle_error,
            "cleanupEngines": ["analytic_fill", "local_inpaint"],
        }
    except Exception as exc:
        return {
            "ready": False,
            "error": str(exc),
        }


COMMANDS = {
    "health": health,
    "detect_text_boxes": detect_text_boxes,
    "cleanup_text_boxes": cleanup_text_boxes,
    "recognize_text": recognize_text,
    "recognize_texts_batch": recognize_texts_batch,
    "apply_text_edits": apply_text_edits,
}


def main():
    if len(sys.argv) < 2:
        emit({"success": False, "error": "missing command"})
        return

    command = sys.argv[1]
    payload = read_payload()
    try:
        handler = COMMANDS.get(command)
        if not handler:
            emit({"success": False, "error": f"unknown command: {command}"})
            return
        result = handler(payload)
        emit(result)
    except Exception as exc:
        emit(
            {
                "success": False,
                "error": str(exc),
                "traceback": traceback.format_exc(),
            }
        )


if __name__ == "__main__":
    main()
