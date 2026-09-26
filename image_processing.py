from __future__ import annotations

from dataclasses import dataclass, replace
from pathlib import Path
from typing import Tuple

import numpy as np
from PIL import Image, ImageOps

RGB = Tuple[int, int, int]


@dataclass(frozen=True)
class ProcessingSettings:
    low: int = 80
    high: int = 180
    invert: bool = False
    denoise: str = "none"
    morphology: str = "none"
    iterations: int = 1
    gray_level: int = 130
    colorized: bool = True
    black_color: RGB = (0, 0, 0)
    gray_color: RGB = (120, 120, 120)
    white_color: RGB = (255, 255, 255)

    def normalized(self) -> "ProcessingSettings":
        low = max(0, min(255, int(self.low)))
        high = max(0, min(255, int(self.high)))
        if low > high:
            low, high = high, low
        if low == high:
            if high < 255:
                high += 1
            else:
                low -= 1
        return replace(
            self,
            low=low,
            high=high,
            iterations=max(1, min(10, int(self.iterations))),
            gray_level=max(32, min(223, int(self.gray_level))),
        )


def load_image_arrays(path: str | Path):
    """Load an image as grayscale + alpha arrays while honoring EXIF orientation."""
    with Image.open(path) as src:
        image = ImageOps.exif_transpose(src).convert("RGBA")
        rgba = np.asarray(image, dtype=np.uint8)

    rgb = rgba[..., :3].astype(np.float32)
    gray = np.rint(
        0.299 * rgb[..., 0] + 0.587 * rgb[..., 1] + 0.114 * rgb[..., 2]
    ).astype(np.uint8)
    alpha = rgba[..., 3].copy()
    height, width = gray.shape
    return width, height, gray, alpha


def classify_gray(gray: np.ndarray, low: int, high: int) -> np.ndarray:
    labels = np.empty(gray.shape, dtype=np.uint8)
    labels[gray <= low] = 0
    middle = (gray > low) & (gray <= high)
    labels[middle] = 1
    labels[gray > high] = 2
    return labels


def _neighbor_counts(labels: np.ndarray):
    h, w = labels.shape
    padded = np.pad(labels, 1, mode="edge")
    c0 = np.zeros((h, w), dtype=np.uint8)
    c1 = np.zeros((h, w), dtype=np.uint8)
    c2 = np.zeros((h, w), dtype=np.uint8)
    for dy in range(3):
        for dx in range(3):
            view = padded[dy : dy + h, dx : dx + w]
            c0 += (view == 0)
            c1 += (view == 1)
            c2 += (view == 2)
    return c0, c1, c2


def apply_trinary_denoise(labels: np.ndarray, method: str) -> np.ndarray:
    if method == "none":
        return labels

    c0, c1, c2 = _neighbor_counts(labels)

    if method == "majority3":
        out = labels.copy()
        out[c0 >= 5] = 0
        out[c1 >= 5] = 1
        out[c2 >= 5] = 2
        return out

    if method == "mode3":
        out = np.zeros_like(labels)
        best = c0.copy()
        mask = c1 > best
        out[mask] = 1
        best[mask] = c1[mask]
        mask = c2 > best
        out[mask] = 2
        return out

    return labels


def _max_filter_labels(labels: np.ndarray) -> np.ndarray:
    h, w = labels.shape
    padded = np.pad(labels, 1, mode="edge")
    out = np.zeros_like(labels)
    for dy in range(3):
        for dx in range(3):
            np.maximum(out, padded[dy : dy + h, dx : dx + w], out=out)
    return out


def _min_filter_labels(labels: np.ndarray) -> np.ndarray:
    h, w = labels.shape
    padded = np.pad(labels, 1, mode="edge")
    out = np.full_like(labels, 2)
    for dy in range(3):
        for dx in range(3):
            np.minimum(out, padded[dy : dy + h, dx : dx + w], out=out)
    return out


def apply_morphology(
    labels: np.ndarray, method: str, iterations: int = 1
) -> np.ndarray:
    result = labels
    iterations = max(1, min(10, int(iterations)))
    for _ in range(iterations):
        if method == "dilate":
            result = _max_filter_labels(result)
        elif method == "erode":
            result = _min_filter_labels(result)
        elif method == "open":
            result = _max_filter_labels(_min_filter_labels(result))
        elif method == "close":
            result = _min_filter_labels(_max_filter_labels(result))
        else:
            break
    return result


def process_image(gray: np.ndarray, settings: ProcessingSettings) -> np.ndarray:
    settings = settings.normalized()
    labels = classify_gray(gray, settings.low, settings.high)
    labels = apply_trinary_denoise(labels, settings.denoise)
    if settings.morphology != "none":
        labels = apply_morphology(labels, settings.morphology, settings.iterations)
    return labels


def display_labels(labels: np.ndarray, invert: bool) -> np.ndarray:
    if not invert:
        return labels
    return (2 - labels).astype(np.uint8, copy=False)


def count_labels(labels: np.ndarray, invert: bool = False):
    shown = display_labels(labels, invert)
    black = int(np.count_nonzero(shown == 0))
    gray = int(np.count_nonzero(shown == 1))
    white = int(np.count_nonzero(shown == 2))
    return {
        "black": black,
        "gray": gray,
        "white": white,
        "total": int(shown.size),
    }


def compute_two_threshold_otsu(gray: np.ndarray):
    hist = np.bincount(gray.ravel(), minlength=256).astype(np.float64)
    total = hist.sum()
    if total <= 0:
        return 80, 180

    prob = hist / total
    omega = np.cumsum(prob)
    mu = np.cumsum(prob * np.arange(256, dtype=np.float64))
    mu_t = mu[-1]

    best_t1, best_t2 = 80, 180
    best_sigma = -1.0

    for t1 in range(0, 254):
        w0 = omega[t1]
        if w0 <= 0:
            continue
        mu0 = mu[t1] / w0
        for t2 in range(t1 + 1, 255):
            w1 = omega[t2] - omega[t1]
            w2 = 1.0 - omega[t2]
            if w1 <= 0 or w2 <= 0:
                continue
            mu1 = (mu[t2] - mu[t1]) / w1
            mu2 = (mu[-1] - mu[t2]) / w2
            sigma_b = (
                w0 * (mu0 - mu_t) ** 2
                + w1 * (mu1 - mu_t) ** 2
                + w2 * (mu2 - mu_t) ** 2
            )
            if sigma_b > best_sigma:
                best_sigma = sigma_b
                best_t1, best_t2 = t1, t2

    return int(best_t1), int(best_t2)


def _sample_nearest(array: np.ndarray, max_side: int | None):
    if not max_side:
        return array
    h, w = array.shape[:2]
    if max(h, w) <= max_side:
        return array
    scale = max_side / max(h, w)
    new_w = max(1, int(round(w * scale)))
    new_h = max(1, int(round(h * scale)))
    y_idx = np.linspace(0, h - 1, new_h).astype(np.intp)
    x_idx = np.linspace(0, w - 1, new_w).astype(np.intp)
    return array[np.ix_(y_idx, x_idx)]


def grayscale_preview_image(
    gray: np.ndarray, alpha: np.ndarray, max_side: int = 1200
) -> Image.Image:
    g = _sample_nearest(gray, max_side)
    a = _sample_nearest(alpha, max_side)
    rgba = np.empty((*g.shape, 4), dtype=np.uint8)
    rgba[..., 0] = g
    rgba[..., 1] = g
    rgba[..., 2] = g
    rgba[..., 3] = a
    return Image.fromarray(rgba, "RGBA")


def result_image(
    labels: np.ndarray,
    alpha: np.ndarray,
    settings: ProcessingSettings,
    max_side: int | None = None,
) -> Image.Image:
    settings = settings.normalized()
    shown = _sample_nearest(display_labels(labels, settings.invert), max_side)
    a = _sample_nearest(alpha, max_side)

    if settings.colorized:
        palette = np.array(
            [settings.black_color, settings.gray_color, settings.white_color],
            dtype=np.uint8,
        )
    else:
        palette = np.array(
            [(0, 0, 0), (settings.gray_level,) * 3, (255, 255, 255)],
            dtype=np.uint8,
        )

    rgb = palette[shown]
    rgba = np.empty((*shown.shape, 4), dtype=np.uint8)
    rgba[..., :3] = rgb
    rgba[..., 3] = a
    return Image.fromarray(rgba, "RGBA")


def save_result_image(
    path: str | Path,
    labels: np.ndarray,
    alpha: np.ndarray,
    settings: ProcessingSettings,
    fmt: str,
):
    out = result_image(labels, alpha, settings)
    fmt = fmt.lower()
    if fmt in {"jpeg", "jpg"}:
        background = Image.new("RGB", out.size, "white")
        background.paste(out, mask=out.getchannel("A"))
        background.save(path, "JPEG", quality=92)
    else:
        out.save(path, "PNG")
