import io

from PIL import Image as PILImage
import PIL.ImageEnhance as IE
import PIL.ImageFilter as IF
import PIL.ImageChops as IC
import PIL.ImageOps as IO


class ImagePreprocessor:

    @staticmethod
    def normalize_and_enhance(img):
        if img.mode == "RGBA":
            img = img.convert("RGB")

        img = img.convert("L")
        img = img.filter(IF.MedianFilter(size=3))

        bg = img.filter(IF.GaussianBlur(radius=25))
        img = IC.subtract(img, bg, scale=1.0, offset=128)

        img = IO.autocontrast(img, cutoff=2)
        img = IE.Contrast(img).enhance(1.8)
        img = img.filter(IF.SHARPEN)

        try:
            import numpy as np
            arr = np.array(img).astype(np.float32)
            hist, _ = np.histogram(arr.flatten(), bins=256, range=(0, 256))
            total = arr.size
            sum_total = np.sum(np.arange(256) * hist)
            sum_bg = 0.0
            weight_bg = 0
            max_var = 0.0
            threshold = 128
            for t in range(256):
                weight_bg += hist[t]
                if weight_bg == 0:
                    continue
                weight_fg = total - weight_bg
                if weight_fg == 0:
                    break
                sum_bg += t * hist[t]
                mean_bg = sum_bg / weight_bg
                mean_fg = (sum_total - sum_bg) / weight_fg
                var_between = weight_bg * weight_fg * (mean_bg - mean_fg) ** 2
                if var_between > max_var:
                    max_var = var_between
                    threshold = t
            k = 0.03
            arr = 255.0 / (1.0 + np.exp(-k * (arr - threshold)))
            arr = np.clip(arr, 0, 255).astype(np.uint8)
            img = PILImage.fromarray(arr)
        except ImportError:
            pass

        img = img.convert("RGB")
        return img

    @staticmethod
    def deskew_image(img):
        try:
            import numpy as np

            gray = img.convert("L") if img.mode != "L" else img
            arr = np.array(gray)
            inv = 255 - arr

            best_angle = 0
            best_score = 0
            for angle_10x in range(-50, 51, 5):
                angle = angle_10x / 10.0
                rotated = img.rotate(angle, expand=False, fillcolor=255)
                rot_arr = 255 - np.array(rotated.convert("L"))
                projection = np.sum(rot_arr, axis=1)
                score = np.var(projection)
                if score > best_score:
                    best_score = score
                    best_angle = angle

            if abs(best_angle) > 0.1:
                img = img.rotate(best_angle, expand=True, fillcolor=255)
            return img
        except ImportError:
            return img

    @staticmethod
    def preprocess_image(image_path: str) -> bytes:
        img = PILImage.open(image_path)

        if img.mode == "RGBA":
            img = img.convert("RGB")

        img = ImagePreprocessor.deskew_image(img)

        max_dim = 4096
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)

        if max(img.size) < 3072:
            scale = 3072 / max(img.size)
            img = img.resize(
                (int(img.width * scale), int(img.height * scale)),
                PILImage.LANCZOS,
            )

        img = ImagePreprocessor.normalize_and_enhance(img)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()

    @staticmethod
    def preprocess_strip(img) -> bytes:
        if img.mode == "RGBA":
            img = img.convert("RGB")

        max_dim = 4096
        if max(img.size) > max_dim:
            img.thumbnail((max_dim, max_dim), PILImage.LANCZOS)

        if max(img.size) < 3072:
            scale = 3072 / max(img.size)
            img = img.resize(
                (int(img.width * scale), int(img.height * scale)),
                PILImage.LANCZOS,
            )

        img = ImagePreprocessor.normalize_and_enhance(img)

        buf = io.BytesIO()
        img.save(buf, format="PNG")
        return buf.getvalue()
