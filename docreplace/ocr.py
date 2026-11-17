import io
import os
from .model import Unit


MAX_PIXELS = 40_000_000


def load_image(data, mode='RGB'):
    from PIL import Image, ImageOps
    with Image.open(io.BytesIO(data)) as source:
        if source.width * source.height > MAX_PIXELS:
            raise ValueError('图片超过 4000 万像素，请单独处理')
        if getattr(source, 'n_frames', 1) > 1:
            raise ValueError('多帧图片暂不支持，避免丢失其他帧')
        return ImageOps.exif_transpose(source).convert(mode)


class OCR:
    def __init__(self):
        try:
            from rapidocr_onnxruntime import RapidOCR
            from PIL import Image
            import onnxruntime
        except ImportError as exc:
            raise RuntimeError('未安装图片识别组件，请使用含 OCR 的便携包') from exc
        onnxruntime.disable_telemetry_events()
        self.engine = RapidOCR()

    def scan(self, key, data, matcher):
        import numpy as np
        from PIL import Image, ImageOps
        im = load_image(data)
        result, _ = self.engine(np.array(im)[:, :, ::-1].copy())
        units = []
        for i, (box, text, confidence) in enumerate(result or []):
            hits = matcher.matches(text)
            if hits:
                units.append(Unit(key+':ocr:'+str(i), key+' / 图片文字（整行区域）', text,
                                  matcher.replace(text), len(hits), 'image', False,
                                  {'part': key, 'box': box, 'confidence': float(confidence)}))
        return units


def render(data, units, font_path, foreground='#000000', background='#ffffff'):
    from PIL import Image, ImageOps, ImageDraw, ImageFont
    im = load_image(data, 'RGBA')
    draw = ImageDraw.Draw(im)
    for unit in units:
        if unit.payload.get('replacement'):
            replacement = load_image(unit.payload['replacement'], 'RGBA')
            im = replacement.resize(im.size)
            break
        if not font_path or not os.path.isfile(font_path):
            raise ValueError('图片文字替换需要指定支持中文的 TTF/TTC/OTF 字体文件')
        box = unit.payload['box']
        x0, y0 = max(0, int(min(p[0] for p in box))), max(0, int(min(p[1] for p in box)))
        x1, y1 = min(im.width, int(max(p[0] for p in box))+1), min(im.height, int(max(p[1] for p in box))+1)
        if x1 <= x0 or y1 <= y0:
            raise ValueError('图片选区无效')
        size = max(1, y1-y0-2)
        while size > 1:
            font = ImageFont.truetype(font_path, size)
            rect = draw.textbbox((0, 0), unit.after, font=font)
            if rect[2]-rect[0] <= x1-x0-2 and rect[3]-rect[1] <= y1-y0-2:
                break
            size -= 1
        font = ImageFont.truetype(font_path, size)
        rect = draw.textbbox((0, 0), unit.after, font=font)
        if rect[2]-rect[0] > x1-x0-2 or rect[3]-rect[1] > y1-y0-2:
            raise ValueError('替换文字无法放入选区，请扩大选区或缩短文字')
        draw.rectangle((x0, y0, x1-1, y1-1), fill=background)
        draw.text((x0+1-rect[0], y0+1-rect[1]), unit.after, font=font, fill=foreground)
    buf = io.BytesIO()
    ext = units[0].payload['part'].lower().rsplit('.', 1)[-1]
    fmt = {'jpg': 'JPEG', 'jpeg': 'JPEG', 'bmp': 'BMP', 'tif': 'TIFF', 'tiff': 'TIFF'}.get(ext, 'PNG')
    (im.convert('RGB') if fmt in ('JPEG', 'BMP') else im).save(buf, format=fmt)
    return buf.getvalue()
