"""
Prints CD / DVD thermal labels to a user-selected physical printer.

The correct approach for thermal label printers:
- Do NOT use `window.print()` (WebView2 browser dialog), because it depends on
  a Windows default printer and typically shows NO dialog / fails if no printer
  is installed.
- Instead, send a high-resolution rendered label image (PIL/Pillow) directly to
  the user-selected physical printer driver.

Backend:
- Windows : `win32print` + `win32ui` (GDI) — native Windows driver.
- Primary : `win32print` for printer enumeration and printing.
- Fallback: QtPrintSupport (PySide6, Windows-only) as an alternative.
- macOS/Linux: printing is handled by the built-in WebKit window.print() on the
  front-end; this backend is only used on Windows.

All imports are "lazy" so the module can be loaded even when a library is not
installed (e.g., development testing on macOS without win32print).
"""
import io
import os
import sys
import tempfile


def _get_pil():
    """Lazily imports Pillow (PIL), raising a clear error if it is not installed."""
    try:
        from PIL import Image, ImageDraw, ImageFont
        return Image, ImageDraw, ImageFont
    except ImportError as e:
        raise RuntimeError(
            "Pillow (Pillow) is not installed. Install it with: pip install Pillow"
        ) from e


# ==============================================================================
# RENDER LABEL AS A HIGH-RESOLUTION IMAGE (PIL)
# ==============================================================================

def _load_font(size_px):
    """Loads an Arial/Bold font if available, otherwise falls back to the PIL default font."""
    _, _, ImageFont = _get_pil()
    candidates = []
    if sys.platform == 'win32':
        windir = os.environ.get('WINDIR', r'C:\Windows')
        candidates = [
            os.path.join(windir, 'Fonts', 'arialbd.ttf'),
            os.path.join(windir, 'Fonts', 'arial.ttf'),
        ]
    elif sys.platform == 'darwin':
        candidates = [
            '/System/Library/Fonts/Helvetica.ttc',
            '/Library/Fonts/Arial.ttf',
        ]
    else:
        candidates = [
            '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf',
            '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
        ]
    for path in candidates:
        if os.path.exists(path):
            try:
                return ImageFont.truetype(path, size_px)
            except Exception:
                continue
    return ImageFont.load_default()


def render_label_image(label_data, width_mm=50, height_mm=30, font_px=34):
    """
    Renders a label into a high-resolution RGB image (mm -> pixels at ~300 DPI).
    label_data: dict with keys:
        facility, date, ic, nama, xray, exam, circle (bool)
    The returned image is linearly sized to DPI 300.
    """
    Image, ImageDraw, _ = _get_pil()
    dpi = 300
    px_w = max(10, int(round(width_mm / 25.4 * dpi)))
    px_h = max(10, int(round(height_mm / 25.4 * dpi)))

    is_circle = bool(label_data.get("circle", False))

    if is_circle:
        # Safe default rectangular image; the circle is reproduced via the printer
        # driver's cut lines. Keep content within the central viewing area.
        canvas = Image.new("RGB", (px_w, px_h), "#ffffff")
        draw = ImageDraw.Draw(canvas)
        draw.rectangle([0, 0, px_w - 1, px_h - 1], outline="#000000", width=6)
        margin = int(px_w * 0.18)
        inner_box = [margin, margin, px_w - margin, px_h - margin]
        lines = _build_lines(label_data)
        _draw_wrapped(draw, inner_box, lines, px_w, px_h, font_px, center=True)
        return canvas

    canvas = Image.new("RGB", (px_w, px_h), "#ffffff")
    draw = ImageDraw.Draw(canvas)
    # Outer border
    draw.rectangle([0, 0, px_w - 1, px_h - 1], outline="#000000", width=6)

    margin = int(px_w * 0.06)
    box = [margin, margin, px_w - margin, px_h - margin]

    # Top group: facility + date + ID
    top_lines = []
    if label_data.get("facility") and label_data.get("show_facility", True):
        top_lines.append(("fw", label_data["facility"]))
    if label_data.get("date") and label_data.get("show_date", True):
        top_lines.append(("n", "DATE: " + str(label_data["date"])))
    if label_data.get("ic") and label_data.get("show_ic", True):
        top_lines.append(("n", "ID: " + str(label_data["ic"])))

    # Middle group: patient name
    nama = label_data.get("nama", "")
    if nama and label_data.get("show_nama", True):
        top_lines.append(("fw", str(nama).upper()))

    y = box[1]
    line_h = int(font_px * 1.35)
    for kind, text in top_lines:
        if not text:
            continue
        font = _load_font(font_px if kind == "n" else int(font_px * 1.15))
        y = _draw_line(draw, (box[0], box[2], y), text, font, bold=(kind == "fw"), underline=(kind == "fw"))
        y += line_h
        if y > box[3] - line_h:
            break

    # Separator line
    if box[3] - y > line_h * 2:
        draw.line([box[0], y, box[2], y], fill="#000000", width=4)

    # Bottom group: X-Ray + Exam
    bottom_x = box[0]
    if label_data.get("xray") and label_data.get("show_xray", True):
        draw.line([box[0], box[3] - line_h * 2, box[2], box[3] - line_h * 2], fill="#000000", width=4)
        draw.text((box[0], box[3] - line_h * 2 + 6), "NO. X-RAY: " + str(label_data["xray"]),
                  font=_load_font(int(font_px * 1.25)), fill="#000000")
    if label_data.get("exam") and label_data.get("show_exam", True):
        draw.text((box[0], box[3] - int(font_px * 1.15)), "EXAM: " + str(label_data["exam"]),
                  font=_load_font(font_px), fill="#000000")

    return canvas


def _build_lines(label_data):
    lines = []
    if label_data.get("facility") and label_data.get("show_facility", True):
        lines.append(("fw", str(label_data["facility"]).upper()))
    if label_data.get("date") and label_data.get("show_date", True):
        lines.append(("n", "DATE: " + str(label_data["date"])))
    if label_data.get("ic") and label_data.get("show_ic", True):
        lines.append(("n", "ID: " + str(label_data["ic"])))
    if label_data.get("nama") and label_data.get("show_nama", True):
        lines.append(("fw", str(label_data["nama"]).upper()))
    if label_data.get("xray") and label_data.get("show_xray", True):
        lines.append(("fw", "NO. X-RAY: " + str(label_data["xray"])))
    if label_data.get("exam") and label_data.get("show_exam", True):
        lines.append(("n", "EXAM: " + str(label_data["exam"])))
    return lines


def _draw_wrapped(draw, box, lines, px_w, px_h, font_px, center=False):
    y = box[1]
    available = box[2] - box[0]
    for kind, text in lines:
        if not text:
            continue
        size = font_px if kind == "n" else int(font_px * 1.1)
        font = _load_font(size)
        while text:
            # Wrap the text according to the available width
            split_at = None
            for i in range(len(text), 0, -1):
                if draw.textlength(text[:i], font=font) <= available:
                    split_at = i
                    break
            if split_at is None:
                split_at = 1
            chunk = text[:split_at]
            text = text[split_at:].strip()
            if center:
                w = draw.textlength(chunk, font=font)
                x = box[0] + (available - w) / 2
            else:
                x = box[0]
            draw.text((x, y), chunk, font=font, fill="#000000")
            y += int(size * 1.35)
            if y > box[3] - int(size * 1.4):
                return


def _draw_line(draw, box, text, font, bold=False, underline=False):
    x0, x1, y = box
    draw.text((x0, y), text, font=font, fill="#000000")
    if underline:
        w = draw.textlength(text, font=font)
        draw.line([x0, y + int(font.size * 1.15), x0 + w, y + int(font.size * 1.15)], fill="#000000", width=4)
    return y


# ==============================================================================
# PRINTER SYSTEM
# ==============================================================================

def list_printers():
    """
    Returns a list of available printer names.
    Windows : win32print.EnumPrinters
    Others  : QtPrintSupport (QPrinterInfo) if PySide6 is installed.
    """
    printers = []
    if sys.platform == 'win32':
        try:
            import win32print
            for p in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
                printers.append(p[2])
        except Exception as e:
            print(f"[LabelPrinter] EnumPrinters warning: {e}")

    if not printers:
        # Try Qt (PySide6 / PyQt5) for other platforms
        try:
            from PySide6.QtPrintSupport import QPrinterInfo
        except Exception:
            try:
                from PyQt5.QtPrintSupport import QPrinterInfo
            except Exception:
                QPrinterInfo = None
        if QPrinterInfo:
            try:
                printers = [name for name in QPrinterInfo.availablePrinterNames()]
            except Exception as e:
                print(f"[LabelPrinter] Qt printer enum warning: {e}")
    return printers


def get_default_printer():
    """Returns the default printer name, if any."""
    if sys.platform == 'win32':
        try:
            import win32print
            return win32print.GetDefaultPrinter()
        except Exception:
            return ""
    return (list_printers() or [""])[0]


def _print_win32(printer_name, image_bytes, width_mm, height_mm):
    """Prints a label image to a Windows printer via GDI (Windows driver)."""
    import win32print
    import win32ui
    from PIL import ImageWin

    with tempfile.NamedTemporaryFile(suffix=".bmp", delete=False) as tmp:
        tmp_path = tmp.name
    try:
        pil_image = Image.open(io.BytesIO(image_bytes))
        pil_image.load()
        # Save a temporary BMP (GDI handles BMP easily)
        pil_image.save(tmp_path, "BMP")

        hprinter = win32print.OpenPrinter(printer_name)
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)
            # Start the print document
            hdc.StartDoc("Daftar Radiologi CD Label")
            hdc.StartPage()

            dib = ImageWin.Dib(pil_image, "RGB")
            # Draw on the DC at the physical size (printer pixels)
            hdc_dc = hdc.GetHandleOutput()
            dib.draw(hdc_dc, (0, 0, pil_image.width, pil_image.height))

            hdc.EndPage()
            hdc.EndDoc()
            hdc.DeleteDC()
        finally:
            win32print.ClosePrinter(hprinter)
    finally:
        try:
            os.remove(tmp_path)
        except Exception:
            pass
    return True


def _print_qt(printer_name, image_bytes, width_mm, height_mm):
    """Prints a label image via QtPrintSupport (Windows)."""
    app = None
    try:
        from PySide6.QtWidgets import QApplication
        from PySide6.QtCore import QRectF, QSizeF
        from PySide6.QtGui import QImage, QPainter
        from PySide6.QtPrintSupport import QPrinter
        QS_BINDING = "pyside6"
    except Exception:
        try:
            from PyQt5.QtWidgets import QApplication
            from PyQt5.QtCore import QRectF, QSizeF
            from PyQt5.QtGui import QImage, QPainter
            from PyQt5.QtPrintSupport import QPrinter
            QS_BINDING = "pyqt5"
        except Exception as e:
            return False, "QtPrintSupport is not available (PySide6 / PyQt5)."

    if not QApplication.instance():
        app = QApplication(sys.argv)

    img = QImage()
    img.loadFromData(image_bytes)

    printer = QPrinter(QPrinter.HighResolution)
    if printer_name:
        printer.setPrinterName(printer_name)
    printer.setFullPage(True)
    try:
        printer.setPaperSize(QSizeF(width_mm, height_mm), QPrinter.Millimeter)
    except Exception:
        pass

    painter = QPainter(printer)
    try:
        page = printer.pageRect(QPrinter.DevicePixel)
        target = QRectF(0, 0, page.width(), page.height())
        src = QRectF(0, 0, img.width(), img.height())
        painter.drawImage(target, img, src)
    finally:
        painter.end()
    if app:
        app.processEvents()
    return True, "ok"


def print_label_image(printer_name, image_bytes, width_mm=50, height_mm=30):
    """
    Main entry point to print a label image.
    Returns a tuple (success: bool, message: str).
    """
    printers = list_printers()
    if printer_name:
        printer_name = str(printer_name).strip()
    else:
        printer_name = get_default_printer()

    if not printer_name:
        detail = ""
        if printers:
            detail = " Available printers: " + ", ".join(printers[:5]) + ("..." if len(printers) > 5 else "")
        return False, "No printer is installed or selected on Windows." + detail

    if sys.platform == 'win32':
        try:
            _print_win32(printer_name, image_bytes, width_mm, height_mm)
            return True, "Label sent to printer: " + printer_name
        except Exception as e:
            # Try Qt as a fallback (PySide6 installed only on Windows)
            try:
                ok, _ = _print_qt(printer_name, image_bytes, width_mm, height_mm)
                if ok:
                    return True, "Label printed (Qt) to: " + printer_name
                return False, "QtPrintSupport is not available."
            except Exception as e2:
                return False, f"Failed to print via win32print ({e}). Qt fallback also failed ({e2})."
    else:
        # macOS / Linux: use the built-in WebKit window.print() —
        # print_label_image is not needed here since printing is handled on the front-end.
        return False, "Not supported: use the WebKit (window.print()) path on this platform."
