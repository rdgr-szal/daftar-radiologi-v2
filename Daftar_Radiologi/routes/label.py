import io
from flask import Blueprint, render_template, request, jsonify
from core.config import load_config, save_config
from core.label_printer import (
    list_printers,
    get_default_printer,
    render_label_image,
    print_label_image
)

label_bp = Blueprint('label', __name__)


@label_bp.route('/api/label/printers', methods=['GET'])
def api_list_label_printers():
    """Lists all physical printers available on the Windows system."""
    try:
        printers = list_printers()
        default_printer = get_default_printer()
        return jsonify({
            "success": True,
            "printers": printers,
            "default_printer": default_printer
        })
    except Exception as e:
        return jsonify({"success": False, "message": f"Error listing printers: {e}"}), 500


@label_bp.route('/api/label/config', methods=['GET'])
def api_get_label_config():
    """Returns the label configuration (size, printer, fields) from config.json."""
    config = load_config()
    label_config = config.get("label_config", {})
    return jsonify({
        "success": True,
        "label_config": label_config,
        "facility_name": config.get("klinik_asal", "KLINIK KESIHATAN"),
        "singkatan": config.get("singkatan_klinik", "")
    })


@label_bp.route('/api/label/save-config', methods=['POST'])
def api_save_label_config():
    """Saves label preferences (printer, size, fields) to config.json."""
    data = request.get_json() or {}
    config = load_config()

    label_config = config.get("label_config", {})

    if "printer_name" in data:
        label_config["printer_name"] = str(data.get("printer_name", "")).strip()

    for key in ("width_mm", "height_mm"):
        if key in data:
            try:
                label_config[key] = float(data.get(key) or 0)
            except (ValueError, TypeError):
                pass

    if "preset" in data:
        label_config["preset"] = str(data.get("preset", "RECT_50x30"))
    if "font_size" in data:
        label_config["font_size"] = str(data.get("font_size", "11px"))
    for key in ("show_facility", "show_nama", "show_ic", "show_exam", "show_date", "show_xray"):
        if key in data:
            label_config[key] = bool(data.get(key))

    config["label_config"] = label_config
    save_config(config)
    return jsonify({"success": True, "message": "Label configuration saved."})


@label_bp.route('/api/label/preview', methods=['POST'])
def api_preview_label():
    """
    Generates a PNG image of the label for web preview (from patient data).
    Body: { nama, ic_pasport, tarikh, xray, exam, width_mm, height_mm, font_px, circle }
    """
    from flask import send_file
    data = request.get_json(force=True) or {}
    width_mm = _to_float(data.get("width_mm"), 50)
    height_mm = _to_float(data.get("height_mm"), 30)
    font_px = _to_int(data.get("font_px"), 34)

    label_data = {
        "facility": data.get("facility", ""),
        "nama": data.get("nama", ""),
        "ic": data.get("ic_pasport", data.get("ic", "")),
        "tarikh": data.get("tarikh", ""),
        "xray": str(data.get("xray", "")),
        "exam": str(data.get("exam", "")),
        "circle": bool(data.get("circle", False)),
    }
    for key in ("show_facility", "show_nama", "show_ic", "show_exam", "show_date", "show_xray"):
        if key in data:
            label_data[key] = bool(data.get(key))
    img = render_label_image(label_data, width_mm=width_mm, height_mm=height_mm, font_px=font_px)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return send_file(buf, mimetype="image/png")


@label_bp.route('/api/label/print', methods=['POST'])
def api_print_label():
    """
    Prints a CD/DVD label to the user-selected physical printer.
    Body: {
        printer_name, nama, ic_pasport, tarikh, xray, exam,
        width_mm, height_mm, font_px, circle, facility
    }
    """
    data = request.get_json(force=True) or {}
    printer_name = str(data.get("printer_name", "")).strip()
    width_mm = _to_float(data.get("width_mm"), 50)
    height_mm = _to_float(data.get("height_mm"), 30)
    font_px = _to_int(data.get("font_px"), 34)

    if width_mm < 10 or height_mm < 10 or width_mm > 200 or height_mm > 200:
        return jsonify({"success": False, "message": "Invalid label size (10-200 mm)."}), 400

    config = load_config()
    label_config = config.get("label_config", {})

    if not printer_name:
        printer_name = label_config.get("printer_name", "")

    # Ensure the facility name value is correct
    facility = data.get("facility", "") or config.get("klinik_asal", "KLINIK KESIHATAN")

    label_data = {
        "facility": facility,
        "nama": data.get("nama", ""),
        "ic": data.get("ic_pasport", data.get("ic", "")),
        "tarikh": data.get("tarikh", ""),
        "xray": str(data.get("xray", "")),
        "exam": str(data.get("exam", "")),
        "circle": bool(data.get("circle", False)),
    }
    # Fields can be specified from the label UI (field display checkboxes)
    for key in ("show_facility", "show_nama", "show_ic", "show_exam", "show_date", "show_xray"):
        if key in data:
            label_data[key] = bool(data.get(key))

    try:
        img = render_label_image(label_data, width_mm=width_mm, height_mm=height_mm, font_px=font_px)
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        image_bytes = buf.getvalue()

        ok, msg = print_label_image(printer_name, image_bytes, width_mm=width_mm, height_mm=height_mm)

        # Save the last successfully used printer
        if ok and printer_name:
            label_config["printer_name"] = printer_name
            config["label_config"] = label_config
            save_config(config)

        return jsonify({"success": ok, "message": msg}), (200 if ok else 500)
    except Exception as e:
        return jsonify({"success": False, "message": f"Error printing label: {e}"}), 500


def _to_float(value, default):
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def _to_int(value, default):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default
