"""OZDEKAN Windows paketinin giriş noktası."""

import json
import os
import shutil
import sys
from pathlib import Path


def application_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def bundled_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))


def prepare_runtime() -> tuple[Path, dict]:
    app_dir = application_dir()
    bundle = bundled_dir()
    data_dir = app_dir / "data"
    static_dir = app_dir / "static"
    data_dir.mkdir(parents=True, exist_ok=True)

    # One-file paket her çalışmada geçici klasöre açılır. Web dosyalarını ilk
    # çalıştırmada kalıcı uygulama klasörüne kopyala; upload'lar burada korunur.
    bundled_static = bundle / "static"
    if bundled_static.exists():
        shutil.copytree(bundled_static, static_dir, dirs_exist_ok=True)

    config_path = app_dir / "config.json"
    bundled_config = bundle / "config.example.json"
    if not config_path.exists() and bundled_config.exists():
        shutil.copy2(bundled_config, config_path)

    config = {}
    if config_path.exists():
        config = json.loads(config_path.read_text(encoding="utf-8"))

    mappings = {
        "modbus_port": "MODBUS_PORT",
        "modbus_baudrate": "MODBUS_BAUDRATE",
        "modbus_parity": "MODBUS_PARITY",
        "modbus_stopbits": "MODBUS_STOPBITS",
        "modbus_timeout": "MODBUS_TIMEOUT",
        "modbus_retries": "MODBUS_RETRIES",
        "inter_request_delay": "MODBUS_INTER_REQUEST_DELAY",
        "inter_device_delay": "MODBUS_INTER_DEVICE_DELAY",
    }
    for config_key, environment_key in mappings.items():
        if config_key in config:
            os.environ[environment_key] = str(config[config_key])

    os.environ["OZDEKAN_DATA_DIR"] = str(data_dir)
    os.environ["OZDEKAN_STATIC_DIR"] = str(static_dir)
    os.chdir(app_dir)
    return app_dir, config


def main() -> None:
    _, config = prepare_runtime()

    # Ortam değişkenleri connection_api import edilmeden önce hazırlanmalıdır.
    import uvicorn
    from main import app

    uvicorn.run(
        app,
        host=str(config.get("host", "127.0.0.1")),
        port=int(config.get("port", 8000)),
        reload=False,
        log_level=str(config.get("log_level", "info")),
    )


if __name__ == "__main__":
    main()
