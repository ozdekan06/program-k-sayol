import json
import os
import shutil
import sys
from pathlib import Path


def prepare_runtime():
    app_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    bundle_dir = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    data_dir = app_dir / "data"
    static_dir = app_dir / "static"
    data_dir.mkdir(parents=True, exist_ok=True)
    if (bundle_dir / "static").exists():
        shutil.copytree(bundle_dir / "static", static_dir, dirs_exist_ok=True)
    config_path = app_dir / "config.json"
    if not config_path.exists():
        shutil.copy2(bundle_dir / "config.example.json", config_path)
    config = json.loads(config_path.read_text(encoding="utf-8"))
    mappings = {
        "modbus_port": "MODBUS_PORT", "modbus_baudrate": "MODBUS_BAUDRATE",
        "modbus_parity": "MODBUS_PARITY", "modbus_stopbits": "MODBUS_STOPBITS",
        "modbus_timeout": "MODBUS_TIMEOUT", "modbus_retries": "MODBUS_RETRIES",
        "inter_request_delay": "MODBUS_INTER_REQUEST_DELAY",
        "inter_device_delay": "MODBUS_INTER_DEVICE_DELAY",
    }
    for key, env_key in mappings.items():
        if key in config:
            os.environ[env_key] = str(config[key])
    os.environ["OZDEKAN_DATA_DIR"] = str(data_dir)
    os.environ["OZDEKAN_STATIC_DIR"] = str(static_dir)
    os.chdir(app_dir)
    return config


def main():
    config = prepare_runtime()
    import uvicorn
    from main import app
    uvicorn.run(app, host=str(config.get("host", "127.0.0.1")), port=int(config.get("port", 8000)), reload=False)


if __name__ == "__main__":
    main()
