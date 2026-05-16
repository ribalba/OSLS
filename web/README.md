# OSLS Web

This is the Django web version of the Open Source Labeling Scale Project.

It stores cuts, session data, label layout, scale state, hardware settings, and print analytics in a local SQLite database. Static CSS, fonts, and the default logo are kept inside this folder so the app can run without internet access after installation.

## Setup

```bash
cd web
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python manage.py migrate
python manage.py import_legacy_data
```

`import_legacy_data` reads the existing project JSON files from `../config/` and imports them into SQLite. Use `--force` to replace existing imported data.

## Run

```bash
cd web
source .venv/bin/activate
python manage.py runserver 0.0.0.0:8000
```

Open `http://<raspberry-pi-ip>:8000/` from a browser on the same network.

The scale monitor starts with the Django process. For maintenance commands or testing without the scale attached, set:

```bash
OSLS_DISABLE_SCALE_MONITOR=1 python manage.py runserver 0.0.0.0:8000
```

## Runtime Data

- SQLite database: `web/data/osls.sqlite3`
- Printed label PNGs: `web/data/printed_labels/`
- Archived analytics groups: stored in the database
- Uploaded logos: `web/media/logos/`

## Hardware

The settings page stores:

- Brother printer USB identifier, for example `usb://0x04f9:0x209c`
- Scale serial port, for example `/dev/ttyUSB0`
- Scale baudrate, usually `9600`

If the printer reports a busy USB resource on Linux, `ipp-usb` may be claiming the printer. A common fix is:

```bash
sudo systemctl stop ipp-usb
```

For a Raspberry Pi service, run Django with a single process so only one scale monitor talks to the serial device.
