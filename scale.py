#!/usr/bin/env python3
import serial
import re

PORT = "/dev/ttyUSB0"
BAUDRATE = 9600
STABLE_ITERATIONS = 10

# Regex to extract a number before 'kg'
VALUE_RE = re.compile(r'([-+]?\d+(?:\.\d+)?)\s*kg')

def main():
    # Open serial port with same settings as your stty command
    ser = serial.Serial(
        PORT,
        BAUDRATE,
        bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE,
        stopbits=serial.STOPBITS_ONE,
        xonxoff=False,
        rtscts=False,
    )

    print(f"Reading from {PORT} at {BAUDRATE} baud. Press Ctrl+C to stop.\n")

    last_value = None
    stable_count = 0
    last_printed_value = None

    try:
        while True:
            raw = ser.readline()  # read one line (blocks until newline)
            try:
                line = raw.decode(errors="ignore").strip()
            except UnicodeDecodeError:
                # If decoding fails for some reason, just skip this line
                continue

            if not line:
                continue

            # Example line: "+ 0.0335kg"
            m = VALUE_RE.search(line)
            if m:
                value_str = m.group(1)
                value = float(value_str)
                if value == last_value:
                    stable_count += 1
                else:
                    last_value = value
                    stable_count = 1

                # Print only when the value has stayed unchanged for enough
                # consecutive readings, and print it once per stable value.
                if stable_count >= STABLE_ITERATIONS and value != last_printed_value:
                    print(f"{value:.4f} kg")
                    last_printed_value = value
            else:
                # If the line doesn't match, you can print it for debugging
                # or just ignore it.
                print(f"Unparsed line: {line}")

    except KeyboardInterrupt:
        print("\nStopping...")
    finally:
        ser.close()

if __name__ == "__main__":
    main()
