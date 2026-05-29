# HEADERS = {"User-Agent": "CodespaceMapStitcher/1.0 (student@example.com)"}

# TILE_SERVER_URL = "https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}&apistyle=s.e%3Al%7Cp.v%3Aoff"

import math
import requests
from PIL import Image
from io import BytesIO
from pathlib import Path

# ─── CONFIG ────────────────────────────────────────────
START_LAT = 19.127188
START_LON = 72.866657
END_LAT   = 19.041070
END_LON   = 73.015568
ZOOM      = 17

OUTPUT_DIR = "map_parts"
TILE_SIZE = 256
NUM_PARTS = 1

HEADERS = {"User-Agent": "ColabMapStitcher/1.0 (student@gmail.com)"}

# More Google-like than light_nolabels
# TILE_SERVER_URL = "https://a.basemaps.cartocdn.com/rastertiles/voyager_nolabels/{z}/{x}/{y}.png"
#TILE_SERVER_URL = "https://a.basemaps.cartocdn.com/light_nolabels/{z}/{x}/{y}.png"
TILE_SERVER_URL = "https://mt1.google.com/vt/lyrs=m&x={x}&y={y}&z={z}&apistyle=s.e%3Al%7Cp.v%3Aoff"


def lat_lon_to_tile(lat, lon, zoom):
    n = 2 ** zoom
    x = int((lon + 180.0) / 360.0 * n)
    y = int(
        (1.0 - math.log(
            math.tan(math.radians(lat)) +
            1.0 / math.cos(math.radians(lat))
        ) / math.pi) / 2.0 * n
    )
    return x, y


def download_tile(z, x, y):
    url = TILE_SERVER_URL.format(z=z, x=x, y=y)
    resp = requests.get(url, headers=HEADERS, timeout=15)
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "")
    if "image" not in content_type:
        raise ValueError(f"Non-image response: {content_type}")

    return Image.open(BytesIO(resp.content)).convert("RGB")


def stitch_map(lat1, lon1, lat2, lon2, zoom):
    min_lat = min(lat1, lat2)
    max_lat = max(lat1, lat2)
    min_lon = min(lon1, lon2)
    max_lon = max(lon1, lon2)

    x_min_tile, y_max_tile = lat_lon_to_tile(min_lat, min_lon, zoom)
    x_max_tile, y_min_tile = lat_lon_to_tile(max_lat, max_lon, zoom)

    x_tiles = range(x_min_tile, x_max_tile + 1)
    y_tiles = range(y_min_tile, y_max_tile + 1)

    cols, rows = len(x_tiles), len(y_tiles)
    canvas = Image.new("RGB", (cols * TILE_SIZE, rows * TILE_SIZE), "white")

    total = cols * rows
    print(f"Downloading {total} tiles ({cols}x{rows} grid) at zoom {zoom}...")

    for j, ty in enumerate(y_tiles):
        for i, tx in enumerate(x_tiles):
            done = j * cols + i + 1
            print(f"  [{done}/{total}] tile ({tx}, {ty})", end="\r")
            try:
                tile = download_tile(zoom, tx, ty)
                canvas.paste(tile, (i * TILE_SIZE, j * TILE_SIZE))
            except Exception as e:
                print(f"\nSkipped tile ({tx},{ty}): {e}")

    print(f"\nDone! Canvas size: {canvas.width} x {canvas.height} px")
    return canvas


def generate_parts(start_lat, start_lon, end_lat, end_lon, zoom, num_parts=5):
    output_dir = Path(OUTPUT_DIR)
    output_dir.mkdir(parents=True, exist_ok=True)

    top_lat = max(start_lat, end_lat)
    bottom_lat = min(start_lat, end_lat)
    left_lon = min(start_lon, end_lon)
    right_lon = max(start_lon, end_lon)

    part_height = (top_lat - bottom_lat) / num_parts

    for i in range(num_parts):
        part_top = top_lat - i * part_height
        part_bottom = top_lat - (i + 1) * part_height

        print(f"\n=== Part {i+1}/{num_parts} ===")
        print(f"Top-left:     ({part_top}, {left_lon})")
        print(f"Bottom-right: ({part_bottom}, {right_lon})")

        img = stitch_map(part_top, left_lon, part_bottom, right_lon, zoom)

        output_file = output_dir / f"map_part_{i+1}.png"
        img.save(output_file)
        print(f"Saved image to: {output_file.resolve()}")


def main():
    generate_parts(START_LAT, START_LON, END_LAT, END_LON, ZOOM, NUM_PARTS)
    print("\nAll parts saved successfully.")


if __name__ == "__main__":
    main()