import os
from pathlib import Path
import pandas as pd
import streamlit as st
from exif import Image
import simplekml


def decimal_coords(coords, ref):
    decimal_degrees = coords[0] + coords[1] / 60 + coords[2] / 3600
    if ref in ['S', 'W']:
        decimal_degrees = -decimal_degrees
    return decimal_degrees


def get_image_folders(base_path="."):
    image_folders = []
    valid_ext = ('.jpg', '.jpeg', '.png')

    for root, dirs, files in os.walk(base_path):
        if any(file.lower().endswith(valid_ext) for file in files):
            image_folders.append(root)

    return sorted(set(image_folders))


def scan_images(image_folder):
    results = []
    skipped = []
    valid_ext = ('.jpg', '.jpeg', '.png')

    if not os.path.isdir(image_folder):
        return results, skipped, f"Folder '{image_folder}' not found."

    for filename in os.listdir(image_folder):
        if filename.lower().endswith(valid_ext):
            img_path = os.path.join(image_folder, filename)

            try:
                with open(img_path, 'rb') as img_file:
                    img = Image(img_file)

                    if img.has_exif and hasattr(img, 'gps_latitude') and hasattr(img, 'gps_longitude'):
                        lat = decimal_coords(img.gps_latitude, img.gps_latitude_ref)
                        lon = decimal_coords(img.gps_longitude, img.gps_longitude_ref)

                        results.append({
                            "imgname": filename,
                            "filepath": img_path,
                            "lat": lat,
                            "long": lon
                        })
                    else:
                        skipped.append({
                            "imgname": filename,
                            "reason": "No EXIF GPS data found"
                        })

            except Exception as e:
                skipped.append({
                    "imgname": filename,
                    "reason": str(e)
                })

    return results, skipped, None


def create_geotagged_kmz(image_folder, output_filename="Mapped_Images.kmz"):
    kml = simplekml.Kml()
    scanned_data, skipped_data, error = scan_images(image_folder)

    if error:
        return None, None, None, skipped_data, error

    # Create output folder inside "KMZ img embed"
    base_dir = Path(__file__).resolve().parent
    output_dir = base_dir / "output"
    output_dir.mkdir(parents=True, exist_ok=True)

    output_kmz_path = output_dir / output_filename
    output_csv_path = output_dir / "image_data_report.csv"

    data_for_csv = []
    mark_counter = 1

    for item in scanned_data:
        filename = item["imgname"]
        lat = item["lat"]
        lon = item["long"]

        pnt = kml.newpoint(name=f'P{mark_counter}', coords=[(lon, lat)])
        pnt.description = f"""
        <div>
            <p><b>Marker:</b> P{mark_counter}</p>
            <p><b>Image:</b> {filename}</p>
            <img src="{filename}" width="600"/>
        </div>
        """

        data_for_csv.append({
            'imgname': filename,
            'lat': lat,
            'long': lon,
            'mark': f'P{mark_counter}'
        })
        mark_counter += 1

    kml.savekmz(str(output_kmz_path))

    df = pd.DataFrame(data_for_csv)
    df.to_csv(output_csv_path, index=False)

    return str(output_kmz_path), str(output_csv_path), df, skipped_data, None


st.set_page_config(page_title="KMZ Generator", layout="wide")

st.title("KMZ Generator from Geotagged Images")
st.write("Select the folder containing geotagged images. All generated output will be saved inside the 'output' folder in 'KMZ img embed'.")

st.subheader("Locate Folder Containing Images")

available_folders = get_image_folders(".")
default_folder = "./input_folder"

if available_folders:
    try:
        default_index = available_folders.index("./input_folder") if "./input_folder" in available_folders else 0
    except:
        default_index = 0

    selected_folder = st.selectbox(
        "Available image folders in workspace",
        options=available_folders,
        index=default_index
    )
else:
    selected_folder = default_folder
    st.warning("No image folders were auto-detected in the workspace.")

folder_path = st.text_input(
    "Enter folder path containing images",
    value=selected_folder
)

output_kmz_name = st.text_input(
    "Enter output KMZ name",
    value="Mapped_Images.kmz"
)

if st.button("Generate KMZ"):
    with st.spinner("Processing images and generating KMZ..."):
        kmz_file, csv_file, df, skipped_data, error = create_geotagged_kmz(
            folder_path,
            output_kmz_name
        )

    if error:
        st.error(error)
    else:
        st.success("KMZ generated successfully.")

        st.write(f"KMZ saved to: `{kmz_file}`")
        st.write(f"CSV saved to: `{csv_file}`")

        if df is not None and not df.empty:
            st.subheader("Processed Image Details")
            st.dataframe(df, use_container_width=True)
        else:
            st.warning("No geotagged image data found.")

        if skipped_data:
            st.subheader("Skipped Files")
            skipped_df = pd.DataFrame(skipped_data)
            st.dataframe(skipped_df, use_container_width=True)

        if os.path.exists(kmz_file):
            with open(kmz_file, "rb") as f:
                st.download_button(
                    label="Download KMZ",
                    data=f,
                    file_name=os.path.basename(kmz_file),
                    mime="application/vnd.google-earth.kmz"
                )

        if os.path.exists(csv_file):
            with open(csv_file, "rb") as f:
                st.download_button(
                    label="Download CSV Report",
                    data=f,
                    file_name=os.path.basename(csv_file),
                    mime="text/csv"
                )

        st.info("All output files are stored in the 'output' folder inside 'KMZ img embed'.")