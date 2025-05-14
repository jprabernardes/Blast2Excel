import os
import re
import pandas as pd

# Get the directory where the script is located
script_dir = os.path.dirname(os.path.abspath(__file__))

# Define input and output folder paths
input_folder = os.path.join(script_dir, "input_folder")
output_folder = os.path.join(script_dir, "output_folder")
output_file = os.path.join(output_folder, "genbank_parsed_all.xlsx")

# Check and create folders with user feedback
if not os.path.exists(input_folder):
    os.makedirs(input_folder)
    print(f"📁 'input_folder' was not found and has been created at: {input_folder}")

if not os.path.exists(output_folder):
    os.makedirs(output_folder)
    print(f"📁 'output_folder' was not found and has been created at: {output_folder}")

all_entries = []  # List to collect parsed entries

# Iterate over all files in the input folder
for filename in os.listdir(input_folder):
    if not filename.endswith((".gb", ".gbk", ".fasta", ".txt")):
        continue  # Skip unsupported files

    filepath = os.path.join(input_folder, filename)
    with open(filepath, "r") as f:
        lines = f.readlines()

    i = 0
    entry = {}

    # Parse each line
    while i < len(lines):
        line = lines[i].strip()

        if line.startswith("LOCUS"):
            if entry:
                all_entries.append(entry)
            entry = {
                "File": filename,
                "Accession": "",
                "Definition": "",
                "Organism": "",
                "Host": "",
                "Geo_location": "",
                "Isolation_source": "",
                "Features": []
            }

        elif line.startswith("ACCESSION"):
            entry["Accession"] = line.split()[1]

        elif line.startswith("DEFINITION"):
            def_lines = [line.replace("DEFINITION", "").strip()]
            i += 1
            while i < len(lines) and not lines[i].startswith("ACCESSION"):
                def_lines.append(lines[i].strip())
                i += 1
            entry["Definition"] = " ".join(def_lines)
            i -= 1

        elif line.strip().startswith("ORGANISM"):
            org_lines = [line.strip().replace("ORGANISM", "").strip()]
            i += 1
            while i < len(lines) and lines[i].startswith(" "):
                org_lines.append(lines[i].strip())
                i += 1
            entry["Organism"] = " ".join(org_lines)
            i -= 1

        elif "/host=" in line:
            entry["Host"] = line.split('"')[1]
        elif "/geo_loc_name=" in line:
            entry["Geo_location"] = line.split('"')[1]
        elif "/isolation_source=" in line:
            entry["Isolation_source"] = line.split('"')[1]

        elif re.match(r"^\s{5}\w+", line):
            feature_type = line.strip().split()[0]
            location = " ".join(line.strip().split()[1:])
            i += 1
            products = []
            while i < len(lines) and lines[i].strip().startswith("/product="):
                products.append(lines[i].strip().split('"')[1])
                i += 1
            entry["Features"].append(f"{feature_type} ({location}) → {', '.join(products)}")
            i -= 1

        i += 1

    if entry:
        all_entries.append(entry)

# Save results to Excel
df = pd.DataFrame(all_entries)
df["Features"] = df["Features"].apply(lambda x: "\n".join(x))
df.to_excel(output_file, index=False)

print(f"✅ Spreadsheet saved to: {output_file}")
