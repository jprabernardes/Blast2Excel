import os
import re
import pandas as pd
from Bio import SeqIO

# File path
# Get the directory of the current script, put the file in the same directory as the script
script_dir = os.path.dirname(os.path.abspath(__file__))
file = os.path.join(script_dir, "example.gb") # Change this to your file name
excel_output = os.path.join(script_dir, "blast_hits.xlsx")

# Initialize variables
tb = []
sequence_blasted = None

# Function to initialize variables for a new OTU
def reset_variables():
    return {
        "accession": None,
        "definition": None,
        "evalue": None,
        "bitscore": None,
        "Identity": None
    }

# Initialize variables
vars = reset_variables()
with open(file, "r") as f:
    for line in f:
        line = line.strip()

        # Identify the start of a new OTU
        if line.startswith("LOCUS"):
            partes = line.split()
            sequence_blasted = partes[1] if len(partes) > 1 else "Unknown"
            # Reset variables for the new OTU
            vars = reset_variables()

        # Extract fields
        elif "/accession=" in line:
            vars["accession"] = line.split('"')[1]
        elif "/def=" in line:
            vars["definition"] = line.split('"', 1)[1].strip()
        elif "/E-value=" in line:
            vars["evalue"] = line.split('"')[1]
        elif "/bit-score=" in line:
            vars["bitscore"] = line.split('"')[1]
        elif "/identities=" in line:
            match = re.search(r"\(([\d\.]+)%\)", line)
            vars["Identity"] = float(match.group(1)) if match else None

            # Check if all required fields are present
            all_fields_present = all(vars[field] is not None for field in vars)
            
            if all_fields_present and sequence_blasted is not None:
                # Add to the result (after the last required field)
                tb.append({
                    "Sequence-blasted": sequence_blasted,
                    "Accession": vars["accession"],
                    "Identity (%)": vars["Identity"],
                    "E-value": vars["evalue"],
                    "Bit Score": vars["bitscore"],
                    "Definition": vars["definition"]
                })
                # Reset variables after adding to the table to prepare for the next hit
                vars = reset_variables()

# Export to Excel
df = pd.DataFrame(tb)
df.to_excel(excel_output, index=False)

print(f"Spreadsheet saved to: {excel_output}")