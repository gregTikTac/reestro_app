from pathlib import Path
import sys

sys.path.append(r"d:\Project_freelance\parser\reestro")
from reestro_parser import InputRow, generate_pdf
from fixtures_sample import SAMPLE_INFO, SAMPLE_RIGHTS, SAMPLE_FIAS

row = InputRow()
row.cadastral = SAMPLE_INFO["cadastralNumber"]
row.fias_guid = SAMPLE_FIAS

out = Path(r"d:\Project_freelance\parser\test_sample_5rights.pdf")
generate_pdf(SAMPLE_INFO, SAMPLE_RIGHTS, row, out)
print(f"PDF: {out} ({len(SAMPLE_RIGHTS)} прав, как в образце)")
