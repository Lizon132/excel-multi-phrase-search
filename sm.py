import argparse
import sys
from typing import List, Sequence

import numpy as np
from sentence_transformers import SentenceTransformer, util

# ---- You can keep using your current model name
MODEL_NAME = "all-MiniLM-L6-v2"

try:
    import xlwings as xw
except ImportError:
    print("ERROR: xlwings not installed. Run: pip install xlwings", file=sys.stderr)
    sys.exit(2)


# ---------- helpers ----------
def build_fulltext(rows: Sequence[Sequence[object]]) -> List[str]:
    """
    Merge all cell values in a row into one string (like your df['FullText']).
    None/NaN become ''.
    """
    out: List[str] = []
    for cells in rows:
        parts = []
        for v in cells:
            if v is None:
                continue
            s = str(v)
            if s.strip():
                parts.append(s)
        out.append(" ".join(parts))
    return out


def scores_0_1(query: str, fulltexts: Sequence[str], model: SentenceTransformer) -> List[float]:
    """
    Cosine similarity(query, row) -> normalized to [0,1]
    """
    if not fulltexts:
        return []

    # Encode once
    row_emb = model.encode(fulltexts, convert_to_tensor=True, show_progress_bar=False)
    q_emb = model.encode(query, convert_to_tensor=True, show_progress_bar=False)

    cos = util.cos_sim(q_emb, row_emb)[0].detach().cpu().numpy()  # shape (n,)
    # util.cos_sim returns roughly [-1, 1]; clamp & normalize to [0,1]
    cos = np.clip(cos, -1.0, 1.0)
    sm = (cos + 1.0) / 2.0
    return sm.astype(float).tolist()


# ---------- main ----------
def main():
    ap = argparse.ArgumentParser(description="Compute semantic match values and write to Excel column A.")
    ap.add_argument("--query", required=True, help="search text")
    ap.add_argument("--workbook", required=True, help="path to .xlsm/.xlsx")
    ap.add_argument("--sheet", required=True, help="target sheet name (e.g., 'Test Docs')")
    ap.add_argument("--start-row", type=int, default=2, help="data starts here (header is row 1)")
    # If you later want to restrict columns, we can add a --use-columns flag (A1 letters or header names)
    args = ap.parse_args()

    app = None
    wb = None
    created_app = False

    try:
        # attach to running Excel if possible
        try:
            app = xw.apps.active
        except Exception:
            app = None
        if app is None:
            app = xw.App(visible=False, add_book=False)
            created_app = True

        # open or reuse workbook
        for b in app.books:
            if b.fullname.lower() == args.workbook.lower():
                wb = b
                break
        if wb is None:
            wb = app.books.open(args.workbook)

        sht = wb.sheets[args.sheet]

        # Find last row and last col (we read B..last for text, write scores to A)
        last_row = sht.range("B" + str(sht.cells.last_cell.row)).end("up").row
        last_col = sht.range((1, sht.cells.last_cell.column)).end("left").column
        if last_row < args.start_row or last_col < 2:
            # nothing to score
            return 0

        # Read rows B..last_col as a 2D list
        rng = sht.range((args.start_row, 2), (last_row, last_col))
        values = rng.value  # list of lists (or list if single col)
        if last_col == 2:  # only one text column
            values = [[v] for v in values]

        fulltexts = build_fulltext(values)

        # Load model (cached by process; subsequent runs are faster)
        model = SentenceTransformer(MODEL_NAME)

        # Compute scores
        scores = scores_0_1(args.query, fulltexts, model)

        # Write back to A2..A{last_row}
        out = scores[: (last_row - args.start_row + 1)]
        sht.range((args.start_row, 1)).options(transpose=True).value = out

        # Do not save here; VBA decides formatting/Save
        return 0

    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        return 1
    finally:
        if created_app and app is not None:
            app.quit()


if __name__ == "__main__":
    sys.exit(main())
