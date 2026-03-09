from docx import Document
import re
from difflib import SequenceMatcher

# ===== CONFIGURE HERE =====
DOC_A = "Mobius.docx"
DOC_B = "suiboM.docx"
# ===========================


def normalize_number(number):
    """
     Normalize number:
    1   -> 001
    12  -> 012
    100 -> 100
    403-2 -> 403-2 (preserved)
    """
    if "-" in number:
        return number.strip()
    return number.zfill(3)


def extract_blocks_docx(doc_path):
    doc = Document(doc_path)
    blocks = {}

    for p in doc.paragraphs:
        text = p.text.strip()

        if not text:
            continue

       # Look for a number at the end of the paragraph
        match = re.search(r"\(?(\d+(?:-\d+)?)\)?$", text)

        if match:
            number = normalize_number(match.group(1))

            # Remove the number from the end
            content = re.sub(r"\(?\d+(?:-\d+)?\)?$", "", text).strip()

            # Normalize spaces but not punctuation
            content = re.sub(r"\s+", " ", content)

            blocks[number] = content

    return blocks


def show_difference(a, b):
    matcher = SequenceMatcher(None, a, b)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            print("\n--- DOC A ---")
            print(a[i1:i2])
            print("\n--- DOC B ---")
            print(b[j1:j2])
            break


def compare_blocks(blocks_a, blocks_b):
    divergences = 0

    all_numbers = sorted(
        set(blocks_a.keys()).union(set(blocks_b.keys()))
    )

    for number in all_numbers:

        if number not in blocks_a:
            print(f"Block {number} — MISSING in DOC A")
            divergences += 1
            continue

        if number not in blocks_b:
            print(f"Block {number} — MISSING in DOC B")
            divergences += 1
            continue

        texto_a = blocks_a[number]
        texto_b = blocks_b[number]

        if texto_a == texto_b:
            print(f"Block {number} — OK")
        else:
            divergences += 1
            print(f"\n🚨 Divergence detected in block {number}")
            show_difference(texto_a, texto_b)

    print("\n==========================")
    if divergences == 0:
        print("FINAL RESULT: ABSOLUTE TEXTUAL IDENTITY.")
    else:
        print(f"FINAL RESULT: {divergences} divergence(s) found.")
    print("==========================\n")


if __name__ == "__main__":
    print("Extracting blocks from DOC A...")
    blocos_a = extract_blocks_docx(DOC_A)

    print("Extracting blocks from DOC B...")
    blocos_b = extract_blocks_docx(DOC_B)

    print("\nStarting comparison...\n")
    compare_blocks(blocos_a, blocos_b)
