from docx import Document
import re
from difflib import SequenceMatcher

# ===== CONFIGURE AQUI =====
DOC_A = "Mobius.docx"
DOC_B = "suiboM.docx"
# ===========================


def normalizar_numero(numero):
    """
    Normaliza número:
    1   -> 001
    12  -> 012
    100 -> 100
    403-2 -> 403-2 (mantém)
    """
    if "-" in numero:
        return numero.strip()
    return numero.zfill(3)


def extrair_blocos_docx(doc_path):
    doc = Document(doc_path)
    blocos = {}

    for p in doc.paragraphs:
        texto = p.text.strip()

        if not texto:
            continue

        # Procura número no final do parágrafo
        match = re.search(r"\(?(\d+(?:-\d+)?)\)?$", texto)

        if match:
            numero = normalizar_numero(match.group(1))

            # Remove o número do final
            conteudo = re.sub(r"\(?\d+(?:-\d+)?\)?$", "", texto).strip()

            # Normaliza espaços mas não pontuação
            conteudo = re.sub(r"\s+", " ", conteudo)

            blocos[numero] = conteudo

    return blocos


def mostrar_diferenca(a, b):
    matcher = SequenceMatcher(None, a, b)
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag != "equal":
            print("\n--- DOC A ---")
            print(a[i1:i2])
            print("\n--- DOC B ---")
            print(b[j1:j2])
            break


def comparar_blocos(blocos_a, blocos_b):
    divergencias = 0

    todos_numeros = sorted(
        set(blocos_a.keys()).union(set(blocos_b.keys()))
    )

    for numero in todos_numeros:

        if numero not in blocos_a:
            print(f"Bloco {numero} — AUSENTE no DOC A")
            divergencias += 1
            continue

        if numero not in blocos_b:
            print(f"Bloco {numero} — AUSENTE no DOC B")
            divergencias += 1
            continue

        texto_a = blocos_a[numero]
        texto_b = blocos_b[numero]

        if texto_a == texto_b:
            print(f"Bloco {numero} — OK")
        else:
            divergencias += 1
            print(f"\n🚨 Divergência detectada no bloco {numero}")
            mostrar_diferenca(texto_a, texto_b)

    print("\n==========================")
    if divergencias == 0:
        print("RESULTADO FINAL: IDENTIDADE TEXTUAL ABSOLUTA.")
    else:
        print(f"RESULTADO FINAL: {divergencias} divergência(s) encontrada(s).")
    print("==========================\n")


if __name__ == "__main__":
    print("Extraindo blocos do DOC A...")
    blocos_a = extrair_blocos_docx(DOC_A)

    print("Extraindo blocos do DOC B...")
    blocos_b = extrair_blocos_docx(DOC_B)

    print("\nIniciando comparação...\n")
    comparar_blocos(blocos_a, blocos_b)
