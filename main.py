from scripts.combine_pdf import combine
from scripts.pptx_to_pdf import convert
from pathlib import Path

p = Path('.')

def main():
    convert(str(p / "in"))
    combine(str(p / "in"))

if __name__ == "__main__":
    main()