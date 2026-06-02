import argparse
import json
from pathlib import Path

from PIL import Image


def main():
    parser = argparse.ArgumentParser(description="Crop named regions from a reference screenshot.")
    parser.add_argument("--source", required=True, help="Source image path.")
    parser.add_argument("--regions", required=True, help="JSON file with {name: [left, top, right, bottom]} pixel boxes.")
    parser.add_argument("--output-dir", required=True, help="Directory for cropped PNG files.")
    args = parser.parse_args()

    source = Path(args.source)
    output_dir = Path(args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    image = Image.open(source)
    regions = json.loads(Path(args.regions).read_text(encoding="utf-8"))

    outputs = {}
    for name, box in regions.items():
        if len(box) != 4:
            raise ValueError(f"Region {name} must have four coordinates")
        crop = image.crop(tuple(box))
        out_path = output_dir / f"{name}.png"
        crop.save(out_path)
        outputs[name] = str(out_path)

    print(json.dumps(outputs, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
