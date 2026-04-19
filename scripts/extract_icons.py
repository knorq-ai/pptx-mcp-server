#!/usr/bin/env python3
"""
Extract icons from iconography_NEW.pptx into a single icons.json asset file.

Each icon is stored as:
  - id: slug-based identifier
  - name: human-readable name from slide title
  - category: auto-classified category
  - keywords: search terms
  - type: "sp" or "grpSp"
  - orig_width_emu / orig_height_emu: original dimensions
  - aspect_ratio: width / height
  - colors: hex colors found in the icon
  - xml: serialized XML string of the shape element

Usage:
    python scripts/extract_icons.py /path/to/iconography_NEW.pptx
"""

from __future__ import annotations

import json
import re
import sys
import zipfile
from lxml import etree

NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"

# Category classification by keyword
CATEGORY_KEYWORDS = {
    "people": [
        "person", "man", "woman", "people", "family", "child", "kid", "boy", "girl",
        "baby", "infant", "nurse", "doctor", "elderly", "father", "mother", "couple",
        "wedding", "team", "group", "businessman", "businesswoman", "secretary",
        "pregnant", "handicap", "toddler", "caregiver", "silhouette",
    ],
    "business": [
        "briefcase", "chart", "presentation", "profit", "money", "dollar", "credit",
        "bank", "piggy", "calculator", "calendar", "clipboard", "document", "folder",
        "org chart", "resume", "id", "badge", "license", "kiosk", "card", "safe",
        "handshake", "meeting", "interview", "gavel", "scales", "justice",
    ],
    "technology": [
        "computer", "laptop", "desktop", "monitor", "keyboard", "mouse", "phone",
        "smartphone", "tablet", "server", "cloud", "internet", "web", "browser",
        "email", "fax", "printer", "scanner", "usb", "software", "data", "database",
        "network", "router", "signal", "headset", "headphone", "microphone",
        "camera", "video", "film", "television", "tv", "cd", "floppy",
    ],
    "transport": [
        "car", "bus", "truck", "van", "train", "airplane", "bike", "bicycle",
        "boat", "sailboat", "ambulance", "rocket", "wheel",
    ],
    "medical": [
        "medical", "hospital", "hospice", "heartbeat", "stethoscope", "syringe",
        "pill", "medication", "thermometer", "tooth", "crutch", "diaper",
    ],
    "education": [
        "book", "education", "graduation", "diploma", "certificate", "teacher",
        "training", "library", "pencil", "pen", "writing", "notebook", "abc",
    ],
    "nature": [
        "tree", "leaf", "sun", "moon", "mountain", "flower", "rain", "fire",
        "lightning", "butterfly", "bird", "footprint",
    ],
    "objects": [
        "key", "lock", "gift", "lamp", "bulb", "light", "clock", "alarm",
        "glasses", "umbrella", "flag", "bell", "bottle", "cup", "coffee",
        "food", "grocery", "bread", "apple", "compass", "globe", "map",
        "house", "building", "door", "tool", "wrench", "screwdriver", "hammer",
        "paintbrush", "broom", "dustpan", "scissors", "paperclip", "thumbtack",
        "ribbon", "trophy", "award", "star", "target", "puzzle", "chess",
    ],
}


def slugify(text: str) -> str:
    """Convert text to a URL-friendly slug."""
    text = text.lower().strip()
    text = re.sub(r"[''`]s?\b", "", text)  # remove possessives
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = text.strip("_")
    return text


def classify_icon(name: str) -> str:
    """Classify an icon into a category based on its name."""
    name_lower = name.lower()
    for category, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if kw in name_lower:
                return category
    return "general"


def generate_keywords(name: str, category: str) -> list[str]:
    """Generate search keywords from icon name and category."""
    words = set(re.findall(r"[a-z]+", name.lower()))
    words.add(category)
    # Remove very short words
    words = {w for w in words if len(w) > 1}
    return sorted(words)


def parse_slide_title(spTree, ns) -> str:
    """Extract the text content from the slide's text placeholder."""
    for child in spTree:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if tag != "sp":
            continue
        # Check if it's a placeholder
        name = ""
        for el in child.iter():
            local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
            if local == "cNvPr":
                name = el.get("name", "")
                break
        if "Placeholder" not in name:
            continue
        # Get text
        texts = []
        for t in child.iter(f"{{{A_NS}}}t"):
            if t.text:
                texts.append(t.text)
        return " ".join(texts)
    return ""


def parse_icon_names_from_title(title: str) -> list[str]:
    """Parse comma-separated icon names from the slide title text."""
    if not title:
        return []
    # Split on commas, clean up whitespace
    parts = [p.strip() for p in title.split(",")]
    # Filter empty
    return [p for p in parts if p]


def extract_icons(pptx_path: str) -> dict:
    """Extract all icons from the PPTX file."""
    zf = zipfile.ZipFile(pptx_path, "r")

    slide_files = sorted(
        [n for n in zf.namelist() if re.match(r"ppt/slides/slide\d+\.xml$", n)],
        key=lambda x: int(re.search(r"slide(\d+)", x).group(1)),
    )

    all_icons = []
    seen_ids = set()
    color_report = {"standard": 0, "nonstandard": 0, "deviations": []}

    for slide_file in slide_files:
        slide_num = int(re.search(r"slide(\d+)", slide_file).group(1))
        xml = zf.read(slide_file)
        root = etree.fromstring(xml)
        spTree = root.find(f".//{{{P_NS}}}cSld/{{{P_NS}}}spTree")
        if spTree is None:
            continue

        # Parse title for names
        title = parse_slide_title(spTree, NS)
        icon_names = parse_icon_names_from_title(title)

        # Collect icon shapes (skip spTree's own nvGrpSpPr/grpSpPr and text placeholders)
        icon_elements = []
        for child in spTree:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

            if tag in ("nvGrpSpPr", "grpSpPr"):
                continue

            # Get shape name
            shape_name = ""
            for el in child.iter():
                local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
                if local == "cNvPr":
                    shape_name = el.get("name", "")
                    break

            # Skip text placeholders
            if "Placeholder" in shape_name:
                continue

            # Must have custom geometry to be an icon
            has_custGeom = child.find(f".//{{{A_NS}}}custGeom") is not None
            if not has_custGeom:
                continue

            icon_elements.append((child, tag, shape_name))

        # Map names to elements
        for i, (elem, tag, shape_name) in enumerate(icon_elements):
            # Get name from title list if available
            if i < len(icon_names):
                display_name = icon_names[i].strip()
            else:
                display_name = shape_name or f"icon_s{slide_num}_{i}"

            # Generate unique ID
            base_id = slugify(display_name) or f"s{slide_num}_{i}"
            icon_id = base_id
            counter = 2
            while icon_id in seen_ids:
                icon_id = f"{base_id}_{counter}"
                counter += 1
            seen_ids.add(icon_id)

            # Extract dimensions from xfrm
            xfrm = elem.find(f".//{{{A_NS}}}xfrm")
            orig_w = 914400  # default 1 inch
            orig_h = 914400
            if xfrm is not None:
                ext = xfrm.find(f"{{{A_NS}}}ext")
                if ext is not None:
                    orig_w = int(ext.get("cx", "914400"))
                    orig_h = int(ext.get("cy", "914400"))

            aspect = orig_w / orig_h if orig_h > 0 else 1.0

            # Extract colors
            colors = set()
            for srgb in elem.findall(f".//{{{A_NS}}}srgbClr"):
                colors.add(srgb.get("val", "").upper())

            # Check color assumptions
            expected = {"00AAE7", "000000"}
            extra = colors - expected - {"FFFFFF", "009DDC"}
            if extra:
                color_report["nonstandard"] += 1
                color_report["deviations"].append(
                    f"{icon_id} (slide {slide_num}): unexpected colors {extra}"
                )
            else:
                color_report["standard"] += 1

            # Serialize XML
            xml_str = etree.tostring(elem, encoding="unicode")

            # Classify
            category = classify_icon(display_name)
            keywords = generate_keywords(display_name, category)

            all_icons.append({
                "id": icon_id,
                "name": display_name,
                "category": category,
                "keywords": keywords,
                "source_slide": slide_num,
                "type": tag,
                "orig_width_emu": orig_w,
                "orig_height_emu": orig_h,
                "aspect_ratio": round(aspect, 4),
                "colors": sorted(colors),
                "xml": xml_str,
            })

    zf.close()

    # Build category summary
    cat_counts: dict[str, int] = {}
    for icon in all_icons:
        cat_counts[icon["category"]] = cat_counts.get(icon["category"], 0) + 1
    categories = [
        {"id": cat, "name": cat.title(), "count": count}
        for cat, count in sorted(cat_counts.items())
    ]

    return {
        "version": "1.0",
        "source": pptx_path.split("/")[-1],
        "total_icons": len(all_icons),
        "categories": categories,
        "color_report": color_report,
        "icons": all_icons,
    }


def main():
    if len(sys.argv) < 2:
        print("Usage: python extract_icons.py <path_to_iconography.pptx> [output.json]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "src/pptx_mcp_server/assets/icons.json"

    print(f"Extracting icons from: {pptx_path}")
    result = extract_icons(pptx_path)

    print(f"Extracted {result['total_icons']} icons")
    print(f"Categories: {result['categories']}")
    print(f"Color report: {result['color_report']['standard']} standard, "
          f"{result['color_report']['nonstandard']} nonstandard")
    if result["color_report"]["deviations"]:
        print("Deviations:")
        for d in result["color_report"]["deviations"][:10]:
            print(f"  {d}")

    # Remove color_report from output (build-time diagnostic only)
    color_report = result.pop("color_report")

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=None)

    file_size = len(open(output_path, "rb").read()) / 1024 / 1024
    print(f"Written to: {output_path} ({file_size:.1f} MB)")


if __name__ == "__main__":
    main()
