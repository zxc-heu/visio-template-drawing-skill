import argparse
import json
from pathlib import Path


def bbox(shape):
    if "points" in shape and shape.get("kind") == "polygon":
        pts = shape.get("points") or []
        xs = [float(pts[i]) for i in range(0, len(pts), 2)]
        ys = [float(pts[i]) for i in range(1, len(pts), 2)]
        if xs and ys:
            return min(xs), min(ys), max(xs), max(ys)
    x = float(shape.get("x", 0))
    y = float(shape.get("y", 0))
    w = float(shape.get("width", 0))
    h = float(shape.get("height", 0))
    return x - w / 2, y - h / 2, x + w / 2, y + h / 2


def area(box):
    l, b, r, t = box
    return max(0.0, r - l) * max(0.0, t - b)


def intersection(a, b):
    l = max(a[0], b[0])
    bottom = max(a[1], b[1])
    r = min(a[2], b[2])
    top = min(a[3], b[3])
    if r <= l or top <= bottom:
        return 0.0
    return (r - l) * (top - bottom)


def point_in_box(x, y, box):
    return box[0] <= x <= box[2] and box[1] <= y <= box[3]


def orientation(ax, ay, bx, by, cx, cy):
    value = (by - ay) * (cx - bx) - (bx - ax) * (cy - by)
    if abs(value) < 1e-9:
        return 0
    return 1 if value > 0 else 2


def on_segment(ax, ay, bx, by, cx, cy):
    return min(ax, cx) <= bx <= max(ax, cx) and min(ay, cy) <= by <= max(ay, cy)


def segments_intersect(a, b, c, d):
    ax, ay = a
    bx, by = b
    cx, cy = c
    dx, dy = d
    o1 = orientation(ax, ay, bx, by, cx, cy)
    o2 = orientation(ax, ay, bx, by, dx, dy)
    o3 = orientation(cx, cy, dx, dy, ax, ay)
    o4 = orientation(cx, cy, dx, dy, bx, by)
    if o1 != o2 and o3 != o4:
        return True
    if o1 == 0 and on_segment(ax, ay, cx, cy, bx, by):
        return True
    if o2 == 0 and on_segment(ax, ay, dx, dy, bx, by):
        return True
    if o3 == 0 and on_segment(cx, cy, ax, ay, dx, dy):
        return True
    if o4 == 0 and on_segment(cx, cy, bx, by, dx, dy):
        return True
    return False


def segment_intersects_box(p1, p2, box):
    if point_in_box(p1[0], p1[1], box) or point_in_box(p2[0], p2[1], box):
        return True
    l, b, r, t = box
    edges = [((l, b), (r, b)), ((r, b), (r, t)), ((r, t), (l, t)), ((l, t), (l, b))]
    return any(segments_intersect(p1, p2, edge[0], edge[1]) for edge in edges)


def connector_points(conn):
    pts = conn.get("points") or []
    if len(pts) < 4:
        return []
    nums = [float(value) for value in pts]
    return [(nums[i], nums[i + 1]) for i in range(0, len(nums), 2)]


def is_text(shape):
    return shape.get("kind") == "text" or shape.get("textRole") in {"label", "caption", "annotation"} or bool(shape.get("text")) or bool(shape.get("richText"))


def is_image(shape):
    return shape.get("kind") == "image"


def is_background_or_container(shape, page_area=0):
    intent = str(shape.get("shapeIntent") or "").lower()
    sid = str(shape.get("id") or "").lower()
    if intent in {"background", "container", "layer", "region", "route-helper", "connector-helper"}:
        return True
    if sid.startswith("j_"):
        return True
    if sid.startswith("bg_") or sid.endswith("_layer"):
        return True
    if page_area > 0 and area(bbox(shape)) / page_area > 0.25 and not is_text(shape):
        return True
    return False


def allows_overlap(shape, other_id):
    if bool(shape.get("allowOverlap")):
        return True
    allowed = shape.get("allowedOverlaps") or []
    return other_id in allowed or "*" in allowed


def main():
    parser = argparse.ArgumentParser(description="Preflight drawing plan layout for obvious collisions.")
    parser.add_argument("plan_json")
    parser.add_argument("--min-overlap-ratio", type=float, default=0.05)
    args = parser.parse_args()

    plan_path = Path(args.plan_json)
    plan = json.loads(plan_path.read_text(encoding="utf-8-sig"))
    page = plan.get("page", {})
    page_w = float(page.get("width", 0) or 0)
    page_h = float(page.get("height", 0) or 0)
    page_area = page_w * page_h if page_w > 0 and page_h > 0 else 0
    shapes = plan.get("shapes", [])

    findings = []
    boxes = {}
    for shape in shapes:
        sid = shape.get("id", "<missing-id>")
        box = bbox(shape)
        boxes[sid] = box
        if page_w > 0 and page_h > 0:
            if box[0] < 0 or box[1] < 0 or box[2] > page_w or box[3] > page_h:
                findings.append({
                    "severity": "error",
                    "type": "page-overflow",
                    "shape": sid,
                    "message": f"shape extends outside page bounds: {box}",
                })

    for i, a in enumerate(shapes):
        for b in shapes[i + 1:]:
            aid = a.get("id", "<missing-id>")
            bid = b.get("id", "<missing-id>")
            ia = intersection(boxes[aid], boxes[bid])
            if ia <= 0:
                continue
            small = min(area(boxes[aid]), area(boxes[bid]))
            ratio = ia / small if small else 0
            if ratio < args.min_overlap_ratio:
                continue

            if (is_text(a) and is_image(b)) or (is_text(b) and is_image(a)):
                allowed = allows_overlap(a, bid) or allows_overlap(b, aid)
                findings.append({
                    "severity": "warning" if allowed else "error",
                    "type": "text-image-overlap",
                    "shapes": [aid, bid],
                    "message": f"text/label overlaps image ({ratio:.2%} of smaller box)"
                    + ("; allowed by plan" if allowed else ""),
                })
            elif is_image(a) and is_image(b):
                allowed = allows_overlap(a, bid) or allows_overlap(b, aid)
                findings.append({
                    "severity": "warning" if allowed else "error",
                    "type": "image-image-overlap",
                    "shapes": [aid, bid],
                    "message": f"images overlap ({ratio:.2%} of smaller box)"
                    + ("; allowed by plan" if allowed else ""),
                })

    for conn in plan.get("connectors", []):
        pts = connector_points(conn)
        if len(pts) < 2:
            continue
        conn_id = conn.get("id", "<missing-id>")
        for shape in shapes:
            sid = shape.get("id", "<missing-id>")
            if sid in {conn.get("from"), conn.get("to")}:
                continue
            if is_background_or_container(shape, page_area):
                continue
            if allows_overlap(shape, conn_id):
                continue
            box = boxes[sid]
            for p1, p2 in zip(pts, pts[1:]):
                if segment_intersects_box(p1, p2, box):
                    finding_type = "text-connector-overlap" if is_text(shape) else "node-connector-overlap"
                    findings.append({
                        "severity": "error",
                        "type": finding_type,
                        "shapes": [sid],
                        "connector": conn_id,
                        "message": "connector path intersects text/label or non-endpoint node",
                    })
                    break

    print(json.dumps({"plan": str(plan_path), "findings": findings}, ensure_ascii=False, indent=2))
    if any(item["severity"] == "error" for item in findings):
        raise SystemExit(2)


if __name__ == "__main__":
    main()
