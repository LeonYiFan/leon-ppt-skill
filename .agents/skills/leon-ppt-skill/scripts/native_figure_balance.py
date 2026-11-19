"""Reflow GUT evidence lanes around measured text and uncropped originals."""

from native_reference import _geometry


SINGLE_PANEL = "gut-single-panel"
METRIC_PANELS = "gut-figure-metrics-figure"
PAIRED_READOUTS = "gut-paired-evidence-and-readouts"
GUT_LEGEND = ("27", "28")
STACKED_PICTURE = "stacked-picture"


def body_geometry(profile, boxes, line_count):
    body = profile.get("body")
    if not body or not line_count:
        return None, 3.40
    x, y, width, _ = boxes[body]
    y = min(2.55, y)
    comparison = profile["id"] != SINGLE_PANEL
    height = (1.50 if line_count == 1 else 2.70) if comparison else (
        1.42 if line_count == 1 else 2.62)
    card_top = max(4.05 if line_count == 1 else 5.02, y + height + .19)
    return [x, y, width, height], card_top


def _move_header(geometry, boxes, slot, card, aligned_top=None):
    x, y, width, _ = card
    old = boxes[slot["container"]]
    header, label = slot["header"], slot["label"]
    bx, by, bw, bh = boxes[header]
    _, ly, _, lh = boxes[label]
    banner_y = y + by - old[1] if aligned_top is None else aligned_top
    # Keep the native horizontal inset and center the label across the safe width.
    geometry[header] = [x + bx - old[0], banner_y, width + bw - old[2], bh]
    label_y = y + ly - old[1] if aligned_top is None else banner_y + (bh - lh) / 2
    geometry[label] = [x + .25, label_y, width - .50, lh]
    geometry[slot["container"]] = card


def _contained(x, y, width, height, aspect):
    display_w, display_h = min(width, height * aspect), min(height, width / aspect)
    return [x + (width - display_w) / 2, y, display_w, display_h]


def _single_panel(boxes, profile, aspects, top, crop_bottom, legend):
    slot = profile["figures"][0]
    x, _, width, _ = boxes[slot["container"]]
    # Portrait originals need less lateral processing area than a microscopy strip.
    if len(aspects) == 1:
        narrower = min(width, max(18.5, (crop_bottom - top - 2.4) * aspects[0] / .92))
        x += (width - narrower) / 2
        width = narrower
    image_top = top + (1.58 if len(aspects) == 1 else 1.35)
    legend_height = 1.36 if legend else 0
    vertical_budget = crop_bottom - image_top - .55 - legend_height
    usable_w = width * .92
    if len(aspects) == 1:
        pictures = [_contained(x + (width - usable_w) / 2, image_top, usable_w,
                               vertical_budget, aspects[0])]
    else:
        # Give a short, wide schematic a little more width than the deeper plot.
        gap = .40
        weights = [a ** -.72 for a in aspects]
        pictures = []
        for aspect, weight in zip(aspects, weights, strict=True):
            height = min(usable_w / aspect, (vertical_budget - gap) * weight / sum(weights))
            picture_width = height * aspect
            pictures.append([x + (width - picture_width) / 2, image_top, picture_width, height])
            image_top += height + gap
    image_bottom = pictures[-1][1] + pictures[-1][3]
    minimum_height = 10.00 if legend else 8.25
    card_bottom = min(crop_bottom, max(top + minimum_height, image_bottom + legend_height + .55))
    geometry = {}
    _move_header(geometry, boxes, slot, [x, top, width, card_bottom - top])
    if legend:
        band, text = GUT_LEGEND
        legend_y = image_bottom + max(.42, min(.86, (card_bottom - image_bottom - .92) / 2))
        geometry[band] = [x + width * .04, legend_y, usable_w, .92]
        geometry[text] = [x + width * .05, legend_y + .08, width * .90, .76]
    return geometry, pictures, card_bottom


def balance_panels(profile, boxes, aspects, top, crop_bottom, legend, scale=1):
    if profile["id"] == SINGLE_PANEL:
        return _single_panel(boxes, profile, aspects, top, crop_bottom, legend)

    # Reserve a common clear band above the data and keep panel boundaries aligned.
    slots = profile["figures"]
    left = min(boxes[slot["container"]][0] for slot in slots)
    right = max(boxes[slot["container"]][0] + boxes[slot["container"]][2] for slot in slots)
    centre = (left + right) / 2
    lanes = [[centre + (x - centre) * scale, w * scale] for x, _, w, _ in
             (boxes[slot["container"]] for slot in slots)]
    paired = profile["id"] == PAIRED_READOUTS
    readout_height = 1.72 if paired else 0
    image_top = top + 1.40
    image_height = crop_bottom - image_top - .44 - readout_height
    geometry = {}
    pictures = []
    for (x, width), aspect in zip(lanes, aspects, strict=True):
        frame_width = width - .60
        pictures.append(_contained(x + (width - frame_width) / 2, image_top,
                                   frame_width, image_height * scale, aspect))
    tallest = max(picture[3] for picture in pictures)
    for picture in pictures:
        picture[1] += .32 * (tallest - picture[3])
    picture_bottom = max(y + height for _, y, _, height in pictures)
    card_bottom = min(crop_bottom, max(top + 8.25, picture_bottom + .44 + readout_height))
    banner_top = top + boxes[slots[0]["header"]][1] - boxes[slots[0]["container"]][1]
    for slot, (x, width) in zip(slots, lanes, strict=True):
        _move_header(geometry, boxes, slot, [x, top, width, card_bottom - top], banner_top)
    if paired:
        readout_y = card_bottom - 1.47
        for (x, width), (band, text) in zip(lanes, (("14", "15"), ("20", "21")), strict=True):
            _, band_y, _, band_height = boxes[band]
            _, text_y, _, text_height = boxes[text]
            geometry[band] = [x + .56, readout_y, width - 1.12, band_height]
            geometry[text] = [x + .76, readout_y + text_y - band_y, width - 1.52, text_height]
    return geometry, pictures, card_bottom


def balance_metric_panels(profile, boxes, aspects, top, crop_bottom, scale=1):
    """Keep the three GUT 41 headers aligned and reserve room for two metrics."""
    metric = "88"
    x, origin_y, width, origin_height = boxes[metric]
    height = min(origin_height * scale, crop_bottom - top)
    geometry = {metric: [x, top, width, height]}
    for ident in ("89", "90", "95", "96", "99", "100", "91", "92"):
        lx, ly, lw, lh = boxes[ident]
        fraction = (ly - origin_y) / origin_height
        geometry[ident] = [lx, top + height * fraction, lw, lh]
    first = profile["figures"][0]
    banner_top = top + boxes[first["header"]][1] - boxes[first["container"]][1]
    geometry["89"][1] = banner_top
    geometry["90"][1] = banner_top + (geometry["89"][3] - geometry["90"][3]) / 2
    pictures = []
    for slot, aspect in zip(profile["figures"], aspects, strict=True):
        left, _, panel_width, _ = boxes[slot["container"]]
        if left < x:
            left += panel_width * (1 - scale)
        panel_width *= scale
        _move_header(geometry, boxes, slot, [left, top, panel_width, height], banner_top)
        usable_w = panel_width - .55
        picture = _contained(left + (panel_width - usable_w) / 2, top + 1.40,
                             usable_w, height - 1.80, aspect)
        picture[1] += .47 * (height - 1.80 - picture[3])
        pictures.append(picture)
    return geometry, pictures, top + height


def coordinates_in_group(donor, node, box, cm_emu):
    """Convert slide coordinates to the original group's local coordinates."""
    ancestors = list(node.iterancestors())
    mapping = (1, 1, 0, 0)
    for parent in reversed(ancestors):
        if parent is donor.tree:
            continue
        if parent.tag.endswith("}grpSp"):
            _, mapping = _geometry(parent, donor, mapping)
    sx, sy, dx, dy = mapping
    x, y, width, height = box
    return [(x - dx / cm_emu) / sx, (y - dy / cm_emu) / sy, width / sx, height / sy]
