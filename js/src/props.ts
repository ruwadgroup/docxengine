/**
 * The §16 closed formatting prop set: parsing the tool shorthand into canonical
 * spec names, and emitting `w:rPr`/`w:pPr` children in the fixed §16 order.
 *
 * Shorthands map verbatim (§Phase-2 preamble): `size`→`size_pt` (points),
 * `spacing_after`→`spacing_after_pt`, `spacing_before`→`spacing_before_pt`;
 * `spacing` is a line multiplier and is **ignored** by the style/format
 * writers. A boolean `false` emits the toggle-off form (`<w:b w:val="0"/>`).
 *
 * The Python twin is `python/src/docxengine/_props.py`-equivalent — byte parity.
 */
import {
  type ElementSlice,
  type SpliceEdit,
  childElements,
  escapeAttr,
  findElement,
  splice,
} from "./xmlscan.js";

/** Canonical (spec-named) properties resolved from the tool shorthand. */
export interface CanonProps {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: string; // RRGGBB, uppercased, no '#'
  size_pt?: number;
  alignment?: string; // left|center|right|both
  spacing_before_pt?: number;
  spacing_after_pt?: number;
}

const ALIGN_MAP: Readonly<Record<string, string>> = {
  left: "left",
  center: "center",
  right: "right",
  justify: "both",
  both: "both",
};

/** Map the tool/style shorthand object onto canonical §16 prop names. */
export function canonicalizeProps(raw: Record<string, unknown> | undefined): CanonProps {
  const out: CanonProps = {};
  if (raw == null || typeof raw !== "object") return out;
  const r = raw as Record<string, unknown>;
  if (typeof r["bold"] === "boolean") out.bold = r["bold"];
  if (typeof r["italic"] === "boolean") out.italic = r["italic"];
  if (typeof r["underline"] === "boolean") out.underline = r["underline"];
  const color = r["color"];
  if (typeof color === "string" && color !== "") {
    out.color = color.replace(/^#/, "").toUpperCase();
  }
  // size_pt is canonical; `size` is the tool shorthand (points).
  const size = r["size_pt"] ?? r["size"];
  if (typeof size === "number") out.size_pt = size;
  const align = r["alignment"];
  if (typeof align === "string" && align in ALIGN_MAP) {
    out.alignment = ALIGN_MAP[align] as string;
  }
  const before = r["spacing_before_pt"] ?? r["spacing_before"];
  if (typeof before === "number") out.spacing_before_pt = before;
  const after = r["spacing_after_pt"] ?? r["spacing_after"];
  if (typeof after === "number") out.spacing_after_pt = after;
  // `spacing` (line multiplier) is intentionally ignored by these writers.
  return out;
}

// ---------------------------------------------------------------------------
// Emission of individual children (§16 closed prop set, fixed order)
// ---------------------------------------------------------------------------

/** One toggle element: present-true → `<w:x/>`, false → `<w:x w:val="0"/>`. */
function toggle(name: string, on: boolean): string {
  return on ? `<${name}/>` : `<${name} w:val="0"/>`;
}

/** Run-property children in §16 order: b, i, u, color, sz. Each is `{name → xml}`. */
export function runPropChildren(p: CanonProps): { name: string; xml: string }[] {
  const out: { name: string; xml: string }[] = [];
  if (p.bold !== undefined) out.push({ name: "w:b", xml: toggle("w:b", p.bold) });
  if (p.italic !== undefined) out.push({ name: "w:i", xml: toggle("w:i", p.italic) });
  if (p.underline !== undefined) {
    out.push({ name: "w:u", xml: p.underline ? '<w:u w:val="single"/>' : '<w:u w:val="none"/>' });
  }
  if (p.color !== undefined) {
    out.push({ name: "w:color", xml: `<w:color w:val="${escapeAttr(p.color)}"/>` });
  }
  if (p.size_pt !== undefined) {
    out.push({ name: "w:sz", xml: `<w:sz w:val="${Math.round(p.size_pt * 2)}"/>` });
  }
  return out;
}

/** Paragraph-property children in §16 order (after pStyle): jc, spacing. */
export function paraPropChildren(p: CanonProps): { name: string; xml: string }[] {
  const out: { name: string; xml: string }[] = [];
  if (p.alignment !== undefined) {
    out.push({ name: "w:jc", xml: `<w:jc w:val="${escapeAttr(p.alignment)}"/>` });
  }
  if (p.spacing_before_pt !== undefined || p.spacing_after_pt !== undefined) {
    const attrs: string[] = [];
    if (p.spacing_before_pt !== undefined) {
      attrs.push(`w:before="${Math.round(p.spacing_before_pt * 20)}"`);
    }
    if (p.spacing_after_pt !== undefined) {
      attrs.push(`w:after="${Math.round(p.spacing_after_pt * 20)}"`);
    }
    out.push({ name: "w:spacing", xml: `<w:spacing ${attrs.join(" ")}/>` });
  }
  return out;
}

/** The concatenated §16 run-property markup (no wrapper). */
export function runPropsInner(p: CanonProps): string {
  return runPropChildren(p)
    .map((c) => c.xml)
    .join("");
}

/** The concatenated §16 paragraph-property markup (no wrapper). */
export function paraPropsInner(p: CanonProps): string {
  return paraPropChildren(p)
    .map((c) => c.xml)
    .join("");
}

// ---------------------------------------------------------------------------
// Merging props into an existing rPr/pPr (§16: replace same-named children,
// create in §16 order) — used by docx_style define and docx_format.
// ---------------------------------------------------------------------------

const RPR_ORDER = ["w:b", "w:i", "w:u", "w:color", "w:sz"];
const PPR_ORDER = ["w:pStyle", "w:jc", "w:spacing"];

/**
 * Merge `children` (name → xml) into the inner markup of an rPr/pPr, replacing
 * same-named existing children and inserting new ones at their §16-order slot.
 * `existingInner` is the current child markup (no wrapper); returns the merged
 * inner markup. `order` pins the canonical child sequence.
 */
export function mergeChildren(
  existingInner: string,
  children: { name: string; xml: string }[],
  order: readonly string[],
): string {
  // Snapshot existing direct children with their markup.
  const present: { name: string; xml: string }[] = [];
  for (const el of childElements(existingInner, 0, existingInner.length)) {
    present.push({ name: el.name, xml: existingInner.slice(el.start, el.end) });
  }
  // Overlay: replace same-named, keep order for unknown children at their spot.
  const byName = new Map<string, string>();
  for (const c of present) byName.set(c.name, c.xml);
  for (const c of children) byName.set(c.name, c.xml);
  // Re-emit: ordered known children first (in §16 order), then any extras that
  // are not in the order list, preserving their original relative order.
  const emitted: string[] = [];
  const usedNames = new Set<string>();
  for (const name of order) {
    if (byName.has(name)) {
      emitted.push(byName.get(name) as string);
      usedNames.add(name);
    }
  }
  for (const c of present) {
    if (!order.includes(c.name) && !usedNames.has(c.name)) {
      emitted.push(byName.get(c.name) ?? c.xml);
      usedNames.add(c.name);
    }
  }
  return emitted.join("");
}

/** Merge run props into a paragraph/style/run, creating `w:rPr` when absent. */
export function mergeRunProps(p: CanonProps): {
  children: { name: string; xml: string }[];
  order: readonly string[];
} {
  return { children: runPropChildren(p), order: RPR_ORDER };
}

/** Merge paragraph props, creating `w:pPr` when absent. */
export function mergeParaProps(p: CanonProps): {
  children: { name: string; xml: string }[];
  order: readonly string[];
} {
  return { children: paraPropChildren(p), order: PPR_ORDER };
}

export { RPR_ORDER, PPR_ORDER };

// ---------------------------------------------------------------------------
// Splicing direct rPr/pPr into a run/paragraph (§16 docx_format direct path)
// ---------------------------------------------------------------------------

/**
 * Build the §16 splice that merges run props into one `w:r`, creating `w:rPr`
 * as the first child when absent. Returns null when there is nothing to write.
 */
export function runPropsEdit(xml: string, run: ElementSlice, p: CanonProps): SpliceEdit | null {
  const children = runPropChildren(p);
  if (children.length === 0) return null;
  if (run.selfClosed) {
    const open = xml.slice(run.start, run.end - 2); // strip "/>"
    return {
      start: run.start,
      end: run.end,
      text: `${open}><w:rPr>${runPropsInner(p)}</w:rPr></w:r>`,
    };
  }
  const rPr = findElement(xml, "w:rPr", run.contentStart, run.contentEnd);
  if (rPr && rPr.start === run.contentStart) {
    const inner = rPr.selfClosed ? "" : xml.slice(rPr.contentStart, rPr.contentEnd);
    const merged = mergeChildren(inner, children, RPR_ORDER);
    return { start: rPr.start, end: rPr.end, text: `<w:rPr>${merged}</w:rPr>` };
  }
  return {
    start: run.contentStart,
    end: run.contentStart,
    text: `<w:rPr>${runPropsInner(p)}</w:rPr>`,
  };
}

/**
 * Build the §16 splice that merges paragraph props into one `w:p`'s `w:pPr`,
 * creating it (as the first child) when absent. pStyle stays first.
 */
export function paraPropsEdit(xml: string, para: ElementSlice, p: CanonProps): SpliceEdit | null {
  const children = paraPropChildren(p);
  if (children.length === 0) return null;
  if (para.selfClosed) {
    const open = xml.slice(para.start, para.end - 2);
    return {
      start: para.start,
      end: para.end,
      text: `${open}><w:pPr>${paraPropsInner(p)}</w:pPr></w:p>`,
    };
  }
  const kids = childElements(xml, para.contentStart, para.contentEnd);
  const pPr = kids.find((k) => k.name === "w:pPr");
  if (pPr) {
    const inner = pPr.selfClosed ? "" : xml.slice(pPr.contentStart, pPr.contentEnd);
    const merged = mergeChildren(inner, children, PPR_ORDER);
    return { start: pPr.start, end: pPr.end, text: `<w:pPr>${merged}</w:pPr>` };
  }
  return {
    start: para.contentStart,
    end: para.contentStart,
    text: `<w:pPr>${paraPropsInner(p)}</w:pPr>`,
  };
}

void splice;
