import type { ExtendedPropertyValue, SingleValueExtendedProperty } from "./types.js";

/**
 * MAPI property tags read as Outlook single-value legacy extended properties.
 * Graph id format is "{type} {proptag}" — see extended properties overview.
 * Types are PtypString (0x001F), so the Graph type name is String.
 */
export const CREATOR_PROPERTIES = [
  { graphId: "String 0x3FF8", canonicalName: "PidTagCreatorName" },
  { graphId: "String 0x3FFA", canonicalName: "PidTagLastModifierName" },
  { graphId: "String 0x5D01", canonicalName: "PidTagSenderSmtpAddress" },
  { graphId: "String 0x5D02", canonicalName: "PidTagSentRepresentingSmtpAddress" },
  { graphId: "String 0x0C1A", canonicalName: "PidTagSenderName" },
  { graphId: "String 0x0042", canonicalName: "PidTagSentRepresentingName" },
  { graphId: "String 0x0C1F", canonicalName: "PidTagSenderEmailAddress" },
  { graphId: "String 0x0065", canonicalName: "PidTagSentRepresentingEmailAddress" },
] as const;

const MINIMAL_PROPERTY_IDS = ["String 0x3FF8", "String 0x3FFA"] as const;

const CANONICAL_BY_ID = new Map(
  CREATOR_PROPERTIES.map((property) => [property.graphId.toLowerCase(), property.canonicalName]),
);

export function extendedPropertyExpand(graphIds: readonly string[]): string {
  const filter = graphIds.map((id) => `id eq '${id}'`).join(" or ");
  return `singleValueExtendedProperties($filter=${filter})`;
}

export const FULL_PROPERTY_EXPAND = extendedPropertyExpand(
  CREATOR_PROPERTIES.map((property) => property.graphId),
);

export const MINIMAL_PROPERTY_EXPAND = extendedPropertyExpand(MINIMAL_PROPERTY_IDS);

export function readExtendedProperties(
  properties: SingleValueExtendedProperty[] | undefined,
): { list: ExtendedPropertyValue[]; byName: Record<string, string> } {
  const list: ExtendedPropertyValue[] = [];
  const byName: Record<string, string> = {};
  for (const property of properties ?? []) {
    if (!property.id || property.value === undefined || property.value === null) {
      continue;
    }
    const canonicalName = CANONICAL_BY_ID.get(property.id.toLowerCase()) ?? null;
    list.push({ id: property.id, canonicalName, value: property.value });
    if (canonicalName) {
      byName[canonicalName] = property.value;
    }
  }
  return { list, byName };
}
