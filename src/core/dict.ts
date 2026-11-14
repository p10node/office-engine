/**
 * Lookups by a key read from a file (an attribute value, an element name, a style id).
 *
 * A plain object answers "constructor", "toString" or "__proto__" with what it inherits: a function or an object,
 * which then travels as a color, a border style, an alignment… and can't be sent out of the worker (the document
 * comes out blank) or breaks the code that expects a string.
 */

/** A lookup table with no inherited keys: a key the table was not given is undefined. */
export function dict<T extends object>(table: T): T {
  return Object.assign(Object.create(null) as T, table)
}

/** `obj[key]`, only when the object itself holds the key: for objects built elsewhere (XML attributes, cloned models). */
export function own<T>(obj: Record<string, T> | undefined, key: string | undefined): T | undefined {
  return obj !== undefined && key !== undefined && Object.hasOwn(obj, key) ? obj[key] : undefined
}
