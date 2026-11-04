import { DocumentError } from "@dagent-office/core";

/** Quoted CSS values remain data, including hostile quotes and backslashes. */
export function cssFontName(name: string): string {
  if (!name || /[\u0000-\u001f\u007f]/.test(name)) throw new DocumentError("InvalidDocument", "A font name is empty or contains control characters.");
  return '"' + name.replace(/["\\]/g, "\\$&") + '"';
}
