/** Join class names, skipping empty parts. */
export function cx(...parts: Array<string | false | undefined | null>): string {
	return parts.filter(Boolean).join(' ');
}
