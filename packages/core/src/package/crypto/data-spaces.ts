/**
 * The `\u0006DataSpaces` storage of an encrypted package ([MS-OFFCRYPTO] 2.1
 * and 2.3.4.1): it declares that the `EncryptedPackage` stream goes through
 * the strong encryption transform. Excel refuses encrypted files without it.
 */
import type { CfbStreamInput } from '../cfb/write';

const ROOT = '\u0006DataSpaces';
const DATA_SPACE = 'StrongEncryptionDataSpace';
const TRANSFORM = 'StrongEncryptionTransform';
const ENCRYPTION_TRANSFORM_ID = '{FF9A3F03-56EF-4613-BDD5-5A41C1D07246}';

/** Little-endian structure writer. */
class StructWriter {
	private readonly out: number[] = [];

	get length(): number {
		return this.out.length;
	}

	u16(value: number): this {
		this.out.push(value & 0xff, (value >>> 8) & 0xff);
		return this;
	}

	u32(value: number): this {
		return this.u16(value & 0xffff).u16(value >>> 16);
	}

	/** Version structure (2.1.4): major, minor. */
	version(major: number, minor: number): this {
		return this.u16(major).u16(minor);
	}

	/** UNICODE-LP-P4 (2.1.2): byte length, UTF-16LE characters, zero padding to 4 bytes. */
	unicode(text: string): this {
		this.u32(text.length * 2);
		for (let i = 0; i < text.length; i++) {
			this.u16(text.charCodeAt(i));
		}
		if (text.length % 2) {
			this.u16(0);
		}
		return this;
	}

	bytes(data: Uint8Array): this {
		this.out.push(...data);
		return this;
	}

	toBytes(): Uint8Array {
		return Uint8Array.from(this.out);
	}
}

/** DataSpaceVersionInfo (2.1.5): reader, updater and writer version 1.0. */
function versionStream(): Uint8Array {
	const w = new StructWriter().unicode('Microsoft.Container.DataSpaces');
	return w.version(1, 0).version(1, 0).version(1, 0).toBytes();
}

/** DataSpaceMap (2.1.6): the EncryptedPackage stream uses StrongEncryptionDataSpace. */
function dataSpaceMap(): Uint8Array {
	// DataSpaceReferenceComponent (2.1.6.2): type 0 is a stream.
	const entry = new StructWriter().u32(1).u32(0).unicode('EncryptedPackage').unicode(DATA_SPACE);
	const map = new StructWriter().u32(8).u32(1);
	return map
		.u32(entry.length + 4)
		.bytes(entry.toBytes())
		.toBytes();
}

/** DataSpaceDefinition (2.1.7): the data space applies StrongEncryptionTransform. */
function dataSpaceDefinition(): Uint8Array {
	return new StructWriter().u32(8).u32(1).unicode(TRANSFORM).toBytes();
}

/** TransformInfoHeader (2.1.8) and an empty EncryptionTransformInfo (2.1.9). */
function primaryStream(): Uint8Array {
	// The header length counts the bytes up to and including the transform id.
	const id = new StructWriter().u32(1).unicode(ENCRYPTION_TRANSFORM_ID);
	const w = new StructWriter().u32(id.length + 4).bytes(id.toBytes());
	w.unicode('Microsoft.Container.EncryptionTransform');
	w.version(1, 0).version(1, 0).version(1, 0);
	// EncryptionName (empty UTF-8-LP-P4), EncryptionBlockSize, CipherMode, Reserved (4).
	return w.u32(0).u32(0).u32(0).u32(4).toBytes();
}

/** The streams of the `\u0006DataSpaces` storage, as compound file paths. */
export function dataSpacesStreams(): CfbStreamInput[] {
	return [
		{ path: `${ROOT}/Version`, data: versionStream() },
		{ path: `${ROOT}/DataSpaceMap`, data: dataSpaceMap() },
		{ path: `${ROOT}/DataSpaceInfo/${DATA_SPACE}`, data: dataSpaceDefinition() },
		{ path: `${ROOT}/TransformInfo/${TRANSFORM}/\u0006Primary`, data: primaryStream() },
	];
}
