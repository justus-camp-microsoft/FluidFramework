/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { assert, unreachableCase, fail } from "@fluidframework/core-utils/internal";
import type { IIdCompressor } from "@fluidframework/id-compressor";

import {
	CursorLocationType,
	type FieldKey,
	type FieldKindData,
	type FieldKindIdentifier,
	type ITreeCursorSynchronous,
	type TreeChunk,
	type TreeFieldStoredSchema,
	type TreeNodeSchemaIdentifier,
	type Value,
	forEachNode,
} from "../../../core/index.js";
import { getOrCreate } from "../../../util/index.js";

import type { Counter, DeduplicationTable } from "./chunkCodecUtilities.js";
import {
	type BufferFormat as BufferFormatGeneric,
	Shape as ShapeGeneric,
	updateShapesAndIdentifiersEncoding,
} from "./chunkEncodingGeneric.js";
// eslint-disable-next-line @typescript-eslint/no-unused-vars -- Referenced by doc comments
import type { FieldBatchEncodingContext, IncrementalEncoder } from "./codecs.js";
import type { FieldBatch } from "./fieldBatch.js";
import {
	type EncodedAnyShape,
	type EncodedChunkShapeV1,
	type EncodedChunkShape,
	type EncodedChunkShapeV2,
	type EncodedFieldBatchV1OrV2,
	type EncodedNestedArrayShape,
	type EncodedValueShape,
	FieldBatchFormatVersion,
	SpecialField,
	supportsIncrementalEncoding,
} from "./format/index.js";

/**
 * Minimum number of consecutive same-shape sibling chunks required to be encoded as an
 * {@link EncodedInlineArrayShape} run (one shape reference for the whole run) rather than
 * one shape reference per chunk.
 *
 * @remarks
 * The break-even point depends on the size of shape references on the wire (1–3 bytes each
 * after JSON serialization) and the per-shape-table-entry cost of a new
 * {@link EncodedInlineArrayShape} (~30 bytes serialized when not deduplicated). Caching by
 * `(length, shape)` in {@link EncoderContext.inlineArrayShapeForRun} amortizes the
 * shape-table cost when the same `(length, shape)` recurs across fields, which is common
 * for uniform-content workloads (e.g. long runs of the same character).
 */
const RUN_DETECTION_MIN_LENGTH = 4;

/**
 * Encode data from `FieldBatch` into an `EncodedFieldBatch`.
 *
 * Optimized for encoded size and encoding performance.
 *
 * Most of the compression strategy comes from the policy provided via `context`.
 */
export function compressedEncode(
	fieldBatch: FieldBatch,
	context: EncoderContext,
): EncodedFieldBatchV1OrV2 {
	const onComplete = context.preEncodeHook?.(fieldBatch, context);
	try {
		const batchBuffer: BufferFormat[] = [];

		// Populate buffer, including shape and identifier references
		for (const cursor of fieldBatch) {
			const buffer: BufferFormat = [];
			anyFieldEncoder.encodeField(cursor, context, buffer);
			batchBuffer.push(buffer);
		}
		return updateShapesAndIdentifiersEncoding(context.version, batchBuffer);
	} finally {
		onComplete?.();
	}
}

export type BufferFormat = BufferFormatGeneric<EncodedChunkShape>;
export type Shape = ShapeGeneric<EncodedChunkShape>;

/**
 * Like {@link FieldEncoder}, except data will be prefixed with the key.
 */
export interface KeyedFieldEncoder {
	readonly key: FieldKey;
	readonly encoder: FieldEncoder;
}

/**
 * An encoder with an associated shape.
 */
export interface Encoder {
	/**
	 * The shape which describes how the encoded data is laid out.
	 * Used by decoders to interpret the output of `encodeNode`.
	 */
	readonly shape: Shape;
}

/**
 * An encoder for a specific shape of node.
 *
 * Can only be used with compatible nodes.
 */
export interface NodeEncoder extends Encoder {
	/**
	 * @param cursor - in Nodes mode. Does not move cursor.
	 */
	encodeNode(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void;
}

/**
 * Like {@link NodeEncoder}, except encodes a run of nodes.
 */
export interface NodesEncoder extends Encoder {
	/**
	 * @param cursor - in Nodes mode. Moves cursor however many nodes it encodes.
	 */
	encodeNodes(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void;
}

/**
 * Like {@link NodeEncoder}, except encodes a field.
 */
export interface FieldEncoder extends Encoder {
	/**
	 * @param cursor - in Fields mode. Encodes entire field.
	 */
	encodeField(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void;
}

/**
 * Makes a {@link FieldEncoder} which runs `encoder` on every node in the field.
 * This does not encode the number nodes: the user of this may need to encode that elsewhere.
 */
export function asFieldEncoder(encoder: NodeEncoder): FieldEncoder {
	return {
		encodeField(
			cursor: ITreeCursorSynchronous,
			context: EncoderContext,
			outputBuffer: BufferFormat,
		): void {
			forEachNode(cursor, () => encoder.encodeNode(cursor, context, outputBuffer));
		},
		shape: encoder.shape,
	};
}

/**
 * Adapt a {@link NodeEncoder} to a {@link NodesEncoder} which invokes `encoder` once.
 */
export function asNodesEncoder(encoder: NodeEncoder): NodesEncoder {
	return {
		encodeNodes(
			cursor: ITreeCursorSynchronous,
			context: EncoderContext,
			outputBuffer: BufferFormat,
		): void {
			encoder.encodeNode(cursor, context, outputBuffer);
			cursor.nextNode();
		},
		shape: encoder.shape,
	};
}

/**
 * Encodes a chunk with {@link EncodedAnyShape} by prefixing the data with its shape.
 */
export class AnyShape extends ShapeGeneric<EncodedChunkShape> {
	private constructor() {
		super();
	}
	public static readonly instance = new AnyShape();

	public encodeShape(
		identifiers: DeduplicationTable<string>,
		shapes: DeduplicationTable<Shape>,
	): EncodedChunkShapeV1 {
		const encodedAnyShape: EncodedAnyShape = 0;
		return { d: encodedAnyShape };
	}

	public countReferencedShapesAndIdentifiers(
		identifiers: Counter<string>,
		shapeDiscovered: (shape: Shape) => void,
	): void {}

	public static encodeField(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
		encoder: FieldEncoder,
	): void {
		outputBuffer.push(encoder.shape);
		encoder.encodeField(cursor, context, outputBuffer);
	}

	public static encodeNode(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
		encoder: NodeEncoder,
	): void {
		outputBuffer.push(encoder.shape);
		encoder.encodeNode(cursor, context, outputBuffer);
	}

	public static encodeNodes(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
		encoder: NodesEncoder,
	): void {
		outputBuffer.push(encoder.shape);
		encoder.encodeNodes(cursor, context, outputBuffer);
	}
}

/**
 * Encodes a single node polymorphically.
 */
export const anyNodeEncoder: NodeEncoder = {
	encodeNode(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		// TODO: Fast path uniform chunk content.
		const nodeEncoder = context.nodeEncoderFromSchema(cursor.type);
		AnyShape.encodeNode(cursor, context, outputBuffer, nodeEncoder);
	},

	shape: AnyShape.instance,
};

/**
 * Encodes a field polymorphically.
 */
export const anyFieldEncoder: FieldEncoder = {
	encodeField(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		// TODO: Fast path uniform chunks.

		if (cursor.getFieldLength() === 0) {
			const shape = InlineArrayEncoder.empty;
			AnyShape.encodeField(cursor, context, outputBuffer, shape);
		} else if (cursor.getFieldLength() === 1) {
			// Fast path chunk of size one size one at least: skip nested array.
			cursor.enterNode(0);
			anyNodeEncoder.encodeNode(cursor, context, outputBuffer);
			cursor.exitNode();
		} else {
			// TODO: more efficient encoding for common cases.
			// Could try to find more specific shape compatible with all children than `anyNodeEncoder`.

			const shape = context.nestedArrayEncoder(anyNodeEncoder);
			AnyShape.encodeField(cursor, context, outputBuffer, shape);
		}
	},

	shape: AnyShape.instance,
};

/**
 * Encodes a chunk using {@link EncodedInlineArrayShape}.
 * @remarks
 * The fact this is also a Shape is an implementation detail of the encoder: that allows the shape it uses to be itself,
 * which is an easy way to keep all the related code together without extra objects.
 */
export class InlineArrayEncoder
	extends ShapeGeneric<EncodedChunkShape>
	implements NodesEncoder, FieldEncoder
{
	public static readonly empty: InlineArrayEncoder = new InlineArrayEncoder(0, {
		get shape() {
			// Not actually used, makes count work without adding an additional shape.
			return InlineArrayEncoder.empty;
		},
		encodeNodes(
			cursor: ITreeCursorSynchronous,
			context: EncoderContext,
			outputBuffer: BufferFormat,
		): void {
			fail(0xb4d /* Empty array should not encode any nodes */);
		},
	});

	/**
	 * @param length - number of invocations of `inner`.
	 */
	public constructor(
		public readonly length: number,
		public readonly inner: NodesEncoder,
	) {
		super();
	}

	public encodeNodes(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		// Linter is wrong about this loop being for-of compatible.
		// eslint-disable-next-line @typescript-eslint/prefer-for-of
		for (let index = 0; index < this.length; index++) {
			this.inner.encodeNodes(cursor, context, outputBuffer);
		}
	}

	public encodeField(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		// Its possible individual items from this array encode multiple nodes, so don't assume === here.
		assert(
			cursor.getFieldLength() >= this.length,
			0x73c /* unexpected length for fixed length array */,
		);
		cursor.firstNode();
		this.encodeNodes(cursor, context, outputBuffer);
		assert(
			cursor.mode === CursorLocationType.Fields,
			0x73d /* should return to fields mode when finished encoding */,
		);
	}

	public encodeShape(
		identifiers: DeduplicationTable<string>,
		shapes: DeduplicationTable<Shape>,
	): EncodedChunkShapeV1 {
		return {
			b: {
				length: this.length,
				shape: shapes.valueToIndex.get(this.inner.shape) ?? fail(0xb4e /* missing shape */),
			},
		};
	}

	public countReferencedShapesAndIdentifiers(
		identifiers: Counter<string>,
		shapeDiscovered: (shape: Shape) => void,
	): void {
		shapeDiscovered(this.inner.shape);
	}

	public get shape(): this {
		return this;
	}
}

/**
 * Encodes the shape for a nested array as {@link EncodedNestedArrayShape} shape.
 */
export class NestedArrayShape extends ShapeGeneric<EncodedChunkShape> {
	/**
	 * @param innerShape - The shape of each item in this nested array.
	 */
	public constructor(public readonly innerShape: Shape) {
		super();
	}

	public encodeShape(
		identifiers: DeduplicationTable<string>,
		shapes: DeduplicationTable<Shape>,
	): EncodedChunkShape {
		const shape: EncodedNestedArrayShape =
			shapes.valueToIndex.get(this.innerShape) ??
			fail(0xb4f /* index for shape not found in table */);
		return {
			a: shape,
		};
	}

	public countReferencedShapesAndIdentifiers(
		identifiers: Counter<string>,
		shapeDiscovered: (shape: Shape) => void,
	): void {
		shapeDiscovered(this.innerShape);
	}
}

/**
 * Encodes a field as a nested array with the {@link EncodedNestedArrayShape} shape.
 * @remarks
 * The fact this is also exposes a Shape is an implementation detail: it allows the shape it uses to be itself
 * which is an easy way to keep all the related code together without extra objects.
 */
export class NestedArrayEncoder implements FieldEncoder {
	public constructor(
		public readonly innerEncoder: NodeEncoder,
		public readonly shape: NestedArrayShape = new NestedArrayShape(innerEncoder.shape),
	) {}

	public encodeField(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		const length = cursor.getFieldLength();

		// Run detection only fires for vText, where the polymorphic dispatch (inner shape =
		// AnyShape) emits a per-node shape reference that can be hoisted when adjacent nodes
		// share the same shape. For non-AnyShape inners, the field's wire shape already pins
		// the inner shape once, so per-node references don't exist to dedupe.
		const runDetectionApplies =
			context.version === FieldBatchFormatVersion.vTextExperimental &&
			this.innerEncoder.shape === AnyShape.instance;

		if (!runDetectionApplies) {
			const flatBuffer: BufferFormat = [];
			let allNonZeroSize = true;
			forEachNode(cursor, () => {
				const before = flatBuffer.length;
				this.innerEncoder.encodeNode(cursor, context, flatBuffer);
				allNonZeroSize &&= flatBuffer.length - before !== 0;
			});
			if (flatBuffer.length === 0) {
				// This relies on the number of inner chunks being the same as the number of nodes.
				// If making inner a `NodesEncoder`, this code will have to be adjusted accordingly.
				outputBuffer.push(length);
			} else {
				assert(
					allNonZeroSize,
					0x73e /* either all or none of the members of a nested array must be 0 sized, or there is no way the decoder could process the content correctly. */,
				);
				outputBuffer.push(flatBuffer);
			}
			return;
		}

		// Run-detection path: pre-encode each child into a side buffer, group adjacent children
		// whose first emitted entry (the per-node shape reference pushed by AnyShape semantics)
		// is the same, and emit each run of length >= RUN_DETECTION_MIN_LENGTH as one
		// EncodedInlineArrayShape header followed by the concatenated data parts. Shorter runs
		// keep the per-node shape-reference layout. The wire format remains a NestedArrayShape
		// with inner=AnyShape; runs are valid AnyShape-dispatched chunks whose shape is
		// EncodedInlineArrayShape, which is supported in v1 and later.
		const sideBuffers: BufferFormat[] = [];
		forEachNode(cursor, () => {
			const sideBuf: BufferFormat = [];
			this.innerEncoder.encodeNode(cursor, context, sideBuf);
			sideBuffers.push(sideBuf);
		});

		if (sideBuffers.every((b) => b.length === 0)) {
			outputBuffer.push(length);
			return;
		}

		assert(
			sideBuffers.every((b) => b.length > 0),
			"either all or none of the members of a nested array must be 0 sized",
		);

		const buffer: BufferFormat = [];
		let i = 0;
		while (i < sideBuffers.length) {
			const head = sideBuffers[i] ?? fail("missing side buffer");
			const headShape = head[0];
			let j = i + 1;
			while (j < sideBuffers.length && sideBuffers[j]?.[0] === headShape) {
				j++;
			}
			const runLength = j - i;
			if (runLength >= RUN_DETECTION_MIN_LENGTH && headShape instanceof ShapeGeneric) {
				const inlineShape = context.inlineArrayShapeForRun(runLength, headShape);
				buffer.push(inlineShape);
				for (let k = i; k < j; k++) {
					// Skip the leading shape reference; the InlineArrayShape carries it once.
					const child = sideBuffers[k] ?? fail("missing side buffer");
					for (let m = 1; m < child.length; m++) {
						const item = child[m];
						if (item !== undefined) {
							buffer.push(item);
						}
					}
				}
			} else {
				for (let k = i; k < j; k++) {
					const child = sideBuffers[k] ?? fail("missing side buffer");
					for (const item of child) {
						buffer.push(item);
					}
				}
			}
			i = j;
		}

		outputBuffer.push(buffer);
	}
}

/**
 * Encodes the shape for an incremental chunk as {@link EncodedIncrementalChunkShape} shape.
 */
export class IncrementalChunkShape extends ShapeGeneric<EncodedChunkShapeV2> {
	public encodeShape(
		identifiers: DeduplicationTable<string>,
		shapes: DeduplicationTable<Shape>,
	): EncodedChunkShapeV2 {
		return {
			e: 0 /* EncodedIncrementalChunkShape */,
		};
	}

	public countReferencedShapesAndIdentifiers(
		identifiers: Counter<string>,
		shapeDiscovered: (shape: Shape) => void,
	): void {}

	public get shape(): this {
		return this;
	}
}

/**
 * Encodes an incremental field whose tree chunks are encoded separately and referenced by their {@link ChunkReferenceId}.
 * The shape of the content of this field is {@link NestedArrayShape}.
 * The inner items of the array have shape {@link IncrementalChunkShape} and are {@link ChunkReferenceId}s
 * of the encoded chunks.
 */
export const incrementalFieldEncoder: FieldEncoder = {
	encodeField(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		assert(
			context.incrementalEncoder !== undefined,
			0xc88 /* incremental encoder must be defined to use incrementalFieldEncoder */,
		);
		assert(
			supportsIncrementalEncoding(context.version),
			0xca1 /* Unsupported FieldBatchFormatVersion for incremental encoding; must be v2 or higher */,
		);

		const chunkReferenceIds = context.incrementalEncoder.encodeIncrementalField(
			cursor,
			(chunk: TreeChunk) => compressedEncode([chunk.cursor()], context),
		);
		outputBuffer.push(chunkReferenceIds);
	},

	shape: new NestedArrayShape(new IncrementalChunkShape() /* innerShape */),
};

/**
 * Encode `value` with `shape` into `outputBuffer`.
 *
 * Requires that `value` is compatible with `shape`.
 */
export function encodeValue(
	value: Value,
	shape: EncodedValueShape,
	outputBuffer: BufferFormat,
): void {
	if (shape === undefined) {
		if (value === undefined) {
			outputBuffer.push(false);
		} else {
			outputBuffer.push(true, value);
		}
	} else {
		if (shape === true) {
			assert(value !== undefined, 0x78d /* required value must not be missing */);
			outputBuffer.push(value);
		} else if (shape === false) {
			assert(value === undefined, 0x73f /* incompatible value shape: expected no value */);
		} else if (Array.isArray(shape)) {
			assert(shape.length === 1, 0x740 /* expected a single constant for value */);
		} else if (shape === SpecialField.Identifier) {
			// This case is a special case handling the encoding of identifier fields.
			assert(value !== undefined, 0x998 /* required value must not be missing */);
			outputBuffer.push(value);
		} else {
			// EncodedCounter case:
			unreachableCase(shape, "Encoding values as deltas is not yet supported");
		}
	}
}

/**
 * Provides common contextual information during encoding, like schema and policy settings.
 * Also, provides a cache to avoid duplicating equivalent shapes during a batch of encode operations.
 * @remarks
 * To avoid Shape duplication, any Shapes used in the encoding should either be:
 * - Singletons defined in a static scope.
 * - Cached in this object for future reuse such that all equivalent Shapes are deduplicated.
 */
export class EncoderContext implements NodeEncodeBuilder, FieldEncodeBuilder {
	private readonly nodeEncodersFromSchema: Map<TreeNodeSchemaIdentifier, NodeEncoder> =
		new Map();
	private readonly nestedArrayEncoders: Map<NodeEncoder, NestedArrayEncoder> = new Map();
	private readonly inlineArrayShapesForRun: Map<number, Map<Shape, InlineArrayEncoder>> =
		new Map();
	public constructor(
		private readonly nodeEncoderFromPolicy: NodeEncoderPolicy,
		private readonly fieldEncoderFromPolicy: FieldEncoderPolicy,
		public readonly fieldShapes: ReadonlyMap<FieldKindIdentifier, FieldKindData>,
		public readonly idCompressor: IIdCompressor,
		/**
		 * To be used to encode incremental chunks, if any.
		 * @remarks
		 * See {@link IncrementalEncoder} for more information.
		 */
		public readonly incrementalEncoder: IncrementalEncoder | undefined,
		public readonly version: FieldBatchFormatVersion,
		/**
		 * See {@link FieldBatchEncodingContext.isSummary}.
		 */
		public readonly isSummary: boolean,
		/**
		 * Optional hook invoked at the start of every {@link compressedEncode} call (including
		 * the recursive sub-chunk calls made by {@link incrementalFieldEncoder}). Used by
		 * encoder policies that need per-batch state — e.g. the VText two-pass encoder uses it
		 * to snapshot/restore per-batch specialization state and to run its counting pass.
		 * @remarks
		 * See {@link PreEncodeHook}.
		 */
		public readonly preEncodeHook: PreEncodeHook | undefined = undefined,
	) {}

	public nodeEncoderFromSchema(schemaName: TreeNodeSchemaIdentifier): NodeEncoder {
		return getOrCreate(this.nodeEncodersFromSchema, schemaName, () =>
			this.nodeEncoderFromPolicy(this, schemaName),
		);
	}

	public fieldEncoderFromSchema(fieldSchema: TreeFieldStoredSchema): FieldEncoder {
		return new LazyFieldEncoder(this, fieldSchema, this.fieldEncoderFromPolicy);
	}

	public nestedArrayEncoder(inner: NodeEncoder): NestedArrayEncoder {
		return getOrCreate(this.nestedArrayEncoders, inner, () => new NestedArrayEncoder(inner));
	}

	/**
	 * Returns an {@link InlineArrayEncoder} representing a run of `length` chunks of `shape`.
	 * Cached by `(length, shape)` so the resulting Shape is deduplicated in the shape table
	 * across all uses within a single encode batch — this amortizes the per-shape-table-entry
	 * cost when the same `(length, shape)` combination recurs (e.g. uniform-content fields).
	 * @remarks
	 * The returned encoder is intended for use as a Shape reference only — pushing it into a
	 * BufferFormat to be replaced by its index at finalization time. Its `encodeNodes` is
	 * unreachable; the data following the shape reference is provided directly by the caller.
	 */
	public inlineArrayShapeForRun(length: number, shape: Shape): InlineArrayEncoder {
		let byShape = this.inlineArrayShapesForRun.get(length);
		if (byShape === undefined) {
			byShape = new Map();
			this.inlineArrayShapesForRun.set(length, byShape);
		}
		let result = byShape.get(shape);
		if (result === undefined) {
			result = new InlineArrayEncoder(length, {
				shape,
				encodeNodes(): void {
					fail("InlineArrayEncoder for run detection should not be invoked");
				},
			});
			byShape.set(shape, result);
		}
		return result;
	}
}

export interface NodeEncodeBuilder {
	nodeEncoderFromSchema(schemaName: TreeNodeSchemaIdentifier): NodeEncoder;
}

export interface FieldEncodeBuilder {
	fieldEncoderFromSchema(schema: TreeFieldStoredSchema): FieldEncoder;
}

/**
 * The policy for building a {@link FieldEncoder} for a field.
 */
export type FieldEncoderPolicy = (
	nodeBuilder: NodeEncodeBuilder,
	schema: TreeFieldStoredSchema,
) => FieldEncoder;

/**
 * The policy for building a {@link NodeEncoder} for a node.
 */
export type NodeEncoderPolicy = (
	fieldBuilder: FieldEncodeBuilder,
	schemaName: TreeNodeSchemaIdentifier,
) => NodeEncoder;

/**
 * Hook for {@link EncoderContext.preEncodeHook}: invoked at the start of every
 * {@link compressedEncode} call. May optionally return a cleanup callback that is invoked
 * (in a `finally` block) after the encode completes — typically used to restore state
 * snapshotted at hook entry, so recursive sub-chunk encodes do not corrupt the outer
 * batch's state when they unwind.
 */
export type PreEncodeHook = (
	fieldBatch: FieldBatch,
	context: EncoderContext,
) => (() => void) | undefined;

class LazyFieldEncoder implements FieldEncoder {
	private encoderLazy: FieldEncoder | undefined;

	public constructor(
		public readonly nodeBuilder: NodeEncodeBuilder,
		public readonly fieldSchema: TreeFieldStoredSchema,
		private readonly fieldEncoderFromPolicy: FieldEncoderPolicy,
	) {}
	public encodeField(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		this.encoder.encodeField(cursor, context, outputBuffer);
	}

	private get encoder(): FieldEncoder {
		// eslint-disable-next-line @typescript-eslint/prefer-nullish-coalescing -- using ??= could change behavior if value is falsy
		if (this.encoderLazy === undefined) {
			this.encoderLazy = this.fieldEncoderFromPolicy(this.nodeBuilder, this.fieldSchema);
		}
		return this.encoderLazy;
	}

	public get shape(): Shape {
		return this.encoder.shape;
	}
}
