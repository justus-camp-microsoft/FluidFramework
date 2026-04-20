/*!
 * Copyright (c) Microsoft Corporation and contributors. All rights reserved.
 * Licensed under the MIT License.
 */

import { assert, unreachableCase, fail } from "@fluidframework/core-utils/internal";
import type { IIdCompressor } from "@fluidframework/id-compressor";

import {
	LeafNodeStoredSchema,
	MapNodeStoredSchema,
	ObjectNodeStoredSchema,
	type StoredSchemaCollection,
	type TreeFieldStoredSchema,
	type TreeNodeSchemaIdentifier,
	type FieldKey,
	type ITreeCursorSynchronous,
	ValueSchema,
	Multiplicity,
	identifierFieldKindIdentifier,
	type SchemaPolicy,
	forEachField,
	forEachNode,
} from "../../../core/index.js";
import { brand, oneFromIterable } from "../../../util/index.js";

import type { IncrementalEncoder } from "./codecs.js";
import {
	AnyShape,
	EncoderContext,
	type BufferFormat,
	type FieldEncoder,
	type FieldEncodeBuilder,
	type KeyedFieldEncoder,
	type NodeEncoder,
	type NodeEncodeBuilder,
	anyNodeEncoder,
	asFieldEncoder,
	compressedEncode,
	incrementalFieldEncoder,
} from "./compressedEncode.js";
import type { FieldBatch } from "./fieldBatch.js";
import {
	type EncodedFieldBatchV1,
	type EncodedFieldBatchV1OrV2,
	type EncodedFieldBatchV2,
	type EncodedValueShape,
	FieldBatchFormatVersion,
	SpecialField,
} from "./format/index.js";
import { defaultIncrementalEncodingPolicy } from "./incrementalEncodingPolicy.js";
import { NodeShapeBasedEncoder, SpecializedNodeShapeEncoder } from "./nodeEncoder.js";

/**
 * Encode data from `fieldBatch` in into an `EncodedChunk` using {@link FieldBatchFormatVersion.v1}.
 * @remarks See {@link schemaCompressedEncode} for more details.
 * This version does not support incremental encoding.
 */
export function schemaCompressedEncodeV1(
	schema: StoredSchemaCollection,
	policy: SchemaPolicy,
	fieldBatch: FieldBatch,
	idCompressor: IIdCompressor,
	_incrementalEncoder: IncrementalEncoder | undefined,
	isSummary: boolean,
): EncodedFieldBatchV1 {
	const encoded: EncodedFieldBatchV1OrV2 = schemaCompressedEncode(
		schema,
		policy,
		fieldBatch,
		idCompressor,
		undefined /* incrementalEncoder */,
		brand(FieldBatchFormatVersion.v1),
		isSummary,
	);
	// Since incrementalEncoder was not provided, no V2 features should be used, and this cast should be safe.
	return encoded as EncodedFieldBatchV1;
}

/**
 * Encode data from `fieldBatch` in into an `EncodedChunk` using {@link FieldBatchFormatVersion.v2}.
 * @remarks See {@link schemaCompressedEncode} for more details.
 * Incremental encoding is supported from this version onwards.
 */
export function schemaCompressedEncodeV2(
	schema: StoredSchemaCollection,
	policy: SchemaPolicy,
	fieldBatch: FieldBatch,
	idCompressor: IIdCompressor,
	incrementalEncoder: IncrementalEncoder | undefined,
	isSummary: boolean,
): EncodedFieldBatchV2 {
	return schemaCompressedEncode(
		schema,
		policy,
		fieldBatch,
		idCompressor,
		incrementalEncoder,
		brand(FieldBatchFormatVersion.v2),
		isSummary,
	);
}

/**
 * Encode data from `fieldBatch` in into an `EncodedChunk` using {@link FieldBatchFormatVersion.vTextExperimental}.
 * @remarks
 * Applies the specialized-node-shape (`f`) optimization: required single-valued boolean leaf
 * fields within ObjectNodes are constant-folded into the shape, reducing stream tokens for
 * nodes such as `CharacterFormat` where format flags are identical across many characters.
 * @param minOccurrencesForSpecialization - Minimum number of times a given boolean-value tuple
 * must appear in a single batch to be promoted to a specialized shape. Defaults to
 * {@link DEFAULT_MIN_OCCURRENCES_FOR_SPECIALIZATION}. Lowering this is primarily useful for
 * tests that want to exercise specialization with small inputs.
 */
export function schemaCompressedEncodeVTextExperimental(
	schema: StoredSchemaCollection,
	policy: SchemaPolicy,
	fieldBatch: FieldBatch,
	idCompressor: IIdCompressor,
	incrementalEncoder: IncrementalEncoder | undefined,
	minOccurrencesForSpecialization: number = DEFAULT_MIN_OCCURRENCES_FOR_SPECIALIZATION,
): EncodedFieldBatchV2 {
	const context = buildContextVText(
		schema,
		policy,
		idCompressor,
		incrementalEncoder,
		brand(FieldBatchFormatVersion.vTextExperimental),
		minOccurrencesForSpecialization,
	);
	return compressedEncode(fieldBatch, context) as EncodedFieldBatchV2;
}

/**
 * Pass 1 of the VText two-pass encode: walk every node in `fieldBatch` that will be encoded by
 * the matching pass-2 {@link compressedEncode} call, and for any node whose encoder is a
 * {@link VTextObjectNodeEncoder}, record its tuple occurrence so pass 2 can decide on first
 * sight whether the tuple has crossed the encoder's `minOccurrencesForSpecialization` threshold.
 * @remarks
 * Fields that the {@link IncrementalEncoder} will encode out-of-band are skipped: their
 * sub-chunks get their own count pass when `compressedEncode` is invoked recursively for
 * each chunk, so counting them here would inflate the outer batch's totals with nodes that
 * the outer batch does not actually emit. Each `forEachNode`/`forEachField` walk returns
 * its cursor to the starting position, so the same `fieldBatch` cursors are re-walkable by
 * `compressedEncode` in pass 2.
 */
function countVTextSpecializationCandidates(
	fieldBatch: FieldBatch,
	context: EncoderContext,
): void {
	const shouldEncodeIncrementally = context.incrementalEncoder?.shouldEncodeIncrementally;
	for (const cursor of fieldBatch) {
		countFieldUnlessIncremental(cursor, context, undefined, shouldEncodeIncrementally);
	}
}

function countFieldUnlessIncremental(
	cursor: ITreeCursorSynchronous,
	context: EncoderContext,
	parentNodeType: string | undefined,
	shouldEncodeIncrementally:
		| ((nodeType: string | undefined, fieldKey: string) => boolean)
		| undefined,
): void {
	if (shouldEncodeIncrementally?.(parentNodeType, cursor.getFieldKey()) === true) {
		return;
	}
	forEachNode(cursor, () => {
		const encoder = context.nodeEncoderFromSchema(cursor.type);
		if (encoder instanceof VTextObjectNodeEncoder) {
			encoder.countNode(cursor);
		}
		const nodeType: string = cursor.type;
		forEachField(cursor, () => {
			countFieldUnlessIncremental(cursor, context, nodeType, shouldEncodeIncrementally);
		});
	});
}

/**
 * Like {@link buildContext} but uses the VText-specific node encoder policy that produces
 * {@link SpecializedNodeShapeEncoder} shapes for ObjectNodes with boolean leaf fields.
 * @remarks
 * Wires up a per-batch {@link EncoderContext.preEncodeHook} that resets every
 * {@link VTextObjectNodeEncoder} created in this context and re-runs the counting pass over
 * the current batch — so each `compressedEncode` call (outer or recursive sub-chunk) makes
 * its specialization decisions using only the counts of nodes encoded in *that* batch.
 */
function buildContextVText(
	storedSchema: StoredSchemaCollection,
	policy: SchemaPolicy,
	idCompressor: IIdCompressor,
	incrementalEncoder: IncrementalEncoder | undefined,
	version: FieldBatchFormatVersion,
	minOccurrencesForSpecialization: number,
): EncoderContext {
	const vTextEncoders: VTextObjectNodeEncoder[] = [];
	const context: EncoderContext = new EncoderContext(
		(fieldBuilder: FieldEncodeBuilder, schemaName: TreeNodeSchemaIdentifier) => {
			const encoder = getNodeEncoderVText(
				fieldBuilder,
				storedSchema,
				schemaName,
				incrementalEncoder,
				context,
				minOccurrencesForSpecialization,
			);
			if (encoder instanceof VTextObjectNodeEncoder) {
				vTextEncoders.push(encoder);
			}
			return encoder;
		},
		(nodeBuilder: NodeEncodeBuilder, fieldSchema: TreeFieldStoredSchema) =>
			getFieldEncoder(nodeBuilder, fieldSchema, context, storedSchema),
		policy.fieldKinds,
		idCompressor,
		incrementalEncoder,
		version,
		(fieldBatch: FieldBatch, ctx: EncoderContext) => {
			// Snapshot every existing encoder's batch state, then run the count pass against
			// fresh state. New encoders created mid-batch (via lazy schema lookup) join
			// `vTextEncoders` after this point — they start with empty state from their
			// constructor, so they don't need a snapshot entry. On unwind, snapshotted
			// encoders are restored; new-this-batch encoders get their state cleared so a
			// later batch starts fresh.
			const snapshots = new Map<VTextObjectNodeEncoder, VTextBatchState>();
			for (const enc of vTextEncoders) {
				snapshots.set(enc, enc.swapBatchState());
			}
			countVTextSpecializationCandidates(fieldBatch, ctx);
			return () => {
				for (const enc of vTextEncoders) {
					const snapshot = snapshots.get(enc);
					if (snapshot === undefined) {
						enc.swapBatchState();
					} else {
						enc.restoreBatchState(snapshot);
					}
				}
			};
		},
	);
	return context;
}

/**
 * Like {@link getNodeEncoder} but wraps ObjectNodes that have required single-valued boolean
 * leaf fields in a {@link VTextObjectNodeEncoder} so those fields are constant-folded into
 * specialized shapes at encode time.
 */
function getNodeEncoderVText(
	fieldBuilder: FieldEncodeBuilder,
	storedSchema: StoredSchemaCollection,
	schemaName: TreeNodeSchemaIdentifier,
	incrementalEncoder: IncrementalEncoder | undefined,
	context: EncoderContext,
	minOccurrencesForSpecialization: number,
): NodeEncoder {
	const baseEncoder = getNodeEncoder(
		fieldBuilder,
		storedSchema,
		schemaName,
		incrementalEncoder,
	);

	const schema =
		storedSchema.nodeSchema.get(schemaName) ?? fail(0xb55 /* missing node schema */);
	const boolFields: { key: FieldKey; type: TreeNodeSchemaIdentifier }[] = [];
	if (schema instanceof ObjectNodeStoredSchema) {
		for (const [key, field] of schema.objectNodeFields ?? []) {
			const type = oneFromIterable(field.types);
			if (type === undefined) {
				continue;
			}
			const nodeSchema = storedSchema.nodeSchema.get(type);
			if (
				!(nodeSchema instanceof LeafNodeStoredSchema) ||
				nodeSchema.leafValue !== ValueSchema.Boolean
			) {
				continue;
			}
			if (context.fieldShapes.get(field.kind)?.multiplicity !== Multiplicity.Single) {
				continue;
			}
			boolFields.push({ key, type });
		}
	}

	if (boolFields.length === 0) {
		return baseEncoder;
	}

	assert(
		baseEncoder instanceof NodeShapeBasedEncoder,
		"VText node encoder policy expects NodeShapeBasedEncoder as base",
	);
	return new VTextObjectNodeEncoder(baseEncoder, boolFields, minOccurrencesForSpecialization);
}

/**
 * Default minimum number of occurrences of a given boolean-value tuple in a batch required
 * to generate a {@link SpecializedNodeShapeEncoder} for it. Tuples below this threshold
 * encode through the base {@link NodeShapeBasedEncoder}. Overridable per-call via the
 * `minOccurrencesForSpecialization` parameter on
 * {@link schemaCompressedEncodeVTextExperimental}.
 * @remarks
 * A specialized shape entry costs roughly `2 + 2 * fieldCount` tokens in the shape table;
 * each instance using the specialized shape saves `fieldCount` stream tokens. The encoder
 * runs a counting pass over the batch before encoding so that, when a tuple does cross the
 * threshold, *every* occurrence (including the first) uses the specialized shape.
 */
const DEFAULT_MIN_OCCURRENCES_FOR_SPECIALIZATION = 8;

/**
 * Encodes ObjectNodes using {@link SpecializedNodeShapeEncoder} shapes that constant-fold
 * required single-valued boolean leaf fields.
 *
 * Used in a two-pass encode: a counting pass ({@link countNode}) records each tuple's
 * occurrence count, after which the encoding pass ({@link encodeNode}) consults those counts
 * to decide on first sight whether to specialize. Tuples that occur at least
 * {@link MIN_OCCURRENCES_FOR_SPECIALIZATION} times use a specialized shape for *all* their
 * occurrences; rarer tuples encode through the base shape. The parent field uses
 * {@link AnyShape} dispatch so each node can carry its own shape index — paying one extra
 * dispatch token per node but saving one token per boolean field that gets embedded as a
 * constant in the specialized shape.
 */
/**
 * Per-batch state for a {@link VTextObjectNodeEncoder}. Snapshotted/restored by the VText
 * preEncodeHook so recursive sub-chunk encodes (via {@link incrementalFieldEncoder}) make
 * specialization decisions scoped to *their* batch without corrupting the outer batch.
 */
interface VTextBatchState {
	counts: Map<string, number>;
	specializedEncoders: Map<string, SpecializedNodeShapeEncoder>;
}

function emptyVTextBatchState(): VTextBatchState {
	return { counts: new Map(), specializedEncoders: new Map() };
}

class VTextObjectNodeEncoder implements NodeEncoder {
	private readonly constantNodeEncoders: Map<string, NodeShapeBasedEncoder> = new Map();
	private batch: VTextBatchState = emptyVTextBatchState();

	public constructor(
		private readonly base: NodeShapeBasedEncoder,
		private readonly boolFields: readonly { key: FieldKey; type: TreeNodeSchemaIdentifier }[],
		private readonly minOccurrencesForSpecialization: number,
	) {}

	public get shape(): AnyShape {
		return AnyShape.instance;
	}

	/**
	 * Counting-pass entry point. Records this node's tuple key without producing output.
	 */
	public countNode(cursor: ITreeCursorSynchronous): void {
		const key = this.readBoolValues(cursor).join(",");
		this.batch.counts.set(key, (this.batch.counts.get(key) ?? 0) + 1);
	}

	/**
	 * Replace the active batch state with a fresh empty one and return the previous state.
	 * Used by the VText preEncodeHook to snapshot/restore around recursive sub-chunk encodes.
	 * {@link constantNodeEncoders} is *not* swapped — it is a stateless leaf-shape cache that
	 * is safe (and beneficial) to share across batches.
	 */
	public swapBatchState(): VTextBatchState {
		const previous = this.batch;
		this.batch = emptyVTextBatchState();
		return previous;
	}

	/**
	 * Restore a batch state previously returned by {@link swapBatchState}.
	 */
	public restoreBatchState(state: VTextBatchState): void {
		this.batch = state;
	}

	public encodeNode(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		const boolValues = this.readBoolValues(cursor);
		const key = boolValues.join(",");
		let specialized = this.batch.specializedEncoders.get(key);
		if (
			specialized === undefined &&
			(this.batch.counts.get(key) ?? 0) >= this.minOccurrencesForSpecialization
		) {
			specialized = this.createSpecialized(boolValues);
			this.batch.specializedEncoders.set(key, specialized);
		}
		AnyShape.encodeNode(cursor, context, outputBuffer, specialized ?? this.base);
	}

	private readBoolValues(cursor: ITreeCursorSynchronous): boolean[] {
		return this.boolFields.map(({ key }) => {
			cursor.enterField(brand(key));
			const hasNode = cursor.firstNode();
			assert(hasNode, "required boolean field must contain a node");
			const value = cursor.value as boolean;
			cursor.exitNode();
			cursor.exitField();
			return value;
		});
	}

	private createSpecialized(boolValues: readonly boolean[]): SpecializedNodeShapeEncoder {
		const overrides: KeyedFieldEncoder[] = [];
		for (const [i, { key, type }] of this.boolFields.entries()) {
			const value = boolValues[i] as boolean;
			const cacheKey = `${type}:${value}`;
			let nodeEncoder = this.constantNodeEncoders.get(cacheKey);
			if (nodeEncoder === undefined) {
				nodeEncoder = new NodeShapeBasedEncoder(type, [value], [], undefined);
				this.constantNodeEncoders.set(cacheKey, nodeEncoder);
			}
			overrides.push({ key, encoder: asFieldEncoder(nodeEncoder) });
		}
		return new SpecializedNodeShapeEncoder(this.base, overrides);
	}
}

/**
 * Encode data from `fieldBatch` in into an `EncodedChunk`.
 * @remarks
 * If `incrementalEncoder` is provided,
 * fields that support incremental encoding will encode their chunks separately via the `incrementalEncoder`.
 * See {@link IncrementalEncoder} for more details.
 *
 * Optimized for encoded size and encoding performance.
 * TODO: This function should eventually also take in the root FieldSchema to more efficiently compress the nodes.
 */
function schemaCompressedEncode(
	schema: StoredSchemaCollection,
	policy: SchemaPolicy,
	fieldBatch: FieldBatch,
	idCompressor: IIdCompressor,
	incrementalEncoder: IncrementalEncoder | undefined,
	version: FieldBatchFormatVersion,
	isSummary: boolean,
): EncodedFieldBatchV1OrV2 {
	return compressedEncode(
		fieldBatch,
		buildContext(schema, policy, idCompressor, incrementalEncoder, version, isSummary),
	);
}

export function buildContext(
	storedSchema: StoredSchemaCollection,
	policy: SchemaPolicy,
	idCompressor: IIdCompressor,
	incrementalEncoder: IncrementalEncoder | undefined,
	version: FieldBatchFormatVersion,
	isSummary: boolean,
): EncoderContext {
	const context: EncoderContext = new EncoderContext(
		(fieldBuilder: FieldEncodeBuilder, schemaName: TreeNodeSchemaIdentifier) =>
			getNodeEncoder(fieldBuilder, storedSchema, schemaName, incrementalEncoder),
		(nodeBuilder: NodeEncodeBuilder, fieldSchema: TreeFieldStoredSchema) =>
			getFieldEncoder(nodeBuilder, fieldSchema, context, storedSchema),
		policy.fieldKinds,
		idCompressor,
		incrementalEncoder,
		version,
		isSummary,
	);
	return context;
}

/**
 * Selects an encoder to use to encode fields.
 */
export function getFieldEncoder(
	nodeBuilder: NodeEncodeBuilder,
	field: TreeFieldStoredSchema,
	context: EncoderContext,
	storedSchema: StoredSchemaCollection,
): FieldEncoder {
	const kind = context.fieldShapes.get(field.kind) ?? fail(0xb52 /* missing FieldKind */);
	const type = oneFromIterable(field.types);
	const nodeEncoder =
		type === undefined ? anyNodeEncoder : nodeBuilder.nodeEncoderFromSchema(type);
	if (kind.multiplicity === Multiplicity.Single) {
		if (field.kind === identifierFieldKindIdentifier) {
			assert(type !== undefined, 0x999 /* field type must be defined in identifier field */);
			const nodeSchema = storedSchema.nodeSchema.get(type);
			assert(nodeSchema !== undefined, 0x99a /* nodeSchema must be defined */);
			assert(
				nodeSchema instanceof LeafNodeStoredSchema,
				0x99b /* nodeSchema must be LeafNodeStoredSchema */,
			);
			assert(
				nodeSchema.leafValue === ValueSchema.String,
				0x99c /* identifier field can only be type string */,
			);
			const identifierNodeEncoder = new NodeShapeBasedEncoder(
				type,
				SpecialField.Identifier,
				[],
				undefined,
			);
			return asFieldEncoder(identifierNodeEncoder);
		}
		return asFieldEncoder(nodeEncoder);
	} else {
		return context.nestedArrayEncoder(nodeEncoder);
	}
}

/**
 * Selects an encoder to use to encode nodes.
 */
export function getNodeEncoder(
	fieldBuilder: FieldEncodeBuilder,
	storedSchema: StoredSchemaCollection,
	schemaName: TreeNodeSchemaIdentifier,
	incrementalEncoder?: IncrementalEncoder,
): NodeShapeBasedEncoder {
	const shouldEncodeIncrementally =
		incrementalEncoder?.shouldEncodeIncrementally ?? defaultIncrementalEncodingPolicy;
	const schema =
		storedSchema.nodeSchema.get(schemaName) ?? fail(0xb53 /* missing node schema */);

	// This handles both object and array nodes.
	if (schema instanceof ObjectNodeStoredSchema) {
		// TODO:Performance:
		// consider moving some optional and sequence fields to extra fields if they are commonly empty
		// to reduce encoded size.
		const objectNodeFields: KeyedFieldEncoder[] = [];
		for (const [key, field] of schema.objectNodeFields ?? []) {
			const fieldEncoder = shouldEncodeIncrementally(schemaName, key)
				? incrementalFieldEncoder
				: fieldBuilder.fieldEncoderFromSchema(field);
			objectNodeFields.push({
				key,
				encoder: fieldEncoder,
			});
		}

		const shape = new NodeShapeBasedEncoder(schemaName, false, objectNodeFields, undefined);
		return shape;
	}
	if (schema instanceof LeafNodeStoredSchema) {
		const shape = new NodeShapeBasedEncoder(
			schemaName,
			valueShapeFromSchema(schema.leafValue),
			[],
			undefined,
		);
		return shape;
	}

	// This handles both maps and record nodes.
	if (schema instanceof MapNodeStoredSchema) {
		const fieldEncoder = shouldEncodeIncrementally(schemaName)
			? incrementalFieldEncoder
			: fieldBuilder.fieldEncoderFromSchema(schema.mapFields);
		const shape = new NodeShapeBasedEncoder(schemaName, false, [], fieldEncoder);
		return shape;
	}
	fail(0xb54 /* unsupported node kind */);
}

function valueShapeFromSchema(schema: ValueSchema | undefined): undefined | EncodedValueShape {
	switch (schema) {
		case undefined: {
			return false;
		}
		case ValueSchema.Number:
		case ValueSchema.String:
		case ValueSchema.Boolean:
		case ValueSchema.FluidHandle: {
			return true;
		}
		case ValueSchema.Null: {
			return [null];
		}
		default: {
			unreachableCase(schema);
		}
	}
}
