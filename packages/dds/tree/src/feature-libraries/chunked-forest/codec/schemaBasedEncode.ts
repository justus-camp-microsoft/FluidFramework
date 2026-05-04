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
	type Value,
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
	type Shape,
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
	type EncodedFieldBatchVTextExperimental,
	type EncodedValueShape,
	FieldBatchFormatVersion,
	SpecialField,
} from "./format/index.js";
import {
	defaultIncrementalEncodingPolicy,
	type IncrementalEncodingPolicy,
} from "./incrementalEncodingPolicy.js";
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
 * Uses the default specialization threshold {@link DEFAULT_MIN_OCCURRENCES_FOR_SPECIALIZATION}.
 */
export function schemaCompressedEncodeVTextExperimental(
	schema: StoredSchemaCollection,
	policy: SchemaPolicy,
	fieldBatch: FieldBatch,
	idCompressor: IIdCompressor,
	incrementalEncoder: IncrementalEncoder | undefined,
): EncodedFieldBatchVTextExperimental {
	return schemaCompressedEncodeVTextExperimentalForTests(
		schema,
		policy,
		fieldBatch,
		idCompressor,
		incrementalEncoder,
		DEFAULT_MIN_OCCURRENCES_FOR_SPECIALIZATION,
	);
}

/**
 * Test-only variant of {@link schemaCompressedEncodeVTextExperimental} that accepts a custom
 * `minOccurrencesForSpecialization` threshold so small test inputs can exercise the
 * specialization heuristic. Production callers must use
 * {@link schemaCompressedEncodeVTextExperimental}, which hard-codes
 * {@link DEFAULT_MIN_OCCURRENCES_FOR_SPECIALIZATION}.
 */
export function schemaCompressedEncodeVTextExperimentalForTests(
	schema: StoredSchemaCollection,
	policy: SchemaPolicy,
	fieldBatch: FieldBatch,
	idCompressor: IIdCompressor,
	incrementalEncoder: IncrementalEncoder | undefined,
	minOccurrencesForSpecialization: number,
): EncodedFieldBatchVTextExperimental {
	const context = buildContextVText(
		schema,
		policy,
		idCompressor,
		incrementalEncoder,
		brand(FieldBatchFormatVersion.vTextExperimental),
		minOccurrencesForSpecialization,
	);
	// `compressedEncode`'s return type is too narrow to express the vTextExperimental version
	// (its `version` field is `"text"`, not 1 or 2). The runtime shape matches
	// `EncodedFieldBatchVTextExperimental`; the double cast bridges the static type only.
	return compressedEncode(
		fieldBatch,
		context,
	) as unknown as EncodedFieldBatchVTextExperimental;
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
 * the outer batch does not actually emit. The policy is invoked with the same arguments
 * that {@link getNodeEncoder} uses, per the {@link IncrementalEncodingPolicy} contract:
 * per-field-key for ObjectNode/ArrayNode parents, per-node (with `fieldKey` undefined) for
 * MapNode/RecordNode parents. Each `forEachNode`/`forEachField` walk returns its cursor to
 * the starting position, so the same `fieldBatch` cursors are re-walkable by
 * `compressedEncode` in pass 2.
 */
function countVTextSpecializationCandidates(
	fieldBatch: FieldBatch,
	context: EncoderContext,
	storedSchema: StoredSchemaCollection,
): void {
	const shouldEncodeIncrementally = context.incrementalEncoder?.shouldEncodeIncrementally;
	for (const cursor of fieldBatch) {
		forEachNode(cursor, () => {
			countNodeAndDescendants(cursor, context, storedSchema, shouldEncodeIncrementally);
		});
	}
}

function countNodeAndDescendants(
	cursor: ITreeCursorSynchronous,
	context: EncoderContext,
	storedSchema: StoredSchemaCollection,
	shouldEncodeIncrementally: IncrementalEncodingPolicy | undefined,
): void {
	const nodeType: TreeNodeSchemaIdentifier = cursor.type;
	const schema = storedSchema.nodeSchema.get(nodeType);
	if (schema instanceof ObjectNodeStoredSchema) {
		// Object/Array: per-field policy decision. The cursor's field key is the object field
		// key for objects, or "" for arrays — both forms the contract accepts.
		forEachField(cursor, () => {
			if (shouldEncodeIncrementally?.(nodeType, cursor.getFieldKey()) === true) {
				return;
			}
			forEachNode(cursor, () => {
				countNodeAndDescendants(cursor, context, storedSchema, shouldEncodeIncrementally);
			});
		});
	} else if (schema instanceof MapNodeStoredSchema) {
		// Map/Record: per-node policy decision; the contract requires fieldKey to be undefined.
		// Mirrors the single shouldEncodeIncrementally(schemaName) call in getNodeEncoder.
		if (shouldEncodeIncrementally?.(nodeType) === true) {
			return;
		}
		forEachField(cursor, () => {
			forEachNode(cursor, () => {
				countNodeAndDescendants(cursor, context, storedSchema, shouldEncodeIncrementally);
			});
		});
	}
	// Leaf nodes (LeafNodeStoredSchema): no fields, no policy call. Unknown schemas would
	// have already failed `nodeEncoderFromSchema` below.
	//
	// Post-order: count this node after its descendants so leaf-value cohorts and inner
	// object cohorts have their counts and threshold decisions populated before this node's
	// cohort key is built. {@link VTextObjectNodeEncoder.countNode} computes its tuple via
	// `resolveShape` on each child encoder, which is only correct after the child has been
	// fully counted.
	const encoder = context.nodeEncoderFromSchema(nodeType);
	if (encoder instanceof VTextObjectNodeEncoder) {
		encoder.countNode(cursor);
	} else if (encoder instanceof VTextLeafNodeEncoder) {
		encoder.countNode(cursor);
	}
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
	const vTextLeafEncoders: VTextLeafNodeEncoder[] = [];
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
			} else if (encoder instanceof VTextLeafNodeEncoder) {
				vTextLeafEncoders.push(encoder);
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
			// `vTextEncoders` / `vTextLeafEncoders` after this point — they start with empty
			// state from their constructor, so they don't need a snapshot entry. On unwind,
			// snapshotted encoders are restored; new-this-batch encoders get their state
			// cleared so a later batch starts fresh.
			const objectSnapshots = new Map<VTextObjectNodeEncoder, VTextBatchState>();
			for (const enc of vTextEncoders) {
				objectSnapshots.set(enc, enc.swapBatchState());
			}
			const leafSnapshots = new Map<VTextLeafNodeEncoder, VTextLeafBatchState>();
			for (const enc of vTextLeafEncoders) {
				leafSnapshots.set(enc, enc.swapBatchState());
			}
			const restore = (): void => {
				for (const enc of vTextEncoders) {
					const snapshot = objectSnapshots.get(enc);
					if (snapshot === undefined) {
						enc.swapBatchState();
					} else {
						enc.restoreBatchState(snapshot);
					}
				}
				for (const enc of vTextLeafEncoders) {
					const snapshot = leafSnapshots.get(enc);
					if (snapshot === undefined) {
						enc.swapBatchState();
					} else {
						enc.restoreBatchState(snapshot);
					}
				}
			};
			// Per the {@link PreEncodeHook} contract, if counting throws after snapshotting
			// state, we must restore that state before rethrowing — otherwise the next encode
			// on this context starts with the snapshotted-empty state and the original
			// counts/specializedEncoders are lost forever.
			try {
				// Multi-pass count: each iteration's `resolveShape` reads counts populated by the
				// previous iteration, so cohort decisions propagate up the tree one nesting level
				// per pass. Iterates until counts stabilize, bounded by COUNT_PASS_MAX_ITERATIONS.
				// For text-domain schemas convergence is reached in ~2–4 passes; the bound exists
				// to defend against pathological cases (e.g. a deeply recursive schema where
				// thresholds keep flipping).
				for (let i = 0; i < COUNT_PASS_MAX_ITERATIONS; i++) {
					countVTextSpecializationCandidates(fieldBatch, ctx, storedSchema);
					let changed = false;
					for (const enc of vTextEncoders) {
						if (enc.commitCountIteration()) changed = true;
					}
					for (const enc of vTextLeafEncoders) {
						if (enc.commitCountIteration()) changed = true;
					}
					if (!changed) break;
				}
			} catch (error) {
				restore();
				throw error;
			}
			return restore;
		},
	);
	return context;
}

/**
 * Like {@link getNodeEncoder} but applies VText-specific specialization wrapping.
 *
 * @remarks
 * ObjectNodes with required single-valued boolean leaf fields are wrapped in a
 * {@link VTextObjectNodeEncoder} so those fields are constant-folded into specialized shapes.
 *
 * String/number LeafNodes are wrapped in a {@link VTextLeafNodeEncoder} so per-value cohorts
 * that occur often enough get their value constant-folded into a specialized shape.
 *
 * Booleans are excluded as leaf-node candidates: with only two possible values, a specialized
 * shape never beats the base shape's variable-boolean encoding once the AnyShape dispatch cost
 * is paid. Identifier leaves (special-cased value shape) are likewise excluded.
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

	if (schema instanceof LeafNodeStoredSchema) {
		if (schema.leafValue === ValueSchema.String || schema.leafValue === ValueSchema.Number) {
			assert(
				baseEncoder instanceof NodeShapeBasedEncoder,
				"VText leaf encoder policy expects NodeShapeBasedEncoder as base",
			);
			return new VTextLeafNodeEncoder(baseEncoder, minOccurrencesForSpecialization);
		}
		return baseEncoder;
	}

	const specializableFields: SpecializableField[] = [];
	if (schema instanceof ObjectNodeStoredSchema) {
		for (const [key, field] of schema.objectNodeFields ?? []) {
			if (context.fieldShapes.get(field.kind)?.multiplicity !== Multiplicity.Single) {
				continue;
			}
			// Defer to the caller's incremental policy: if a field is meant to be encoded
			// out-of-band, constant-folding its value into a specialized shape would silently
			// override that decision.
			if (incrementalEncoder?.shouldEncodeIncrementally?.(schemaName, key) === true) {
				continue;
			}
			const type = oneFromIterable(field.types);
			if (type === undefined) {
				// Polymorphic field: cohort key uses the resolved sub-shape per instance.
				// Cohorts only fire when all instances pick the same sub-shape (which implies
				// the same cursor.type), so the override pinning that shape is sound.
				specializableFields.push({ kind: "subShape", key });
				continue;
			}
			const nodeSchema = storedSchema.nodeSchema.get(type);
			if (
				nodeSchema instanceof LeafNodeStoredSchema &&
				(nodeSchema.leafValue === ValueSchema.Boolean ||
					nodeSchema.leafValue === ValueSchema.String ||
					nodeSchema.leafValue === ValueSchema.Number)
			) {
				specializableFields.push({ kind: "leafValue", key, leafType: type });
			} else {
				specializableFields.push({ kind: "subShape", key });
			}
		}
	}

	if (specializableFields.length === 0) {
		return baseEncoder;
	}

	assert(
		baseEncoder instanceof NodeShapeBasedEncoder,
		"VText node encoder policy expects NodeShapeBasedEncoder as base",
	);
	return new VTextObjectNodeEncoder(
		baseEncoder,
		specializableFields,
		minOccurrencesForSpecialization,
		context,
	);
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
 * Upper bound on iterations of the multi-pass VText count loop. Each pass propagates
 * threshold decisions one nesting level up: leaf cohorts settle in pass 1, immediate
 * parents in pass 2, and so on. For typical text-domain schemas this converges in
 * 2–4 passes; the cap is a safety net against pathological cases where thresholds
 * flip between passes (e.g. inputs that bunch right at the threshold in a recursive
 * schema). Hitting the cap doesn't break encoding — encode falls back to whatever
 * counts the last iteration produced — but it would mean some cohorts that *could*
 * have specialized end up using the base shape.
 */
const COUNT_PASS_MAX_ITERATIONS = 10;

/**
 * Encodes ObjectNodes using {@link SpecializedNodeShapeEncoder} shapes that constant-fold the
 * required single-valued fields whose contents are predictable across the cohort.
 *
 * Used in a two-pass encode: a counting pass ({@link countNode}) records each tuple's
 * occurrence count, after which the encoding pass ({@link encodeNode}) consults those counts
 * to decide on first sight whether to specialize. Tuples that occur at least
 * {@link MIN_OCCURRENCES_FOR_SPECIALIZATION} times use a specialized shape for *all* their
 * occurrences; rarer tuples encode through the base shape. The parent field uses
 * {@link AnyShape} dispatch so each node can carry its own shape index — paying one extra
 * dispatch token per node but saving one token per field that gets embedded as a constant
 * (or pinned to a child cohort shape) in the specialized shape.
 *
 * @remarks
 * Two flavors of field specialization are supported. `leafValue` covers required single-valued
 * boolean, string, or number leaf fields: the cohort key includes the field's value, and the
 * cohort's field override pins a constant-value shape that emits zero data per instance.
 * `subShape` covers required single-valued fields whose content is itself encoded by an
 * encoder that can produce different shapes per instance (a {@link VTextLeafNodeEncoder} or
 * {@link VTextObjectNodeEncoder}, transitively): the cohort key includes the resolved
 * sub-shape (by reference identity) and the cohort's field override pins that sub-shape, which
 * may itself be a specialized cohort emitting zero data — allowing wins to compound across
 * nesting levels.
 *
 * For correctness with nested composition, the count pass walks the tree in post-order: a
 * parent's cohort key is computed only after its children's counts are populated, so the
 * parent's `resolveShape` calls on its children return the cohort decisions those children
 * will make at encode time.
 */
type SpecializableField =
	| {
			readonly kind: "leafValue";
			readonly key: FieldKey;
			readonly leafType: TreeNodeSchemaIdentifier;
	  }
	| { readonly kind: "subShape"; readonly key: FieldKey };

/**
 * Per-batch state for a {@link VTextObjectNodeEncoder}. Snapshotted/restored by the VText
 * preEncodeHook so recursive sub-chunk encodes (via {@link incrementalFieldEncoder}) make
 * specialization decisions scoped to *their* batch without corrupting the outer batch.
 */
interface VTextBatchState {
	/**
	 * Counts being populated by the current count-pass iteration. After each iteration, copied
	 * to {@link resolveCounts} so the next iteration's `resolveShape` calls see them.
	 */
	counts: Map<string, number>;
	/**
	 * Counts from the previous count-pass iteration, consulted by {@link resolveShape} to make
	 * threshold decisions. Empty during the first iteration; equal to {@link counts} after the
	 * count loop converges, at which point the encode pass can read it as the final count map.
	 */
	resolveCounts: Map<string, number>;
	specializedEncoders: Map<string, SpecializedNodeShapeEncoder>;
}

function emptyVTextBatchState(): VTextBatchState {
	return { counts: new Map(), resolveCounts: new Map(), specializedEncoders: new Map() };
}

/**
 * Encodes ObjectNodes using {@link SpecializedNodeShapeEncoder} shapes that constant-fold
 * required single-valued boolean leaf fields.
 *
 * Used in a two-pass encode: a counting pass ({@link VTextObjectNodeEncoder.countNode}) records
 * each tuple's occurrence count, after which the encoding pass
 * ({@link VTextObjectNodeEncoder.encodeNode}) consults those counts to decide on first sight
 * whether to specialize. Tuples that occur at least `minOccurrencesForSpecialization` times
 * (default {@link DEFAULT_MIN_OCCURRENCES_FOR_SPECIALIZATION}) use a specialized shape for
 * *all* their occurrences; rarer tuples encode through the base shape. The parent field uses
 * {@link AnyShape} dispatch so each node can carry its own shape index — paying one extra
 * dispatch token per node but saving one token per boolean field that gets embedded as a
 * constant in the specialized shape.
 */
class VTextObjectNodeEncoder implements NodeEncoder {
	private readonly constantNodeEncoders: Map<string, NodeShapeBasedEncoder> = new Map();
	/**
	 * Stable per-encoder identifiers for {@link Shape} instances appearing as a `subShape`
	 * field's resolved shape. Used only to build a string cohort key — keying on the Shape
	 * directly would also work, but stringifying the tuple gives a single primitive key for
	 * the counts/encoder maps.
	 */
	private readonly shapeIds: Map<Shape, number> = new Map();
	private nextShapeId = 0;
	private batch: VTextBatchState = emptyVTextBatchState();

	public constructor(
		private readonly base: NodeShapeBasedEncoder,
		private readonly specializableFields: readonly SpecializableField[],
		private readonly minOccurrencesForSpecialization: number,
		private readonly nodeBuilder: NodeEncodeBuilder,
	) {}

	public get shape(): AnyShape {
		return AnyShape.instance;
	}

	/**
	 * Counting-pass entry point. Records this node's tuple key without producing output.
	 * Must be called *after* counting any children whose encoders contribute to this node's
	 * {@link SpecializableField} resolution — otherwise their `resolveShape` returns the base
	 * encoder before their own threshold has been observed, and the parent under-discriminates.
	 */
	public countNode(cursor: ITreeCursorSynchronous): void {
		const key = this.cohortKey(cursor);
		this.batch.counts.set(key, (this.batch.counts.get(key) ?? 0) + 1);
	}

	/**
	 * Replace the active batch state with a fresh empty one and return the previous state.
	 * Used by the VText preEncodeHook to snapshot/restore around recursive sub-chunk encodes.
	 * {@link constantNodeEncoders} and {@link shapeIds} are *not* swapped — they are stateless
	 * caches that are safe (and beneficial) to share across batches.
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
		const resolved = this.resolveShape(cursor);
		if (resolved instanceof SpecializedNodeShapeEncoder) {
			this.batch.specializedEncoders.set(this.cohortKey(cursor), resolved);
		}
		AnyShape.encodeNode(cursor, context, outputBuffer, resolved);
	}

	/**
	 * Returns the {@link NodeEncoder} this encoder will dispatch to for the cursor's current
	 * node: the cached specialized cohort if the node's tuple has crossed the threshold,
	 * otherwise the base encoder. Safe to call from a parent encoder — never mutates
	 * {@link batch.counts}, and caches its newly-created specialized instance in
	 * {@link batch.specializedEncoders} so subsequent calls (within the same iteration of the
	 * count loop, and in the encode pass) return the *same* shape reference. That stability
	 * is load-bearing: a parent encoder uses the resolved shape's identity to build its own
	 * cohort key via {@link idForShape}, so creating a fresh cohort instance per call would
	 * give every parent instance a unique key and prevent the parent cohort from firing.
	 */
	public resolveShape(
		cursor: ITreeCursorSynchronous,
	): NodeShapeBasedEncoder | SpecializedNodeShapeEncoder {
		const key = this.cohortKey(cursor);
		const existing = this.batch.specializedEncoders.get(key);
		if (existing !== undefined) {
			return existing;
		}
		if ((this.batch.resolveCounts.get(key) ?? 0) >= this.minOccurrencesForSpecialization) {
			const specialized = this.createSpecialized(cursor);
			this.batch.specializedEncoders.set(key, specialized);
			return specialized;
		}
		return this.base;
	}

	/**
	 * Promote {@link batch.counts} to {@link batch.resolveCounts} for the next count-pass
	 * iteration (or for the encode pass), and return whether the new counts differ from the
	 * previous resolveCounts — the multi-pass count loop uses this to detect a fixed point.
	 */
	public commitCountIteration(): boolean {
		const previous = this.batch.resolveCounts;
		const current = this.batch.counts;
		let changed = previous.size !== current.size;
		if (!changed) {
			for (const [k, v] of current) {
				if (previous.get(k) !== v) {
					changed = true;
					break;
				}
			}
		}
		this.batch.resolveCounts = new Map(current);
		this.batch.counts = new Map();
		return changed;
	}

	/**
	 * Build the cohort key for the node at the cursor's current position. Each
	 * {@link SpecializableField} contributes one segment:
	 *
	 * - `leafValue`: the field's leaf value (type-tagged via {@link valueKey}).
	 * - `subShape`: the per-encoder ID of the resolved child shape (reference identity).
	 */
	private cohortKey(cursor: ITreeCursorSynchronous): string {
		const parts: string[] = [];
		for (const field of this.specializableFields) {
			cursor.enterField(brand(field.key));
			const hasNode = cursor.firstNode();
			assert(hasNode, "required specializable field must contain a node");
			if (field.kind === "leafValue") {
				parts.push(`L:${valueKey(cursor.value)}`);
			} else {
				const childEncoder = this.nodeBuilder.nodeEncoderFromSchema(cursor.type);
				const childShape =
					childEncoder instanceof VTextLeafNodeEncoder ||
					childEncoder instanceof VTextObjectNodeEncoder
						? childEncoder.resolveShape(cursor)
						: childEncoder.shape;
				parts.push(`S:${this.idForShape(childShape)}`);
			}
			cursor.exitNode();
			cursor.exitField();
		}
		return parts.join("|");
	}

	private idForShape(shape: Shape): number {
		let id = this.shapeIds.get(shape);
		if (id === undefined) {
			id = this.nextShapeId++;
			this.shapeIds.set(shape, id);
		}
		return id;
	}

	private createSpecialized(cursor: ITreeCursorSynchronous): SpecializedNodeShapeEncoder {
		const overrides: KeyedFieldEncoder[] = [];
		for (const field of this.specializableFields) {
			cursor.enterField(brand(field.key));
			const hasNode = cursor.firstNode();
			assert(hasNode, "required specializable field must contain a node");
			if (field.kind === "leafValue") {
				const value = cursor.value;
				const cacheKey = `${field.leafType}:${valueKey(value)}`;
				let nodeEncoder = this.constantNodeEncoders.get(cacheKey);
				if (nodeEncoder === undefined) {
					nodeEncoder = new NodeShapeBasedEncoder(field.leafType, [value], [], undefined);
					this.constantNodeEncoders.set(cacheKey, nodeEncoder);
				}
				overrides.push({ key: field.key, encoder: asFieldEncoder(nodeEncoder) });
			} else {
				const childEncoder = this.nodeBuilder.nodeEncoderFromSchema(cursor.type);
				const resolvedChild =
					childEncoder instanceof VTextLeafNodeEncoder ||
					childEncoder instanceof VTextObjectNodeEncoder
						? childEncoder.resolveShape(cursor)
						: childEncoder;
				overrides.push({
					key: field.key,
					encoder: asFieldEncoder(resolvedChild),
				});
			}
			cursor.exitNode();
			cursor.exitField();
		}
		return new SpecializedNodeShapeEncoder(this.base, overrides);
	}
}

/**
 * Per-batch state for a {@link VTextLeafNodeEncoder}. Keyed on the encoded leaf value.
 * Snapshotted/restored by the VText preEncodeHook for the same reasons as
 * {@link VTextBatchState}.
 */
interface VTextLeafBatchState {
	counts: Map<string, number>;
	resolveCounts: Map<string, number>;
	specializedEncoders: Map<string, SpecializedNodeShapeEncoder>;
}

function emptyVTextLeafBatchState(): VTextLeafBatchState {
	return { counts: new Map(), resolveCounts: new Map(), specializedEncoders: new Map() };
}

/**
 * Encodes a string- or number-valued leaf node, producing a {@link SpecializedNodeShapeEncoder}
 * that constant-folds the leaf's value when that value occurs at least
 * {@link MIN_OCCURRENCES_FOR_SPECIALIZATION} times in the batch.
 *
 * Counting and dispatch mirror {@link VTextObjectNodeEncoder}: pass 1 records each value's
 * occurrence count, and pass 2 promotes each cohort that crosses the threshold to its own
 * specialized shape; rarer values fall back to the base shape. The parent uses {@link AnyShape}
 * dispatch so each leaf can carry its own shape index.
 *
 * The dispatch costs one extra token per leaf, while the specialized shape saves one token
 * per leaf (the variable value), so this is roughly a wash in the average case. The win is
 * concentrated in workloads where one value dominates an entire array — the parent could in
 * principle skip AnyShape and inline the specialized leaf shape, but that optimization is not
 * implemented here.
 */
class VTextLeafNodeEncoder implements NodeEncoder {
	private readonly constantNodeEncoders: Map<string, SpecializedNodeShapeEncoder> = new Map();
	private batch: VTextLeafBatchState = emptyVTextLeafBatchState();

	public constructor(
		private readonly base: NodeShapeBasedEncoder,
		private readonly minOccurrencesForSpecialization: number,
	) {}

	public get shape(): AnyShape {
		return AnyShape.instance;
	}

	/**
	 * Counting-pass entry point. Records this leaf's value occurrence without producing output.
	 */
	public countNode(cursor: ITreeCursorSynchronous): void {
		const key = valueKey(cursor.value);
		this.batch.counts.set(key, (this.batch.counts.get(key) ?? 0) + 1);
	}

	public swapBatchState(): VTextLeafBatchState {
		const previous = this.batch;
		this.batch = emptyVTextLeafBatchState();
		return previous;
	}

	public restoreBatchState(state: VTextLeafBatchState): void {
		this.batch = state;
	}

	public encodeNode(
		cursor: ITreeCursorSynchronous,
		context: EncoderContext,
		outputBuffer: BufferFormat,
	): void {
		const resolved = this.resolveShape(cursor);
		if (resolved instanceof SpecializedNodeShapeEncoder) {
			this.batch.specializedEncoders.set(valueKey(cursor.value), resolved);
		}
		AnyShape.encodeNode(cursor, context, outputBuffer, resolved);
	}

	/**
	 * Returns the {@link NodeEncoder} this encoder will dispatch to for the cursor's current
	 * value: the cached specialized cohort if the value's count has crossed the threshold,
	 * otherwise the base encoder. Safe to call from a parent encoder's count pass — it never
	 * mutates {@link batch.counts}, and the cached specialized instance is reused at encode time.
	 */
	public resolveShape(
		cursor: ITreeCursorSynchronous,
	): NodeShapeBasedEncoder | SpecializedNodeShapeEncoder {
		const value = cursor.value;
		const key = valueKey(value);
		const existing = this.batch.specializedEncoders.get(key);
		if (existing !== undefined) {
			return existing;
		}
		if ((this.batch.resolveCounts.get(key) ?? 0) >= this.minOccurrencesForSpecialization) {
			return this.getOrCreateSpecialized(key, value);
		}
		return this.base;
	}

	/**
	 * Promote {@link batch.counts} to {@link batch.resolveCounts} for the next count-pass
	 * iteration (or the encode pass), and return whether the new counts differ from the
	 * previous resolveCounts — the multi-pass count loop uses this to detect a fixed point.
	 */
	public commitCountIteration(): boolean {
		const previous = this.batch.resolveCounts;
		const current = this.batch.counts;
		let changed = previous.size !== current.size;
		if (!changed) {
			for (const [k, v] of current) {
				if (previous.get(k) !== v) {
					changed = true;
					break;
				}
			}
		}
		this.batch.resolveCounts = new Map(current);
		this.batch.counts = new Map();
		return changed;
	}

	private getOrCreateSpecialized(key: string, value: Value): SpecializedNodeShapeEncoder {
		const cached = this.constantNodeEncoders.get(key);
		if (cached !== undefined) {
			return cached;
		}
		const specialized = new SpecializedNodeShapeEncoder(this.base, [], { value: [value] });
		this.constantNodeEncoders.set(key, specialized);
		return specialized;
	}
}

/**
 * Encodes a leaf value to a string suitable for use as a Map key. Strings, numbers, and
 * booleans are unambiguous when prefixed with their type tag.
 *
 * {@link VTextLeafNodeEncoder} is only wired up for string/number leaves; other value types
 * shouldn't reach this code path.
 */
function valueKey(value: Value): string {
	const t = typeof value;
	assert(
		t === "string" || t === "number" || t === "boolean",
		"VTextLeafNodeEncoder only supports primitive leaf values",
	);
	return `${t}:${value as string | number | boolean}`;
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
