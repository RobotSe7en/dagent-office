import { z } from "zod";
import { IdentifierSchema, PackagePathSchema } from "./primitives.js";
import { AnchorSchema, AnnotationStatusSchema } from "./annotations.js";
import { PresentationGeometrySchema, PresentationPresetSchema, PresentationSizeSchema } from "./presentation.js";
import { PresentationObjectSchema, PresentationChartSchema } from "./presentation.js";
/** Creation inputs are transient commands; core allocates all persisted identities. */
export const PresentationObjectInputSchema = z.discriminatedUnion("kind", [
    z.strictObject({ kind: z.literal("text"), text: z.string(), style: z.string().optional() }),
    z.strictObject({ kind: z.literal("geometry"), preset: PresentationPresetSchema, text: z.string().optional(), style: z.string().optional() }),
    z.strictObject({ kind: z.literal("image"), assetPath: PackagePathSchema, alt: z.string(), style: z.string().optional() }),
    z.strictObject({ kind: z.literal("connector"), routing: z.enum(["straight", "elbow"]),
        startNodeId: IdentifierSchema.optional(), startSite: z.number().int().min(0).optional(), endNodeId: IdentifierSchema.optional(), endSite: z.number().int().min(0).optional(), style: z.string().optional() }),
    z.strictObject({ kind: z.literal("table"), rows: z.number().int().positive().max(200), columns: z.number().int().positive().max(200), style: z.string().optional() }),
]);
export type PresentationObjectInput = z.infer<typeof PresentationObjectInputSchema>;
const selection = z.array(IdentifierSchema).min(1);
const index = z.number().int().min(0), count = z.number().int().min(1).max(200);
const definitionKind = z.enum(["theme", "master", "layout", "notes-master"]);
const chartData = z.strictObject({ categories: z.array(z.string()), series: z.array(z.strictObject({ seriesId: IdentifierSchema.optional(), name: z.string(), values: z.array(z.number().nullable()) })).min(1) });
const author = z.strictObject({ name: z.string(), initials: z.string().optional() });
export const PresentationTableChangeSchema = z.discriminatedUnion("kind", [
    z.strictObject({ kind: z.enum(["insertRows", "deleteRows", "insertColumns", "deleteColumns"]), at: index, count }),
    z.strictObject({ kind: z.literal("mergeCells"), row: index, column: index, rows: count, columns: count }),
    z.strictObject({ kind: z.literal("splitCell"), row: index, column: index }),
    z.strictObject({ kind: z.literal("resize"), grid: z.array(z.number().positive()).min(1).max(200).optional(), rowHeights: z.array(z.number().positive()).min(1).max(200).optional() }),
    z.strictObject({ kind: z.literal("setCells"), row: index, column: index, values: z.array(z.array(z.string()).min(1).max(200)).min(1).max(200) }),
]);
/** Format commands remain separate from the V1 Word proposal operation set. */
export const PresentationOperationSchema = z.discriminatedUnion("op", [
    z.strictObject({ op: z.literal("setPresentationGeometry"), nodeId: IdentifierSchema, geometry: PresentationGeometrySchema,
        childCoordinates: z.strictObject({ x: z.number(), y: z.number(), width: z.number().positive(), height: z.number().positive() }).optional() }),
    z.strictObject({ op: z.literal("setPresentationSize"), size: PresentationSizeSchema }),
    z.strictObject({ op: z.literal("insertSlide"), layoutId: IdentifierSchema, afterSlideId: IdentifierSchema.optional() }),
    z.strictObject({ op: z.literal("duplicateSlide"), slideId: IdentifierSchema, afterSlideId: IdentifierSchema.optional() }),
    z.strictObject({ op: z.literal("removeSlide"), slideId: IdentifierSchema }),
    z.strictObject({ op: z.literal("moveSlide"), slideId: IdentifierSchema, beforeSlideId: IdentifierSchema.optional() }),
    z.strictObject({ op: z.literal("setSlideProperties"), slideId: IdentifierSchema, changes: z.strictObject({ name: z.string().optional(), hidden: z.boolean().optional() }) }),
    z.strictObject({ op: z.literal("insertPresentationObject"), parentNodeId: IdentifierSchema, beforeNodeId: IdentifierSchema.optional(), geometry: PresentationGeometrySchema, object: PresentationObjectInputSchema }),
    z.strictObject({ op: z.literal("duplicatePresentationObjects"), nodeIds: selection, offset: z.strictObject({ x: z.number(), y: z.number() }) }),
    z.strictObject({ op: z.literal("reorderPresentationObjects"), nodeIds: selection, position: z.enum(["front", "back", "forward", "backward"]) }),
    z.strictObject({ op: z.literal("alignPresentationObjects"), nodeIds: selection, alignment: z.enum(["left", "center", "right", "top", "middle", "bottom"]), relativeTo: z.enum(["selection", "slide"]) }),
    z.strictObject({ op: z.literal("distributePresentationObjects"), nodeIds: selection, axis: z.enum(["horizontal", "vertical"]) }),
    z.strictObject({ op: z.literal("groupPresentationObjects"), nodeIds: selection }),
    z.strictObject({ op: z.literal("ungroupPresentationObject"), nodeId: IdentifierSchema }),
    z.strictObject({ op: z.literal("setPresentationCrop"), nodeId: IdentifierSchema, crop: z.strictObject({ left: z.number().min(0).lt(1), top: z.number().min(0).lt(1), right: z.number().min(0).lt(1), bottom: z.number().min(0).lt(1) }) }),
    z.strictObject({ op: z.literal("setPresentationPreset"), nodeId: IdentifierSchema, preset: PresentationPresetSchema }),
    z.strictObject({ op: z.literal("setPresentationConnector"), nodeId: IdentifierSchema, changes: z.strictObject({ routing: z.enum(["straight", "elbow"]).optional(),
            start: z.union([z.null(), z.strictObject({ nodeId: IdentifierSchema, site: z.number().int().min(0) })]).optional(),
            end: z.union([z.null(), z.strictObject({ nodeId: IdentifierSchema, site: z.number().int().min(0) })]).optional() }) }),
    z.strictObject({ op: z.literal("insertPresentationNotes"), slideId: IdentifierSchema, html: z.string(), notesMasterId: IdentifierSchema.optional() }),
    z.strictObject({ op: z.literal("setPresentationNotesMaster"), slideId: IdentifierSchema, notesMasterId: IdentifierSchema.nullable() }),
    z.strictObject({ op: z.literal("editPresentationTable"), nodeId: IdentifierSchema, change: PresentationTableChangeSchema }),
    z.strictObject({ op: z.literal("setPresentationTheme"), themeId: IdentifierSchema, properties: z.record(z.string(), z.string().min(1)) }),
    z.strictObject({ op: z.literal("insertPresentationDefinition"), kind: definitionKind, copyFromId: IdentifierSchema.optional(), themeId: IdentifierSchema.optional(), masterId: IdentifierSchema.optional(), name: z.string().optional() }),
    z.strictObject({ op: z.literal("updatePresentationDefinition"), kind: z.enum(["master", "layout", "notes-master"]), definitionId: IdentifierSchema, changes: z.strictObject({ themeId: IdentifierSchema.optional(), masterId: IdentifierSchema.optional(), name: z.string().optional(), showMasterShapes: z.boolean().optional() }) }),
    z.strictObject({ op: z.literal("removePresentationDefinition"), kind: definitionKind, definitionId: IdentifierSchema, replacementId: IdentifierSchema.optional() }),
    z.strictObject({ op: z.literal("setPresentationSlideLayout"), slideId: IdentifierSchema, layoutId: IdentifierSchema }),
    z.strictObject({ op: z.literal("setPresentationPlaceholder"), nodeId: IdentifierSchema, placeholder: PresentationObjectSchema.options[0].shape.placeholder.unwrap().nullable() }),
    z.strictObject({ op: z.literal("resetPresentationOverrides"), nodeId: IdentifierSchema, geometry: z.boolean().optional(), appearance: z.boolean().optional(), text: z.boolean().optional() }),
    z.strictObject({ op: z.literal("insertPresentationChart"), parentNodeId: IdentifierSchema, geometry: PresentationGeometrySchema, kind: z.enum(["column", "line", "pie"]), title: z.string().optional(), data: chartData }),
    z.strictObject({ op: z.literal("setPresentationChart"), nodeId: IdentifierSchema, changes: PresentationChartSchema.omit({ chartId: true, nodeId: true, dataSetId: true, titleNodeId: true, seriesStyles: true, pointStyles: true }).partial(), title: z.string().nullable().optional(), clear: z.array(z.enum(["minimum", "maximum", "majorUnit", "numberFormat", "labelContents", "seriesStyles", "pointStyles"])).optional() }),
    z.strictObject({ op: z.literal("setPresentationChartData"), nodeId: IdentifierSchema, data: chartData }),
    z.strictObject({ op: z.literal("setPresentationChartStyle"), nodeId: IdentifierSchema, target: z.enum(["plot", "legend", "text", "label", "categoryAxis", "valueAxis", "grid", "series", "point"]), seriesId: IdentifierSchema.optional(), index: z.number().int().min(0).optional(), properties: z.record(z.string(), z.string().min(1)) }),
    z.strictObject({ op: z.literal("addPresentationComment"), anchor: AnchorSchema, html: z.string(), format: z.enum(["classic", "modern"]).optional(), author: author.optional(), date: z.string().optional() }),
    z.strictObject({ op: z.literal("addPresentationCommentReply"), annotationId: IdentifierSchema, html: z.string(), author: author.optional(), date: z.string().optional() }),
    z.strictObject({ op: z.literal("removePresentationCommentReply"), annotationId: IdentifierSchema, replyId: IdentifierSchema }),
    z.strictObject({ op: z.literal("setPresentationCommentBody"), annotationId: IdentifierSchema, replyId: IdentifierSchema.optional(), html: z.string() }),
    z.strictObject({ op: z.literal("updatePresentationComment"), annotationId: IdentifierSchema, changes: z.strictObject({ anchor: AnchorSchema.optional(), status: AnnotationStatusSchema.optional(), format: z.enum(["classic", "modern"]).optional(), author: author.optional(), date: z.string().optional() }) }),
    z.strictObject({ op: z.literal("setPresentationTextProperties"), nodeId: IdentifierSchema, changes: z.strictObject({
            autoFit: z.enum(["none", "normal", "shape"]).nullable().optional(), wrap: z.boolean().nullable().optional(),
            fontScale: z.number().positive().max(1).nullable().optional(), lineSpacingReduction: z.number().min(0).max(1).nullable().optional(),
        }) }),
]);
export type PresentationOperation = z.infer<typeof PresentationOperationSchema>;
