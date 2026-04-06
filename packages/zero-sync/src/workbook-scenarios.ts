import { z } from "zod";

export const workbookScenarioViewportSchema = z
  .object({
    rowStart: z.number().int().nonnegative(),
    rowEnd: z.number().int().nonnegative(),
    colStart: z.number().int().nonnegative(),
    colEnd: z.number().int().nonnegative(),
  })
  .refine(
    (viewport) => viewport.rowEnd >= viewport.rowStart && viewport.colEnd >= viewport.colStart,
    {
      message: "viewport bounds must be ordered",
    },
  );

export const workbookScenarioCreateRequestSchema = z.object({
  name: z.string().trim().min(1),
  sheetName: z.string().trim().min(1).optional(),
  address: z.string().trim().min(1).optional(),
  viewport: workbookScenarioViewportSchema.optional(),
});

export const workbookScenarioResponseSchema = z.object({
  documentId: z.string().min(1),
  workbookId: z.string().min(1),
  ownerUserId: z.string().min(1),
  name: z.string().min(1),
  baseRevision: z.number().int().nonnegative(),
  sheetId: z.number().int().positive().nullable(),
  sheetName: z.string().nullable(),
  address: z.string().nullable(),
  viewport: workbookScenarioViewportSchema.nullable(),
  createdAt: z.number().int().nonnegative(),
  updatedAt: z.number().int().nonnegative(),
});

export const workbookScenarioDeleteResponseSchema = z.object({
  ok: z.literal(true),
});

export type WorkbookScenarioCreateRequest = z.infer<typeof workbookScenarioCreateRequestSchema>;
export type WorkbookScenarioResponse = z.infer<typeof workbookScenarioResponseSchema>;
export type WorkbookScenarioDeleteResponse = z.infer<typeof workbookScenarioDeleteResponseSchema>;
