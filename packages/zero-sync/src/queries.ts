import { defineQueriesWithType, defineQuery } from "@rocicorp/zero";
import { z } from "zod";
import { schema } from "./schema.js";
import { zql } from "./zql.js";

const defineQueries = defineQueriesWithType<typeof schema>();

export const workbookQueryArgsSchema = z.object({
  documentId: z.string().min(1),
});

export const queries = defineQueries({
  workbooks: {
    byId: defineQuery(workbookQueryArgsSchema, ({ args: { documentId } }) =>
      zql.workbooks
        .where("id", documentId)
        .related("sheets", (sheet) =>
          sheet
            .orderBy("sortOrder", "asc")
            .related("cells", (cell) => cell.orderBy("address", "asc"))
            .related("computedCells", (cell) => cell.orderBy("address", "asc"))
            .related("rowMetadata", (entry) => entry.orderBy("startIndex", "asc"))
            .related("columnMetadata", (entry) => entry.orderBy("startIndex", "asc")),
        )
        .related("definedNames", (entry) => entry.orderBy("name", "asc"))
        .related("workbookMetadataEntries", (entry) => entry.orderBy("key", "asc"))
        .related("calculationSettings", (entry) => entry.one())
        .one(),
    ),
  },
});
