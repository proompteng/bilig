export interface WorkbookNavigationTarget {
  readonly documentId: string;
  readonly serverUrl?: string | null;
  readonly sheetName?: string | null;
  readonly address?: string | null;
}

export function resolveWorkbookNavigationUrl(input: WorkbookNavigationTarget): string {
  const url = new URL(window.location.href);
  url.searchParams.set("document", input.documentId);
  if (input.serverUrl) {
    url.searchParams.set("server", input.serverUrl);
  } else {
    url.searchParams.delete("server");
  }
  if (input.sheetName) {
    url.searchParams.set("sheet", input.sheetName);
  } else {
    url.searchParams.delete("sheet");
  }
  if (input.address) {
    url.searchParams.set("cell", input.address.toUpperCase());
  } else {
    url.searchParams.delete("cell");
  }
  return url.toString();
}
