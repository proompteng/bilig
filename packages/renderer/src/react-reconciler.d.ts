declare module "react-reconciler" {
  interface WorkbookReconcilerInstance {
    createContainer(...args: unknown[]): unknown;
    updateContainer(...args: unknown[]): void;
  }

  const Reconciler: (hostConfig: unknown) => WorkbookReconcilerInstance;
  export default Reconciler;
}

declare module "react-reconciler/constants" {
  export const DefaultEventPriority: number;
}
