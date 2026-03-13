declare module "react-reconciler" {
  const Reconciler: (hostConfig: unknown) => any;
  export default Reconciler;
}

declare module "react-reconciler/constants" {
  export const DefaultEventPriority: number;
}
