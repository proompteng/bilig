import { directLookupColumnInfo } from './formula-binding-dependency-helpers.js'
import { rewriteDirectAggregateDescriptorForStructuralTransform } from './formula-binding-direct-descriptors.js'
import { staticIntegerValue } from './formula-binding-lookup-candidates.js'
import {
  directAggregateStructureEqual,
  directCriteriaOperandEqual,
  directCriteriaStructureEqual,
  directLookupStructureEqual,
  directScalarDependencyCellsEqual,
} from './formula-binding-shape-helpers.js'

export const formulaBindingServiceTestHooks = {
  directAggregateStructureEqual,
  directCriteriaOperandEqual,
  directCriteriaStructureEqual,
  directLookupColumnInfo,
  directLookupStructureEqual,
  directScalarDependencyCellsEqual,
  rewriteDirectAggregateDescriptorForStructuralTransform,
  staticIntegerValue,
}
