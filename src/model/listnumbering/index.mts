// CR-001 Phase B step 3: list numbering (docx4j `org.docx4j.model.listnumbering`, as CR-014
// settled it against Word).
//
// The split is the one docx4j-python's port uses, so that the two read alike:
//
//   definitions.mts  what a list *is*: `LevelDefinition`, `AbstractListDefinition`,
//                    `ListDefinition`, and `NumberingDefinitions` over one `w:numbering`
//   state.mts        what it is *up to*: `Counter`, `NumberingState` (one per story of one
//                    traversal), `NumberingStates`
//   formats.mts      the `w:numFmt` registry, fail-soft to decimal as Word is
//   Emulator.mts     the resolution (through `PropertyResolver`) and the counting
export {
  LEVELS, LevelDefinition, AbstractListDefinition, ListDefinition, NumberingDefinitions,
  type StyleLookup,
} from './definitions.mjs';
export {
  Counter, NumberingState, NumberingStates, storyKindOf, type PartLike, type StoryKind,
} from './state.mjs';
export {
  formatValue, formatterFor, register, registeredFormats, resetWarnings, type LabelFormatter,
} from './formats.mjs';
export {
  Emulator, NumRef, type NumberingResult, type NumberedPackageLike, type NumberingPartLike,
} from './Emulator.mjs';
