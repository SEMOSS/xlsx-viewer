/**
 * Copy / fill translation of formulas (token based, see `language/translate.ts`):
 * `translateFormula(formula, rowDelta, colDelta)` shifts relative references,
 * keeps absolute parts and turns references pushed off the grid into `#REF!`.
 */
export { translateFormula } from './language/translate';
