# Fix Competitor Code KeyError

## What & Why
The app crashes with `KeyError: 'product'` when loading the product reference page. Auto-detected competitor codes from uploaded reports are saved with a `name` key, but the lookup builder expects a `product` key. This mismatch needs to be fixed in both the data writer and the data reader so neither old nor new entries cause errors.

## Done looks like
- The Product Reference page loads without errors, even when auto-detected competitor codes are present in the database
- Newly auto-detected competitor codes are saved with the correct key structure
- Existing entries with the old key structure continue to work

## Out of scope
- Migrating or rewriting existing `codes_db.json` entries (the reader should handle both formats)
- Changes to how Royal Purple or service tier codes are stored

## Tasks
1. In the lookup builder, make the competitor code reader resilient by falling back from `product` to `name` (or vice versa) so both old and new entries work.
2. In the auto-detection code, change the competitor code entry to use `product` as the key instead of `name`, matching the expected schema.

## Relevant files
- `product_reference.py:43-52`
- `code_detector.py:198-206`
- `codes_db.json`
