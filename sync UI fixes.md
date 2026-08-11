# Sync UI Fixes

Windows-side running checklist for LUCAS Mac/Windows UI and behavior changes that need to be mirrored without drifting.

## Entry Template

- Origin: Mac or Windows
- Implemented on origin: Yes/No
- Mirrored to other platform: Yes/No/Pending/Not needed
- Source commit/repo:
- Target to mirror:
- Summary:
- Notes / avoid porting:

## 2026-08-11 - Mac People Rules / Network Mode Payout UI

- Origin: Mac
- Implemented on origin: Yes
- Mirrored to Windows: Pending
- Target to mirror: Windows

- Mac commit `f2cec03` - Added `Balance Share %` support to People Rules. Implemented on Mac: Yes. Mirrored to Windows: Pending.
  - Blank Sheet Type + Balance Share defines a team member payout share from unpaid net profit.
  - Fallback remains 50% when no explicit team balance share exists.

- Mac commit `db818af` - Superseded by later correction. Implemented on Mac: Yes, then corrected. Mirrored to Windows: Do not port.
  - Incorrectly treated balance-share-only rows as direct payout sources.
  - Do not port this behavior by itself.

- Mac commit `82e60ed` - Corrected People Rules balance-share semantics. Implemented on Mac: Yes. Mirrored to Windows: Pending.
  - Blank Sheet Type + Balance Share is a team profit-share rule, e.g. 50%.
  - Sheet Type + Seller Rate/Deduction is a Network Mode pass-through payout row.
  - Balance-share-only rows do not classify the person/source as seller payout.

- Mac commit `f758fcc` - Renamed People Rules `Seller` column to `Person`. Implemented on Mac: Yes. Mirrored to Windows: Pending.
  - Old CSV files with `Seller` header still load.
  - New saves write `Person`.
  - Network Mode Create UI label changed from Seller to Person.
  - Rows with Sheet Type clear Balance Share on save.

- Mac commit `e35fa07` - Disabled Balance Share input when Sheet Type is filled. Implemented on Mac: Yes. Mirrored to Windows: Pending.
  - User cannot type Balance Share while a Sheet Type is present.
  - Clearing Sheet Type re-enables Balance Share.
  - Save-time clearing remains as backup.

- Mac commit `691eee9` - Added blank Sheet Type dropdown option. Implemented on Mac: Yes. Mirrored to Windows: Pending.
  - Blank option appears first in People Rules Sheet Type dropdown.
  - Lets a row reset from company payout rule back to no Sheet Type.

### Current Intended Rules

- `Person` + blank `Sheet Type` + `Balance Share %` = team member unpaid net-profit share.
- `Person` + `Sheet Type` + `Seller Rate %` or `Deduction %` = Network Mode pass-through payout rule.
- `Balance Share %` and `Sheet Type` should not coexist in the same saved row.
- If `Sheet Type` is filled, `Balance Share %` should be disabled in the UI.
- If `Sheet Type` is reset to blank, `Balance Share %` should become editable again.
