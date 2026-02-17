# Contributing to Excel Macro VBA Library

Thanks for considering contributing. These guidelines help keep the library consistent and easy to use.

## How to contribute

1. **Fork and clone** the repo.
2. **Add or edit** VBA in your Excel workbook:
   - Prefer **one logical module per file** (e.g. one `.bas` for sheet helpers, one for AD, etc.), or add to `Module1.bas` if it’s a small addition.
   - Use **meaningful names**: `Custom_` prefix for subs/functions that operate on sheets/ranges; no prefix for generic helpers like `ProperX`, `ReturnNthPartOfString`.
3. **Export** the module from the VBA editor (Right-click module → Export File) and save as `.bas` or `.vb` in the repo.
4. **Document** your code:
   - Add a short comment block at the top of each **Public** Sub/Function: what it does, main parameters, and one example if helpful.
   - Add or update the entry in **docs/API.md** (signature, parameters, brief description, example).
5. **Update README.md** if you add a new file or a new category in “What’s included”.
6. **Commit and push** to your fork, then open a **Pull Request** with a clear description of the change.

## Code style

- **Indentation:** Use 4 spaces.
- **Naming:** PascalCase for procedures; descriptive names for parameters (e.g. `targetSheet`, `columnReference`).
- **Comments:** Use `'` for VBA comments; keep them concise and in English.
- **Scope:** Prefer `Public` for procedures that are part of the library API; use `Private` for internal helpers.

## Testing

- Run your macro or UDF in Excel (with macros enabled) and confirm it behaves as expected.
- If you add or change AD-related code, test on a domain-joined machine if possible, or note limitations in the API doc.

## Documentation

- Every **Public** procedure should have an entry in **docs/API.md** with:
  - Full signature (name and parameters).
  - Short description and, where relevant, parameter meanings and allowed values.
  - One example (cell formula or VBA call) if it helps.

If you’re unsure about structure or naming, open an issue and we can align before you invest time in a PR.
