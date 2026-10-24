# 🧱 Production source

Production VBA source: the only input to the CCR-K-SFTs workbook. **Nothing
is here yet.** The repository is in setup, and no calculation engine, facade
or workbook exists ([docs/REPOSITORY_STRUCTURE.md](../docs/REPOSITORY_STRUCTURE.md)).

| Location | Will hold | Rule |
| --- | --- | --- |
| `core/` | Calculation engine | Never touches Excel; every standard module declares `Option Private Module` |
| `modules/` | Public facade | Validates inputs and delegates to `core/`; supported declarations are listed in [docs/PUBLIC_API.txt](../docs/PUBLIC_API.txt) |
| `classes/` | Class modules | Status stated in each class header |
| `workbook/` | Document modules and the macro-free workbook template | Host glue only; no calculation logic |
| `forms/` | UserForms, `.frm` beside its `.frx` | Only if a form is needed |

Each subdirectory is created with its first real file. Test and example
modules never belong here. Naming, encoding and import/export conventions are
set in [#4](https://github.com/danielep71/CCR-K-SFTs/issues/4); the dependency
direction and placement rules in
[docs/REPOSITORY_STRUCTURE.md](../docs/REPOSITORY_STRUCTURE.md).
