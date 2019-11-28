# VBA List - Yet another VBA array wrapper

VBA List type to fit the IT requirement at work.

## Characteristics

1. No dependency. Not even `mscorlib.dll` and `Microsoft Scripting Runtime`.
2. No `interface` class for simplicity.
3. 1-based.
4. Automatically manage the array capacity to minimize the occurrence of internal `Redim`.
5. Self-implement.
6. Follow the convention of .Net ArrayList and SortedList classes, and some from Python List and Numpy Array. 
7. Abuse of VBA type check. That is, almost no custom type check at all.
8. Abuse of the `Variant` type. Efficiency is to sacrifice.

## Provided Types

- [x] List
- [x] Mapping
- [x] Dictionary, same as mapping, a drop-in replacement for `Dictionary`
- [ ] Matrix
- [ ] Table

*This is a working prototype, not code-reviewed yet. Use at your own risk.*
