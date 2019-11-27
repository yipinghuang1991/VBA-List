# VBA List - Yet another VBA array wrapper

VBA List type to fit the IT requirement at work.

## Characteristics

1. 1-based.
2. Automatically manage the array capacity to minimize the occurrence of internal `Redim`.
3. No dependency. Not even mscorlib and Microsoft scripting language.
4. Self-implement.
5. Follow the convention of .Net ArrayList and SortedList classes, and some from Python List and Numpy Array. 
6. Abuse of VBA type check. That is, no custom type check at all.
7. Abuse of the `Variant` type. Efficiency is to sacrifice.

## Provided Types

- [x] List
- [x] Mapping
- [ ] Matrix
- [ ] Table
