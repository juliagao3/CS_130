# CS_130

## Testing

```
python3 -m sheets.tests
```

## TODO

- [ ] write tests
  - [x] basic dependency tests
  - [x] cyclic dependency test
  - [x] topological sort test
- [ ] formula parsing
  - [x] add
  - [x] mul
  - [x] unary
  - [x] concatenation
  - [x] cell references
  - [ ] finding errors
    - [ ] parse error
    - [ ] bad reference (should move this to the dependency discovery code)
- [ ] dependency discovery
  - [x] maintain dependency graph
    - [x] remove edges around cells whose content changes
  - [x] identify cycles
  - [ ] topological sort