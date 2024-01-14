# CS_130

## Testing

```
python3 -m sheets.tests
```

## TODO

- [ ] write tests
  - [ ] basic dependency tests
  - [ ] cyclic dependency test
  - [ ] topological sort test
- [ ] formula parsing
  - [x] add
  - [x] mul
  - [x] unary
  - [x] concatenation
  - [ ] cell references
  - [ ] finding errors
    - [ ] parse error
    - [ ] bad reference (should move this to the dependency discovery code)
- [ ] dependency discovery
  - [ ] maintain dependency graph
    - [ ] remove edges around cells whose content changes
  - [ ] identify cycles
  - [ ] topological sort