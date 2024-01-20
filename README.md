# CS_130

## Testing

```
python -m unittest tests.tests
```

## TODO

- [ ] write tests
  - [x] basic dependency tests
  - [x] cyclic dependency test
  - [x] topological sort test
- [x] formula parsing
  - [x] add
  - [x] mul
  - [x] unary
  - [x] concatenation
  - [x] cell references
  - [x] finding errors
    - [x] parse error
    - [x] bad reference (should move this to the dependency discovery code)
    - [x] propagate referenced cell errors
- [x] dependency discovery
  - [x] maintain dependency graph
    - [x] remove edges around cells whose content changes
  - [x] identify cycles
  - [x] topological sort