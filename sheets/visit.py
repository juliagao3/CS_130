import lark

#scan through a formula, collecting cell references
#reference workbook, sheet, formula
class CellRefFinder(lark.Visitor):
    def __init__(self):
        self.refs = []
        
    def cell(self, tree):
        #print(tree.children)
        if len(tree.children) == 1:
            self.refs.append(str(tree.children[0]))
        else:
            assert len(tree.children) == 2
            self.refs.append('!'.join(tree.children))


parser = lark.Lark.open('formulas.lark', start='formula')
tree = parser.parse('= a1 + 3 * Sheet2!A2')
print(tree.pretty())

v = CellRefFinder()
print(v.refs)
v.visit(tree)
print(v.refs)