from typing import TypeVar, Generic, Dict, List, Set

T = TypeVar("T")

class Graph(Generic[T]):
    '''
    Represents a directed graph.

        Attributes:
            forward  - maps node A to a set of nodes which contains B if and
                       only if directed edge A -> B exists in the graph
            backward - maps node B to a set of nodes which contains A if and
                       only if directed edge A -> B exists in the graph
    '''

    def __init__(self):
        self.forward: Dict[T, T] = {}
        self.backward: Dict[T, T] = {}
        self.topological_order: List[T] = []
        self.cycles: List[Set[T]] = []
        self.cycles_dirty: bool = False

    def get_forward_links(self, node):
        if node not in self.forward:
            return []
        else:
            return self.forward[node].copy()
        
    def get_backward_links(self, node):
        if node not in self.backward:
            return []
        else:
            return self.backward[node].copy()
        
    def get_nodes(self):
        return self.forward.keys() | self.backward.keys()

    def is_cycle(self, scc):
        if len(scc) > 1:
            return True
        return (scc[0] in self.forward) and (scc[0] in self.forward[scc[0]])

    def get_cycles(self):
        if self.cycles_dirty:
            sccs, topo = self.tarjan()
            self.cycles = list(filter(lambda s: self.is_cycle(s), sccs))
            self.topological_order = topo
            self.cycles_dirty = False
        return self.cycles

    def get_topological_order(self):
        if self.cycles_dirty:
            sccs, topo = self.tarjan()
            self.cycles = list(filter(lambda s: self.is_cycle(s), sccs))
            self.topological_order = topo
            self.cycles_dirty = False
        return self.topological_order

    def link(self, from_node: T, to_node: T):
        '''
        Add entries to the forward and backward maps to create a link between
        the from_node and to_node.
        '''
        if from_node not in self.forward:
            self.forward[from_node] = set()
        elif to_node in self.forward[from_node]:
            assert from_node in self.backward[to_node]
            return

        if to_node not in self.backward:
            self.backward[to_node] = set()

        self.forward[from_node].add(to_node)
        self.backward[to_node].add(from_node)

        self.cycles_dirty = True

    def clear_forward_links(self, node: T):
        '''
        Remove all forward links coming from the given node.
        '''
        if node not in self.forward:
            return

        links = self.forward[node]
        self.forward.pop(node)

        if len(links) == 0:
            return

        for to in links:
            self.backward[to].remove(node)
            if len(self.backward[to]) == 0:
                self.backward.pop(to)
        self.cycles_dirty = True

    def clear_backward_link(self, node: T):
        if node not in self.backward:
            return

        links = self.backward[node]
        self.backward.pop(node)

        if len(links) == 0:
            return

        for from_node in links:
            self.forward[from_node].remove(node)
            if len(self.forward[from_node]) == 0:
                self.forward.pop(from_node)
        self.cycles_dirty = True

    def remove_node(self, node: T):
        self.clear_forward_links(node)
        self.clear_backward_link(node)

    def tarjan(self):
        stack: List[T] = []
        stack_set: Set[T] = set()

        next_number: int = 0
        number: Dict[T, int] ={}
        lowlink: Dict[T, int] = {}

        sccs: List[List[T]] = []
        topological_order = []

        def strongconnect(root):
            nonlocal self
            nonlocal stack
            nonlocal stack_set

            nonlocal next_number
            nonlocal number
            nonlocal lowlink

            nonlocal sccs

            dfs = [(root, None, self.get_forward_links(root))]
            while len(dfs) > 0:
                v, w, children = dfs.pop()

                if w is None:
                    number[v] = next_number
                    lowlink[v] = next_number
                    next_number += 1
                    stack.append(v)
                    stack_set.add(v)
                else:
                    lowlink[v] = min(lowlink[v], lowlink[w])

                recurse = False

                while len(children) > 0:
                    w = children.pop()
                    if w not in number:
                        dfs.append((v, w, children))
                        dfs.append((w, None, self.get_forward_links(w)))
                        recurse = True
                        break
                    elif w in stack_set:
                        lowlink[v] = min(lowlink[v], number[w])

                if recurse:
                    continue

                topological_order.append(v)
                if lowlink[v] == number[v]:
                    scc = []
                    while True:
                        w = stack.pop()
                        stack_set.remove(w)
                        scc.append(w)
                        if w == v:
                            break
                    sccs.append(scc)

        vertices = self.forward.keys() | self.backward.keys()
        for v in vertices:
            if v not in number:
                strongconnect(v)

        return sccs, topological_order

    def get_ancestors_of_set(self, nodes):
        '''
        Returns the set of nodes reachable by following backward links starting
        from any node in the given set. Does not return any nodes in the set.
        '''
        ancestors = set()
        queue = [n for n in nodes]
        while len(queue) > 0:
            v = queue.pop(0)

            if v not in nodes:
                ancestors.add(v)

            for w in self.get_backward_links(v):
                if w not in ancestors and w not in nodes:
                    queue.append(w)
        return ancestors
