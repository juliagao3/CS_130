from typing import TypeVar, Generic, Dict, List, Set, Tuple

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
        self.forward: Dict[T, Tuple[Set[T], Set[T]]] = {}
        self.backward: Dict[T, Tuple[Set[T], Set[T]]] = {}
        self.topological_order: List[T] = []
        self.cycles: List[Set[T]] = []
        self.cycles_dirty: bool = False

    def get_forward_links(self, node):
        if node not in self.forward:
            return {}
        else:
            t = self.forward[node]
            return t[0] | t[1]
        
    def get_backward_links(self, node):
        if node not in self.backward:
            return {}
        else:
            t = self.backward[node]
            return t[0] | t[1]

    def get_nodes(self):
        return self.forward.keys() | self.backward.keys()

    def is_cycle(self, scc):
        if len(scc) > 1:
            return True
        return (scc[0] in self.forward) and (scc[0] in self.get_forward_links(scc[0]))

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
            self.forward[from_node] = (set(), set())
        elif to_node in self.forward[from_node][0]:
            assert from_node in self.backward[to_node][0]
            return

        if to_node not in self.backward:
            self.backward[to_node] = (set(), set())

        self.forward[from_node][0].add(to_node)
        self.backward[to_node][0].add(from_node)

        self.cycles_dirty = True

    def clear_forward_links(self, node: T):
        '''
        Remove all forward links coming from the given node.
        '''
        if node not in self.forward:
            return

        links = self.get_forward_links(node)
        self.forward.pop(node)

        if len(links) == 0:
            return

        for to in links:
            static, runtime = self.backward[to]

            if node in static:
                static.remove(node)

            if node in runtime:
                runtime.remove(node)

            if len(self.get_backward_links(to)) == 0:
                self.backward.pop(to)
        self.cycles_dirty = True

    def clear_backward_link(self, node: T):
        if node not in self.backward:
            return

        links = self.get_backward_links(node)
        self.backward.pop(node)

        if len(links) == 0:
            return

        for from_node in links:
            static, runtime = self.forward[from_node]

            if node in static:
                static.remove(node)

            if node in runtime:
                runtime.remove(node)

            if len(self.get_forward_links(node)) == 0:
                self.forward.pop(from_node)
        self.cycles_dirty = True

    def link_runtime(self, from_node: T, to_node: T):
        '''
        Add entries to the forward and backward maps to create a link between
        the from_node and to_node.
        '''
        if from_node not in self.forward:
            self.forward[from_node] = (set(), set())
        elif to_node in self.forward[from_node][1]:
            assert from_node in self.backward[to_node][1]
            return

        if to_node not in self.backward:
            self.backward[to_node] = (set(), set())

        self.forward[from_node][1].add(to_node)
        self.backward[to_node][1].add(from_node)

        self.cycles_dirty = True
        
    def clear_forward_runtime_links(self, node: T):
        '''
        Remove all forward links coming from the given node.
        '''
        if node not in self.forward:
            return

        links = self.forward[node][1]

        if len(links) == 0:
            return

        for to in links:
            self.backward[to][1].remove(node)
            if len(self.get_backward_links(to)) == 0:
                self.backward.pop(to)

        links.clear()

        self.cycles_dirty = True

    def clear_backward_runtime_link(self, node: T):
        if node not in self.backward:
            return

        links = self.backward[node][1]

        if len(links) == 0:
            return

        for from_node in links:
            self.forward[from_node][1].remove(node)
            if len(self.get_forward_links(node)) == 0:
                self.forward.pop(from_node)
        
        links.clear()

        self.cycles_dirty = True
        
    def remove_node(self, node: T):
        self.clear_forward_links(node)
        self.clear_backward_link(node)

    def strongconnect(
                self,
                v: T,
                next_index: int,
                index: Dict[T, int],
                lowlink: Dict[T, int],
                on_stack: Dict[T, bool],
                stack: List[T],
                sccs: List[List[T]],
                topo: List[T]
            ):

        call_stack = [(v, list(self.get_forward_links(v)) if v in self.forward else [], 0)]

        while len(call_stack) > 0:
            v, forward, i = call_stack.pop()

            if i == 0:
                index[v] = next_index
                lowlink[v] = next_index
                next_index += 1
                stack.append(v)
                on_stack[v] = True
            else:
                w = forward[i-1]
                lowlink[v] = min(lowlink[v], lowlink[w])

            recurse = False

            while i < len(forward):
                w = forward[i]
                i += 1
                if w not in index:
                    call_stack.append((v, forward, i))
                    call_stack.append((w, list(self.get_forward_links(w)) if w in self.forward else [], 0))
                    recurse = True
                    break
                elif w in on_stack and on_stack[w]:
                    lowlink[v] = min(lowlink[v], index[w])

            if recurse:
                continue
            
            topo.append(v)
            if lowlink[v] == index[v]:
                scc = []
                while True:
                    w = stack.pop()
                    on_stack[w] = False
                    scc.append(w)
                    if w == v:
                        break
                sccs.append(scc)

        return next_index

    def tarjan(self):
        next_index = 0
        index = {}
        lowlink = {}
        on_stack = {}
        stack = []
        sccs = []
        topo = []
        for n in self.get_nodes():
            if n not in index:
                next_index = self.strongconnect(n, next_index, index, lowlink, on_stack, stack, sccs, topo)
        return sccs, topo

    def get_ancestors_of_set(self, nodes):
        '''
        Returns the set of nodes reachable by following backward links starting
        from any node in the given set. Does not return any nodes in the set.
        '''
        ancestors = set()
        queue = [n for n in nodes]
        while len(queue) > 0:
            v = queue.pop()

            if v not in nodes:
                ancestors.add(v)

            for w in self.get_backward_links(v):
                if w not in ancestors and w not in nodes:
                    queue.append(w)
        return ancestors
