import enum

from typing import TypeVar, Generic, Dict, List, Set, Tuple

T = TypeVar("T")

class EdgeType(enum.Enum):

    ''' An edge of this type exists from each cell to each cell it references
        regardless of whether the reference gets used in evaluation. '''
    REFERENCE = 0

    ''' An edge of this type exists from each cell to each cell it references
        only when the cell actually used the reference in evaluation. '''
    EVALUATED_REFERENCE = 1

    All = { REFERENCE, EVALUATED_REFERENCE }

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
        self.forward: Dict[T, Dict[T, Set[EdgeType]]] = {}
        self.backward: Dict[T, Dict[T, Set[EdgeType]]] = {}
        self.topological_order: List[T] = []
        self.cycles: List[Set[T]] = []
        self.cycles_dirty: bool = False

    def get_forward_links(self, node):
        if node not in self.forward:
            return {}
        else:
            return self.forward[node].keys()
        
    def get_backward_links(self, node):
        if node not in self.backward:
            return {}
        else:
            return self.backward[node].keys()

    def get_forward_links_of_type(self, node, types: Set[EdgeType]):
        if node not in self.forward:
            return {}
        else:
            return {t[0] for t in self.forward[node].items() if len(types.intersection(t[1])) > 0}

    def get_backward_links_of_type(self, node, types: Set[EdgeType]):
        if node not in self.backward:
            return {}
        else:
            return {t[0] for t in self.backward[node].items() if len(types.intersection(t[1])) > 0}

    def get_nodes(self):
        return self.forward.keys() | self.backward.keys()

    def is_cycle(self, scc):
        if len(scc) > 1:
            return True
        return (scc[0] in self.forward) and (scc[0] in self.get_forward_links(scc[0]))

    def get_cycles(self, edge_types: Set[EdgeType] = None):
        if edge_types is None:
            edge_types = {EdgeType.EVALUATED_REFERENCE}

        if self.cycles_dirty:
            sccs, topo = self.tarjan(edge_types)
            self.cycles = list(filter(lambda s: self.is_cycle(s), sccs))
            self.topological_order = topo
            self.cycles_dirty = False
        return self.cycles

    def get_topological_order(self, edge_types: Set[EdgeType] = None):
        if edge_types is None:
            edge_types = {EdgeType.EVALUATED_REFERENCE}

        if self.cycles_dirty:
            sccs, topo = self.tarjan(edge_types)
            self.cycles = list(filter(lambda s: self.is_cycle(s), sccs))
            self.topological_order = topo
            self.cycles_dirty = False
        return self.topological_order

    def link(self, from_node: T, to_node: T, edge_types: Set[EdgeType]):
        '''
        Add entries to the forward and backward maps to create a link between
        the from_node and to_node.
        '''
        if from_node not in self.forward:
            self.forward[from_node] = dict()

        if to_node not in self.backward:
            self.backward[to_node] = dict()

        if to_node not in self.forward[from_node]:
            self.forward[from_node][to_node] = set()

        if from_node not in self.backward[to_node]:
            self.backward[to_node][from_node] = set()

        self.forward[from_node][to_node] |= edge_types
        self.backward[to_node][from_node] |= edge_types

        self.cycles_dirty = True

    def clear_forward_links(self, node: T, edge_types: Set[EdgeType] = None):
        '''
        Remove all forward links coming from the given node.
        '''
        if edge_types is None:
            edge_types = {EdgeType.REFERENCE, EdgeType.EVALUATED_REFERENCE}

        links = self.get_forward_links_of_type(node, edge_types)

        if len(links) == 0:
            return

        for to in links:
            # LINE A:
            self.forward[node][to] -= edge_types

            if len(self.forward[node][to]) == 0:
                self.forward[node].pop(to)

            # LINE B:
            self.backward[to][node] -= edge_types

            if len(self.backward[to][node]) == 0:
                self.backward[to].pop(node)

        # This will invalidate in the case where LINE A and LINE B above
        # don't actually change the graph. Not ideal but I'll fix later.
        self.cycles_dirty = True

    def clear_backward_link(self, node: T, edge_types: Set[EdgeType] = None):
        '''
        Remove all backward links coming from the given node.
        '''
        if edge_types is None:
            edge_types = {EdgeType.REFERENCE, EdgeType.EVALUATED_REFERENCE}

        links = self.get_backward_links_of_type(node, edge_types)

        if len(links) == 0:
            return

        for from_node in links:
            # LINE A:
            self.forward[from_node][node] -= edge_types

            if len(self.forward[from_node][node]) == 0:
                self.forward[from_node].pop(node)

            # LINE B:
            self.backward[node][from_node] -= edge_types

            if len(self.backward[node][from_node]) == 0:
                self.backward[node].pop(from_node)

        # This will invalidate in the case where LINE A and LINE B above
        # don't actually change the graph. Not ideal but I'll fix later.
        self.cycles_dirty = True

    def remove_node(self, node: T):
        self.clear_forward_links(node, EdgeType.All)
        self.clear_backward_link(node, EdgeType.All)

    def strongconnect(
                self,
                v: T,
                next_index: int,
                index: Dict[T, int],
                lowlink: Dict[T, int],
                on_stack: Dict[T, bool],
                stack: List[T],
                sccs: List[List[T]],
                topo: List[T],
                edge_types: Set[EdgeType]
            ):

        call_stack = [(v, list(self.get_forward_links_of_type(v, edge_types)), 0)]

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
                    call_stack.append((w, list(self.get_forward_links_of_type(w, edge_types)), 0))
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

    def tarjan(self, edge_types: Set[EdgeType]):
        next_index = 0
        index = {}
        lowlink = {}
        on_stack = {}
        stack = []
        sccs = []
        topo = []
        for n in self.get_nodes():
            if n not in index:
                next_index = self.strongconnect(n, next_index, index, lowlink, on_stack, stack, sccs, topo, edge_types)
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
