from typing import *

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

    def link(self, from_node: T, to_node: T):
        '''
        Add entries to the forward and backward maps to create a link between
        the from_node and to_node.
        '''
        pass

    def clear_forward_links(self, node: T):
        '''
        Remove all forward links coming from the given node.
        '''
        pass

    def find_cycle(self, root: T):
        '''
        Returns a cycle containing this node or None if no such cycle exists.
        '''
        return None

    def get_ancestors(self, root: T):
        '''
        Returns the nodes that need to be recomputed following a change in
        the given node. The nodes should be ordered so that no node must be
        recomputed twice.

        In graph terms this is the set of nodes reachable by following backward
        links from the root in topological sort order.
        '''
        return []