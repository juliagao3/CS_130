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

    def get_forward_links(self, node):
        if not node in self.forward:
            return []
        else:
            return self.forward[node]
        
    def get_backward_links(self, node):
        if not node in self.backward:
            return []
        else:
            return self.backward[node]

    def link(self, from_node: T, to_node: T):
        '''
        Add entries to the forward and backward maps to create a link between
        the from_node and to_node.
        '''
        if not from_node in self.forward:
            self.forward[from_node] = set()

        if not to_node in self.backward:
            self.backward[to_node] = set()

        self.forward[from_node].add(to_node)
        self.backward[to_node].add(from_node)

    def clear_forward_links(self, node: T):
        '''
        Remove all forward links coming from the given node.
        '''
        if node in self.forward:
            for to in self.forward[node]:
                self.backward[to].remove(node)
                if len(self.backward[to]) == 0:
                    self.backward.pop(to)
            self.forward.pop(node)

    def find_cycle(self, root: T):
        '''
        Returns a cycle containing this node or None if no such cycle exists.
        '''
        # this is a non-recursive dfs... I think this could've probly been
        # done better
        seen = set()
        path = []
        stack = [(0, root, list(self.get_forward_links(root)))]
        while len(stack) > 0:
            i, cur, children = stack.pop()

            if i == 0:
                if cur in seen and cur in path:
                    path = path[path.index(cur):]
                    break
                seen.add(cur)
                path.append(cur)

            if i < len(children):
                stack.append((i+1, cur, children))
                stack.append((0, children[i], list(self.get_forward_links(children[i]))))
            else:
                path.pop()

        if len(path) == 0:
            return None
        else:
            return path

    def topological_sort(self):
        ordering = []
        vertices = self.forward.keys() | self.backward.keys()
        while len(vertices) > 0:
            no_incoming = vertices - self.backward.keys()
            if len(no_incoming) == 0:
                break
            node = no_incoming.pop()
            vertices.remove(node)
            self.clear_forward_links(node)
            ordering.append(node)
        return ordering

    def get_ancestors(self, root: T):
        '''
        Return the nodes reachable from root by following backward links.
        '''
        seen = set()
        path = []
        stack = [(0, root, list(self.get_backward_links(root)))]
        subgraph = Graph()
        while len(stack) > 0:
            i, cur, children = stack.pop()

            if i == 0:
                if len(path) > 0:
                    subgraph.link(path[-1], cur)
                if cur in seen:
                    continue
                seen.add(cur)
                path.append(cur)

            if i < len(children):
                stack.append((i+1, cur, children))
                stack.append((0, children[i], list(self.get_backward_links(children[i]))))
            else:
                path.pop()
        
        return subgraph