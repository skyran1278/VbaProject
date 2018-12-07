class Solution:
    def numIslands(self, grid):
        """
        :type grid: List[List[str]]
        :rtype: int
        """
        queue = []
        used = []
        num_island = 0
        root = (0, 0)
        if grid[root[0]][root[1]] == '1':
            num_island += 1
        # root = grid[0][0]

        queue.append(root)
        used.append(root)

        while (queue):
            size = len(queue)
            for i in range(size):
                cur_x, cur_y = queue[i]

                for coordinate in ((cur_x + 1, cur_y), (cur_x, cur_y + 1)):
                    if grid[cur_x][cur_y] != grid[coordinate[0]][coordinate[1]]:
                        pass
                    if coordinate not in used:
                        queue.append(coordinate)
                        used.append(coordinate)

                queue.pop(0)


sol = Solution()
sol.numIslands([["1", "1", "1", "1", "0"], ["1", "1", "0", "1", "0"], [
               "1", "1", "0", "0", "0"], ["0", "0", "0", "0", "0"]])

# /**
#  * Return the length of the shortest path between root and target node.
#  */
# int BFS(Node root, Node target) {
#     Queue<Node> queue;  // store all nodes which are waiting to be processed
#     Set<Node> used;     // store all the used nodes
#     int step = 0;       // number of steps neeeded from root to current node
#     // initialize
#     add root to queue;
#     add root to used;
#     // BFS
#     while (queue is not empty) {
#         step = step + 1;
#         // iterate the nodes which are already in the queue
#         int size = queue.size();
#         for (int i = 0; i < size; ++i) {
#             Node cur = the first node in queue;
#             return step if cur is target;
#             for (Node next : the neighbors of cur) {
#                 if (next is not in used) {
#                     add next to queue;
#                     add next to used;
#                 }
#             }
#             remove the first node from queue;
#         }
#     }
#     return -1;          // there is no path from root to target
# }

# class MyCircularQueue:

#     def __init__(self, k):
#         """
#         Initialize your data structure here. Set the size of the queue to be k.
#         :type k: int
#         """
#         self.queue = []
#         self.size = k
#         self._first = None


#     def enQueue(self, value):
#         """
#         Insert an element into the circular queue. Return true if the operation is successful.
#         :type value: int
#         :rtype: bool
#         """
#         if self.isFull():
#             return False
#         self.queue.append(value)
#         return True


#     def deQueue(self):
#         """
#         Delete an element from the circular queue. Return true if the operation is successful.
#         :rtype: bool
#         """
#         if self.isEmpty():
#             return False
#         self.queue.pop(0)
#         return True


#     def Front(self):
#         """
#         Get the front item from the queue.
#         :rtype: int
#         """
#         if self.isEmpty():
#             return -1
#         return self.queue[0]


#     def Rear(self):
#         """
#         Get the last item from the queue.
#         :rtype: int
#         """
#         if self.isEmpty():
#             return -1

#         return self.queue[len(self.queue) - 1]


#     def isEmpty(self):
#         """
#         Checks whether the circular queue is empty or not.
#         :rtype: bool
#         """
#         return len(self.queue) == 0


#     def isFull(self):
#         """
#         Checks whether the circular queue is full or not.
#         :rtype: bool
#         """
#         return len(self.queue) == self.size


# # Your MyCircularQueue object will be instantiated and called as such:
# # obj = MyCircularQueue(k)
# # param_1 = obj.enQueue(value)
# # param_2 = obj.deQueue()
# # param_3 = obj.Front()
# # param_4 = obj.Rear()
# # param_5 = obj.isEmpty()
# # param_6 = obj.isFull()
