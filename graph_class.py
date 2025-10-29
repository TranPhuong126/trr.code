# graph_class.py
class Graph:
    def __init__(self, num_vertices):
        self.num_vertices = num_vertices
        self.adj_matrix = [[0 for _ in range(num_vertices)] for _ in range(num_vertices)]

    def add_edge(self, u, v):
        """Thêm cạnh giữa đỉnh u và v"""
        if u < self.num_vertices and v < self.num_vertices:
            self.adj_matrix[u][v] = 1
            self.adj_matrix[v][u] = 1  # Vì đồ thị vô hướng
        else:
            print("Lỗi: Chỉ số đỉnh không hợp lệ!")

    def display(self):
        """Hiển thị ma trận kề"""
        print("Ma trận kề:")
        for row in self.adj_matrix:
            print(row)
