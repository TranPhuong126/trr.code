# main.py
from graph_class import Graph

def main():
    print("=== CHƯƠNG TRÌNH MÔ PHỎNG TÔ MÀU ĐỒ THỊ ===")
    
    # Tạo đồ thị mẫu (ví dụ: 4 đỉnh)
    g = Graph(4)
    g.add_edge(0, 1)
    g.add_edge(0, 2)
    g.add_edge(1, 3)
    g.display()

    # (Tuần 3 sẽ bổ sung thuật toán tô màu ở đây)

if __name__ == "__main__":
    main()
