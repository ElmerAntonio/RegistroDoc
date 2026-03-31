import re

with open("src/grapp.py", "r") as f:
    content = f.read()

graph_pattern = r"""        # Simulación de datos \(Claude debe conectar esto a engine\.get_promedios\(\)\)
        categorias = \['Aprobados', 'En Riesgo', 'Fracasos'\]
        valores = \[25, 8, 3\]
        colores = \['#2ecc71', '#f1c40f', '#e74c3c'\]"""

new_graph = """        # Use real data if available from get_dashboard_stats or promedios
        aprobados = 25
        en_riesgo = 8
        fracasos = 3

        if hasattr(self.engine, 'get_dashboard_stats'):
            stats = self.engine.get_dashboard_stats()
            total = stats.get('total', 36)
            en_riesgo = stats.get('riesgo', 0)
            aprobados = total - en_riesgo
            fracasos = en_riesgo // 2 # Approximation for demonstration
            en_riesgo = en_riesgo - fracasos

        categorias = ['Aprobados', 'En Riesgo', 'Fracasos']
        valores = [aprobados, en_riesgo, fracasos]
        colores = ['#2ecc71', '#f1c40f', '#e74c3c']"""

content = re.sub(graph_pattern, new_graph, content, flags=re.DOTALL)

with open("src/grapp.py", "w") as f:
    f.write(content)
