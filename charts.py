import matplotlib
# Configurar backend antes de importar pyplot
try:
    matplotlib.use('QtAgg')  # Backend para PySide6/PyQt6
except:
    try:
        matplotlib.use('Qt5Agg')  # Fallback para PySide2/PyQt5
    except:
        pass  # Usar backend padrão se não conseguir configurar

import matplotlib.pyplot as plt
try:
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
except ImportError:
    try:
        from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
    except ImportError:
        from matplotlib.backends.backend_agg import FigureCanvasAgg as FigureCanvas

from matplotlib.figure import Figure
import sqlite3
from datetime import datetime, timedelta
import numpy as np

class GraficosWidget:
    """Widget para exibir gráficos de vendas"""
    
    def __init__(self):
        self.fig = Figure(figsize=(12, 8), facecolor='#2b2b2b')
        self.canvas = FigureCanvas(self.fig)
        self.fig.patch.set_facecolor('#2b2b2b')
        
    def obter_dados(self, filtro_data_inicio=None, filtro_data_fim=None, filtro_cliente=None, filtro_produto=None):
        """Obtém dados do banco aplicando filtros"""
        conn = sqlite3.connect("vendas.db")
        cursor = conn.cursor()
        
        query = """
            SELECT v.data_venda, v.valor_total, v.cliente_nome, i.produto, i.quantidade
            FROM vendas v
            JOIN itens_venda i ON i.venda_id = v.id
            WHERE 1=1
        """
        
        params = []
        
        if filtro_data_inicio:
            query += " AND (v.data_venda IS NULL OR v.data_venda >= ?)"
            params.append(filtro_data_inicio.toString("yyyy-MM-dd"))
        
        if filtro_data_fim:
            query += " AND (v.data_venda IS NULL OR v.data_venda <= ?)"
            params.append(filtro_data_fim.toString("yyyy-MM-dd"))
        
        if filtro_cliente:
            query += " AND v.cliente_nome LIKE ?"
            params.append(f"%{filtro_cliente}%")
        
        if filtro_produto:
            query += " AND i.produto LIKE ?"
            params.append(f"%{filtro_produto}%")
        
        cursor.execute(query, params)
        dados = cursor.fetchall()
        conn.close()
        
        return dados
    
    def grafico_vendas_por_data(self, filtro_data_inicio=None, filtro_data_fim=None, filtro_cliente=None, filtro_produto=None):
        """Gráfico de vendas ao longo do tempo"""
        dados = self.obter_dados(filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
        
        if not dados:
            self.fig.clear()
            ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
            ax.text(0.5, 0.5, 'Não há dados para exibir', 
                   ha='center', va='center', color='white', fontsize=14)
            ax.set_facecolor('#2b2b2b')
            self.canvas.draw()
            return
        
        # Agrupar por data
        vendas_por_data = {}
        for linha in dados:
            data_str = linha[0] if linha[0] else "Sem data"
            valor = float(linha[1]) if linha[1] else 0
            if data_str not in vendas_por_data:
                vendas_por_data[data_str] = 0
            vendas_por_data[data_str] += valor
        
        # Ordenar por data
        datas_ordenadas = sorted([d for d in vendas_por_data.keys() if d != "Sem data"], 
                                key=lambda x: datetime.strptime(x, "%Y-%m-%d") if isinstance(x, str) and len(x) >= 10 else datetime.min)
        valores_ordenados = [vendas_por_data[d] for d in datas_ordenadas]
        
        # Adicionar "Sem data" no final se existir
        if "Sem data" in vendas_por_data:
            datas_ordenadas.append("Sem data")
            valores_ordenados.append(vendas_por_data["Sem data"])
        
        self.fig.clear()
        ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
        
        # Limitar número de datas no eixo X para legibilidade
        if len(datas_ordenadas) > 20:
            # Mostrar apenas algumas datas
            indices = np.linspace(0, len(datas_ordenadas)-1, 10, dtype=int)
            datas_exibidas = [datas_ordenadas[i] if i < len(datas_ordenadas) else "" for i in indices]
            ax.plot(range(len(valores_ordenados)), valores_ordenados, 
                   color='#4CAF50', linewidth=2, marker='o', markersize=4)
            ax.set_xticks(indices)
            ax.set_xticklabels(datas_exibidas, rotation=45, ha='right', color='white')
        else:
            ax.plot(range(len(valores_ordenados)), valores_ordenados, 
                   color='#4CAF50', linewidth=2, marker='o', markersize=4)
            ax.set_xticks(range(len(datas_ordenadas)))
            ax.set_xticklabels(datas_ordenadas, rotation=45, ha='right', color='white')
        
        ax.set_ylabel('Valor Total (R$)', color='white', fontsize=12)
        ax.set_xlabel('Data', color='white', fontsize=12)
        ax.set_title('Vendas por Data', color='white', fontsize=14, fontweight='bold')
        ax.grid(True, alpha=0.3, color='gray')
        ax.set_facecolor('#2b2b2b')
        ax.tick_params(colors='white')
        
        self.fig.tight_layout()
        self.canvas.draw()
    
    def grafico_top_produtos(self, top_n=10, filtro_data_inicio=None, filtro_data_fim=None, filtro_cliente=None, filtro_produto=None):
        """Gráfico dos top produtos mais vendidos"""
        dados = self.obter_dados(filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
        
        if not dados:
            self.fig.clear()
            ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
            ax.text(0.5, 0.5, 'Não há dados para exibir', 
                   ha='center', va='center', color='white', fontsize=14)
            ax.set_facecolor('#2b2b2b')
            self.canvas.draw()
            return
        
        # Agrupar por produto
        produtos_vendas = {}
        for linha in dados:
            produto = linha[3] if linha[3] else "Sem nome"
            valor = float(linha[1]) if linha[1] else 0
            if produto not in produtos_vendas:
                produtos_vendas[produto] = 0
            produtos_vendas[produto] += valor
        
        # Ordenar e pegar top N
        produtos_ordenados = sorted(produtos_vendas.items(), key=lambda x: x[1], reverse=True)[:top_n]
        
        produtos = [p[0][:30] + "..." if len(p[0]) > 30 else p[0] for p in produtos_ordenados]
        valores = [p[1] for p in produtos_ordenados]
        
        self.fig.clear()
        ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
        
        cores = plt.cm.viridis(np.linspace(0, 1, len(produtos)))
        bars = ax.barh(range(len(produtos)), valores, color=cores)
        
        ax.set_yticks(range(len(produtos)))
        ax.set_yticklabels(produtos, color='white')
        ax.set_xlabel('Valor Total (R$)', color='white', fontsize=12)
        ax.set_title(f'Top {top_n} Produtos Mais Vendidos', color='white', fontsize=14, fontweight='bold')
        ax.set_facecolor('#2b2b2b')
        ax.tick_params(colors='white')
        ax.grid(True, alpha=0.3, color='gray', axis='x')
        
        # Adicionar valores nas barras
        for i, (bar, valor) in enumerate(zip(bars, valores)):
            ax.text(valor, i, f' R$ {valor:.2f}', 
                   va='center', color='white', fontsize=9)
        
        self.fig.tight_layout()
        self.canvas.draw()
    
    def grafico_vendas_por_cliente(self, top_n=10, filtro_data_inicio=None, filtro_data_fim=None, filtro_cliente=None, filtro_produto=None):
        """Gráfico de vendas por cliente"""
        dados = self.obter_dados(filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
        
        if not dados:
            self.fig.clear()
            ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
            ax.text(0.5, 0.5, 'Não há dados para exibir', 
                   ha='center', va='center', color='white', fontsize=14)
            ax.set_facecolor('#2b2b2b')
            self.canvas.draw()
            return
        
        # Agrupar por cliente
        clientes_vendas = {}
        for linha in dados:
            cliente = linha[2] if linha[2] else "Sem nome"
            valor = float(linha[1]) if linha[1] else 0
            if cliente not in clientes_vendas:
                clientes_vendas[cliente] = 0
            clientes_vendas[cliente] += valor
        
        # Ordenar e pegar top N
        clientes_ordenados = sorted(clientes_vendas.items(), key=lambda x: x[1], reverse=True)[:top_n]
        
        clientes = [c[0][:25] + "..." if len(c[0]) > 25 else c[0] for c in clientes_ordenados]
        valores = [c[1] for c in clientes_ordenados]
        
        self.fig.clear()
        ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
        
        cores = plt.cm.plasma(np.linspace(0, 1, len(clientes)))
        bars = ax.bar(range(len(clientes)), valores, color=cores)
        
        ax.set_xticks(range(len(clientes)))
        ax.set_xticklabels(clientes, rotation=45, ha='right', color='white')
        ax.set_ylabel('Valor Total (R$)', color='white', fontsize=12)
        ax.set_title(f'Top {top_n} Clientes', color='white', fontsize=14, fontweight='bold')
        ax.set_facecolor('#2b2b2b')
        ax.tick_params(colors='white')
        ax.grid(True, alpha=0.3, color='gray', axis='y')
        
        # Adicionar valores nas barras
        for bar, valor in zip(bars, valores):
            altura = bar.get_height()
            ax.text(bar.get_x() + bar.get_width()/2., altura,
                   f'R$ {valor:.2f}',
                   ha='center', va='bottom', color='white', fontsize=9)
        
        self.fig.tight_layout()
        self.canvas.draw()
    
    def grafico_pizza_categorias(self, filtro_data_inicio=None, filtro_data_fim=None, filtro_cliente=None, filtro_produto=None):
        """Gráfico de pizza com distribuição de vendas"""
        dados = self.obter_dados(filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
        
        if not dados:
            self.fig.clear()
            ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
            ax.text(0.5, 0.5, 'Não há dados para exibir', 
                   ha='center', va='center', color='white', fontsize=14)
            ax.set_facecolor('#2b2b2b')
            self.canvas.draw()
            return
        
        # Agrupar por produto (top 5)
        produtos_vendas = {}
        for linha in dados:
            produto = linha[3] if linha[3] else "Sem nome"
            valor = float(linha[1]) if linha[1] else 0
            if produto not in produtos_vendas:
                produtos_vendas[produto] = 0
            produtos_vendas[produto] += valor
        
        # Ordenar e pegar top 5, resto como "Outros"
        produtos_ordenados = sorted(produtos_vendas.items(), key=lambda x: x[1], reverse=True)
        top_5 = produtos_ordenados[:5]
        outros_valor = sum(p[1] for p in produtos_ordenados[5:])
        
        labels = [p[0][:20] + "..." if len(p[0]) > 20 else p[0] for p in top_5]
        valores = [p[1] for p in top_5]
        
        if outros_valor > 0:
            labels.append("Outros")
            valores.append(outros_valor)
        
        self.fig.clear()
        ax = self.fig.add_subplot(111, facecolor='#2b2b2b')
        
        cores = plt.cm.Set3(np.linspace(0, 1, len(labels)))
        wedges, texts, autotexts = ax.pie(valores, labels=labels, autopct='%1.1f%%', 
                                         colors=cores, startangle=90, textprops={'color': 'white'})
        
        ax.set_title('Distribuição de Vendas por Produto', color='white', fontsize=14, fontweight='bold')
        
        self.fig.tight_layout()
        self.canvas.draw()

