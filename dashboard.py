from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QPushButton, QTableWidget,
    QTableWidgetItem, QHBoxLayout, QFileDialog,
    QLineEdit, QLabel, QDateEdit, QMessageBox,
    QCheckBox, QHeaderView, QFrame, QTabWidget,
    QGridLayout, QComboBox, QDialog, QDialogButtonBox
)
from PySide6.QtCore import QDate, Qt, QTimer
from PySide6.QtGui import QColor, QFont, QPalette, QIcon
import os
import sys
import sqlite3
from import_excel import importar_excel
from pdf_report import gerar_pdf, gerar_pdf_personalizado
from export_excel import exportar_excel, exportar_excel_completo
from charts import GraficosWidget

class Dashboard(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sistema de Controle de Vendas - Mercado Livre")
        # Tamanho mínimo e máximo para responsividade
        self.setMinimumSize(1000, 600)
        self.resize(1400, 800)
        
        # Configurar ícone da janela
        # Tentar diferentes caminhos para o ícone (desenvolvimento e executável)
        icon_paths = [
            "icon.png",  # Pasta atual
            os.path.join(os.path.dirname(__file__), "icon.png"),  # Pasta do script
            os.path.join(sys._MEIPASS, "icon.png") if hasattr(sys, '_MEIPASS') else None,  # PyInstaller temp
        ]
        
        for icon_path in icon_paths:
            if icon_path and os.path.exists(icon_path):
                self.setWindowIcon(QIcon(icon_path))
                break
        
        # Aplicar estilo moderno com fundo escuro
        self.setStyleSheet("""
            QWidget {
                background-color: #2b2b2b;
                font-family: 'Segoe UI', Arial, sans-serif;
                color: #ffffff;
            }
            QLabel {
                color: #ffffff;
                font-weight: 500;
            }
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: 500;
                min-width: 100px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
            QPushButton#btn_importar {
                background-color: #2196F3;
            }
            QPushButton#btn_importar:hover {
                background-color: #1976D2;
            }
            QPushButton#btn_deletar {
                background-color: #f44336;
            }
            QPushButton#btn_deletar:hover {
                background-color: #d32f2f;
            }
            QPushButton#btn_pdf {
                background-color: #FF9800;
            }
            QPushButton#btn_pdf:hover {
                background-color: #F57C00;
            }
            QLineEdit, QDateEdit {
                padding: 6px;
                border: 1px solid #555;
                border-radius: 4px;
                background-color: #3a3a3a;
                color: #ffffff;
            }
            QLineEdit:focus, QDateEdit:focus {
                border: 2px solid #2196F3;
                background-color: #444;
            }
            QTableWidget {
                background-color: #1e1e1e;
                border: 1px solid #444;
                border-radius: 4px;
                gridline-color: #333;
                selection-background-color: #2196F3;
                color: #ffffff;
            }
            QTableWidget::item {
                padding: 8px;
                color: #ffffff;
            }
            QTableWidget::item:selected {
                background-color: #2196F3;
                color: white;
            }
            QTableWidget::item:alternate {
                background-color: #252525;
            }
            QHeaderView::section {
                background-color: #1565C0;
                color: white;
                padding: 12px 10px;
                border: 1px solid #0D47A1;
                font-weight: 600;
                font-size: 10pt;
                min-height: 40px;
                text-align: left;
            }
            QHeaderView {
                background-color: #1565C0;
                min-height: 40px;
            }
            QFrame {
                background-color: #333333;
                border-radius: 8px;
                padding: 10px;
                border: 1px solid #444;
            }
            QCheckBox {
                color: #ffffff;
            }
            QCheckBox::indicator {
                background-color: #444;
                border: 1px solid #666;
                border-radius: 3px;
            }
            QCheckBox::indicator:checked {
                background-color: #2196F3;
                border: 1px solid #1976D2;
            }
        """)

        layout = QVBoxLayout()
        layout.setSpacing(15)
        layout.setContentsMargins(15, 15, 15, 15)

        # Frame para filtros
        frame_filtros = QFrame()
        frame_filtros.setStyleSheet("QFrame { background-color: #333333; border-radius: 8px; padding: 15px; border: 1px solid #444; }")
        filtros_layout = QVBoxLayout()
        filtros_layout.setSpacing(10)
        
        # Título dos filtros
        titulo_filtros = QLabel("🔍 Filtros de Busca")
        titulo_filtros.setStyleSheet("font-size: 14px; font-weight: bold; color: #64B5F6; margin-bottom: 10px;")
        filtros_layout.addWidget(titulo_filtros)
        
        # Linha de filtros - layout responsivo
        filtros = QHBoxLayout()
        filtros.setSpacing(10)
        
        # Data Início
        label_data_inicio = QLabel("📅 Data Início:")
        label_data_inicio.setMinimumWidth(100)
        filtros.addWidget(label_data_inicio)
        self.filtro_data_inicio = QDateEdit()
        self.filtro_data_inicio.setDate(QDate.currentDate().addMonths(-1))
        self.filtro_data_inicio.setCalendarPopup(True)
        self.filtro_data_inicio.setMinimumWidth(120)
        filtros.addWidget(self.filtro_data_inicio)

        # Data Fim
        label_data_fim = QLabel("📅 Data Fim:")
        label_data_fim.setMinimumWidth(100)
        filtros.addWidget(label_data_fim)
        self.filtro_data_fim = QDateEdit()
        self.filtro_data_fim.setDate(QDate.currentDate())
        self.filtro_data_fim.setCalendarPopup(True)
        self.filtro_data_fim.setMinimumWidth(120)
        filtros.addWidget(self.filtro_data_fim)

        # Cliente
        label_cliente = QLabel("👤 Cliente:")
        label_cliente.setMinimumWidth(80)
        filtros.addWidget(label_cliente)
        self.filtro_cliente = QLineEdit()
        self.filtro_cliente.setPlaceholderText("Filtrar por cliente")
        self.filtro_cliente.setMinimumWidth(150)
        filtros.addWidget(self.filtro_cliente)

        # Produto
        label_produto = QLabel("📦 Produto:")
        label_produto.setMinimumWidth(80)
        filtros.addWidget(label_produto)
        self.filtro_produto = QLineEdit()
        self.filtro_produto.setPlaceholderText("Filtrar por produto")
        self.filtro_produto.setMinimumWidth(150)
        filtros.addWidget(self.filtro_produto)

        btn_filtrar = QPushButton("🔍 Filtrar")
        btn_filtrar.setMinimumWidth(100)
        btn_filtrar.clicked.connect(self.carregar_dados)
        filtros.addWidget(btn_filtrar)
        filtros.addStretch()

        filtros_layout.addLayout(filtros)
        frame_filtros.setLayout(filtros_layout)
        layout.addWidget(frame_filtros)

        # Cards de métricas
        self.criar_cards_metricas(layout)
        
        # Botões de ação
        botoes = QHBoxLayout()
        botoes.setSpacing(10)
        
        btn_excel = QPushButton("📥 Importar Excel")
        btn_excel.setObjectName("btn_importar")
        btn_excel.setMinimumWidth(140)
        btn_excel.clicked.connect(self.importar)

        btn_deletar = QPushButton("🗑️ Deletar Selecionadas")
        btn_deletar.setObjectName("btn_deletar")
        btn_deletar.setMinimumWidth(160)
        btn_deletar.clicked.connect(self.deletar_selecionadas)

        btn_pdf = QPushButton("📄 Gerar PDF")
        btn_pdf.setObjectName("btn_pdf")
        btn_pdf.setMinimumWidth(120)
        btn_pdf.clicked.connect(self.exportar_pdf)
        
        btn_pdf_personalizado = QPushButton("📊 PDF Personalizado")
        btn_pdf_personalizado.setObjectName("btn_pdf")
        btn_pdf_personalizado.setMinimumWidth(150)
        btn_pdf_personalizado.clicked.connect(self.exportar_pdf_personalizado)
        
        btn_exportar_excel = QPushButton("📤 Exportar Excel")
        btn_exportar_excel.setStyleSheet("""
            QPushButton {
                background-color: #9C27B0;
            }
            QPushButton:hover {
                background-color: #7B1FA2;
            }
        """)
        btn_exportar_excel.setMinimumWidth(140)
        btn_exportar_excel.clicked.connect(self.exportar_para_excel)

        botoes.addWidget(btn_excel)
        botoes.addWidget(btn_deletar)
        botoes.addWidget(btn_pdf)
        botoes.addWidget(btn_pdf_personalizado)
        botoes.addWidget(btn_exportar_excel)
        botoes.addStretch()

        layout.addLayout(botoes)

        # Abas para tabela e gráficos
        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #444;
                background-color: #2b2b2b;
            }
            QTabBar::tab {
                background-color: #444;
                color: white;
                padding: 8px 20px;
                margin-right: 2px;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #2196F3;
            }
            QTabBar::tab:hover {
                background-color: #555;
            }
        """)
        
        # Aba de Tabela
        tab_tabela = QWidget()
        tab_tabela_layout = QVBoxLayout()
        self.tabela = QTableWidget()
        self.tabela.setMinimumHeight(300)
        tab_tabela_layout.addWidget(self.tabela)
        tab_tabela.setLayout(tab_tabela_layout)
        self.tabs.addTab(tab_tabela, "📋 Tabela de Vendas")
        
        # Aba de Gráficos
        tab_graficos = QWidget()
        tab_graficos_layout = QVBoxLayout()
        
        # Controles de gráficos
        controles_graficos = QHBoxLayout()
        btn_grafico_data = QPushButton("📈 Vendas por Data")
        btn_grafico_data.clicked.connect(lambda: self.mostrar_grafico('data'))
        btn_grafico_produtos = QPushButton("📊 Top Produtos")
        btn_grafico_produtos.clicked.connect(lambda: self.mostrar_grafico('produtos'))
        btn_grafico_clientes = QPushButton("👥 Top Clientes")
        btn_grafico_clientes.clicked.connect(lambda: self.mostrar_grafico('clientes'))
        btn_grafico_pizza = QPushButton("🍰 Distribuição")
        btn_grafico_pizza.clicked.connect(lambda: self.mostrar_grafico('pizza'))
        
        controles_graficos.addWidget(btn_grafico_data)
        controles_graficos.addWidget(btn_grafico_produtos)
        controles_graficos.addWidget(btn_grafico_clientes)
        controles_graficos.addWidget(btn_grafico_pizza)
        controles_graficos.addStretch()
        
        tab_graficos_layout.addLayout(controles_graficos)
        
        # Widget de gráficos
        self.graficos_widget = GraficosWidget()
        tab_graficos_layout.addWidget(self.graficos_widget.canvas)
        tab_graficos.setLayout(tab_graficos_layout)
        self.tabs.addTab(tab_graficos, "📈 Gráficos e Visualizações")
        
        layout.addWidget(self.tabs)

        self.setLayout(layout)
        self.dados = []
        
        # Timer para atualizar métricas a cada 30 segundos
        self.timer_metricas = QTimer()
        self.timer_metricas.timeout.connect(self.atualizar_metricas)
        self.timer_metricas.start(30000)  # 30 segundos
        
        self.carregar_dados()
        self.atualizar_metricas()
    
    def criar_cards_metricas(self, layout):
        """Cria cards com métricas em tempo real"""
        frame_metricas = QFrame()
        frame_metricas.setStyleSheet("QFrame { background-color: #333333; border-radius: 8px; padding: 15px; border: 1px solid #444; }")
        metricas_layout = QGridLayout()
        metricas_layout.setSpacing(15)
        
        # Card 1: Total de Vendas
        card1 = self.criar_card_metrica("💰 Total Vendido", "R$ 0,00", "#4CAF50")
        metricas_layout.addWidget(card1, 0, 0)
        
        # Card 2: Número de Vendas
        card2 = self.criar_card_metrica("📦 Total de Vendas", "0", "#2196F3")
        metricas_layout.addWidget(card2, 0, 1)
        
        # Card 3: Clientes Únicos
        card3 = self.criar_card_metrica("👥 Clientes Únicos", "0", "#FF9800")
        metricas_layout.addWidget(card3, 0, 2)
        
        # Card 4: Produtos Únicos
        card4 = self.criar_card_metrica("📦 Produtos Únicos", "0", "#9C27B0")
        metricas_layout.addWidget(card4, 0, 3)
        
        self.label_total_vendido = card1.findChild(QLabel, "valor")
        self.label_total_vendas = card2.findChild(QLabel, "valor")
        self.label_clientes_unicos = card3.findChild(QLabel, "valor")
        self.label_produtos_unicos = card4.findChild(QLabel, "valor")
        
        frame_metricas.setLayout(metricas_layout)
        layout.addWidget(frame_metricas)
    
    def criar_card_metrica(self, titulo, valor, cor):
        """Cria um card individual de métrica"""
        card = QFrame()
        card.setStyleSheet(f"""
            QFrame {{
                background-color: {cor};
                border-radius: 8px;
                padding: 15px;
                min-width: 200px;
            }}
        """)
        card_layout = QVBoxLayout()
        
        label_titulo = QLabel(titulo)
        label_titulo.setStyleSheet("color: white; font-size: 12px; font-weight: 500;")
        label_titulo.setAlignment(Qt.AlignCenter)
        
        label_valor = QLabel(valor)
        label_valor.setObjectName("valor")
        label_valor.setStyleSheet("color: white; font-size: 20px; font-weight: bold;")
        label_valor.setAlignment(Qt.AlignCenter)
        
        card_layout.addWidget(label_titulo)
        card_layout.addWidget(label_valor)
        card.setLayout(card_layout)
        
        return card
    
    def atualizar_metricas(self):
        """Atualiza as métricas em tempo real"""
        try:
            conn = sqlite3.connect("vendas.db")
            cursor = conn.cursor()
            
            # Aplicar filtros
            query = """
                SELECT COUNT(DISTINCT v.id) as total_vendas,
                       SUM(v.valor_total) as total_valor,
                       COUNT(DISTINCT v.cliente_nome) as clientes_unicos,
                       COUNT(DISTINCT i.produto) as produtos_unicos
                FROM vendas v
                JOIN itens_venda i ON i.venda_id = v.id
                WHERE 1=1
            """
            
            params = []
            
            if hasattr(self, 'filtro_data_inicio') and self.filtro_data_inicio.date():
                query += " AND (v.data_venda IS NULL OR v.data_venda >= ?)"
                params.append(self.filtro_data_inicio.date().toString("yyyy-MM-dd"))
            
            if hasattr(self, 'filtro_data_fim') and self.filtro_data_fim.date():
                query += " AND (v.data_venda IS NULL OR v.data_venda <= ?)"
                params.append(self.filtro_data_fim.date().toString("yyyy-MM-dd"))
            
            if hasattr(self, 'filtro_cliente') and self.filtro_cliente.text():
                query += " AND v.cliente_nome LIKE ?"
                params.append(f"%{self.filtro_cliente.text()}%")
            
            if hasattr(self, 'filtro_produto') and self.filtro_produto.text():
                query += " AND i.produto LIKE ?"
                params.append(f"%{self.filtro_produto.text()}%")
            
            cursor.execute(query, params)
            resultado = cursor.fetchone()
            conn.close()
            
            if resultado:
                total_vendas = resultado[0] or 0
                total_valor = resultado[1] or 0
                clientes_unicos = resultado[2] or 0
                produtos_unicos = resultado[3] or 0
                
                if self.label_total_vendido:
                    self.label_total_vendido.setText(f"R$ {total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
                if self.label_total_vendas:
                    self.label_total_vendas.setText(str(total_vendas))
                if self.label_clientes_unicos:
                    self.label_clientes_unicos.setText(str(clientes_unicos))
                if self.label_produtos_unicos:
                    self.label_produtos_unicos.setText(str(produtos_unicos))
        except Exception as e:
            print(f"Erro ao atualizar métricas: {e}")
    
    def mostrar_grafico(self, tipo):
        """Mostra gráfico do tipo especificado"""
        filtro_data_inicio = self.filtro_data_inicio.date() if hasattr(self, 'filtro_data_inicio') else None
        filtro_data_fim = self.filtro_data_fim.date() if hasattr(self, 'filtro_data_fim') else None
        filtro_cliente = self.filtro_cliente.text() if hasattr(self, 'filtro_cliente') and self.filtro_cliente.text() else None
        filtro_produto = self.filtro_produto.text() if hasattr(self, 'filtro_produto') and self.filtro_produto.text() else None
        
        if tipo == 'data':
            self.graficos_widget.grafico_vendas_por_data(filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
        elif tipo == 'produtos':
            self.graficos_widget.grafico_top_produtos(10, filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
        elif tipo == 'clientes':
            self.graficos_widget.grafico_vendas_por_cliente(10, filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
        elif tipo == 'pizza':
            self.graficos_widget.grafico_pizza_categorias(filtro_data_inicio, filtro_data_fim, filtro_cliente, filtro_produto)
    
    def exportar_pdf_personalizado(self):
        """Exporta PDF com opções personalizadas"""
        dialog = QDialog(self)
        dialog.setWindowTitle("Gerar PDF Personalizado")
        dialog.setMinimumWidth(400)
        layout = QVBoxLayout()
        
        # Tipo de relatório
        layout.addWidget(QLabel("Tipo de Relatório:"))
        combo_tipo = QComboBox()
        combo_tipo.addItems(["completo", "resumo", "por_cliente", "por_produto"])
        layout.addWidget(combo_tipo)
        
        # Opções
        check_resumo = QCheckBox("Incluir Resumo Executivo")
        check_resumo.setChecked(True)
        layout.addWidget(check_resumo)
        
        check_estatisticas = QCheckBox("Incluir Estatísticas")
        check_estatisticas.setChecked(True)
        layout.addWidget(check_estatisticas)
        
        # Botões
        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons)
        
        dialog.setLayout(layout)
        
        if dialog.exec() == QDialog.Accepted:
            arquivo, _ = QFileDialog.getSaveFileName(
                self, "Salvar PDF Personalizado", "relatorio_personalizado.pdf", "*.pdf"
            )
            if arquivo:
                try:
                    filtro_data_inicio = self.filtro_data_inicio.date() if self.filtro_data_inicio.date() else None
                    filtro_data_fim = self.filtro_data_fim.date() if self.filtro_data_fim.date() else None
                    filtro_cliente = self.filtro_cliente.text() if self.filtro_cliente.text() else None
                    filtro_produto = self.filtro_produto.text() if self.filtro_produto.text() else None
                    
                    gerar_pdf_personalizado(
                        arquivo,
                        combo_tipo.currentText(),
                        filtro_data_inicio,
                        filtro_data_fim,
                        filtro_cliente,
                        filtro_produto,
                        check_resumo.isChecked(),
                        check_estatisticas.isChecked()
                    )
                    QMessageBox.information(self, "Sucesso", f"PDF personalizado gerado: {arquivo}")
                except Exception as e:
                    QMessageBox.critical(self, "Erro", f"Erro ao gerar PDF: {str(e)}")
    
    def exportar_para_excel(self):
        """Exporta dados para Excel"""
        arquivo, _ = QFileDialog.getSaveFileName(
            self, "Salvar Excel", "vendas_exportadas.xlsx", "*.xlsx"
        )
        if arquivo:
            try:
                filtro_data_inicio = self.filtro_data_inicio.date() if self.filtro_data_inicio.date() else None
                filtro_data_fim = self.filtro_data_fim.date() if self.filtro_data_fim.date() else None
                filtro_cliente = self.filtro_cliente.text() if self.filtro_cliente.text() else None
                filtro_produto = self.filtro_produto.text() if self.filtro_produto.text() else None
                
                exportar_excel_completo(
                    arquivo,
                    filtro_data_inicio,
                    filtro_data_fim,
                    filtro_cliente,
                    filtro_produto
                )
                QMessageBox.information(self, "Sucesso", f"Excel exportado: {arquivo}")
            except Exception as e:
                QMessageBox.critical(self, "Erro", f"Erro ao exportar Excel: {str(e)}")

    def importar(self):
        arquivo, _ = QFileDialog.getOpenFileName(
            self, "Selecionar Arquivo Excel", "", "Arquivos Excel (*.xlsx *.xls);;Todos os arquivos (*.*)"
        )
        if arquivo:
            try:
                linhas_importadas = importar_excel(arquivo)
                mensagem = f"Dados importados com sucesso!\n\n{linhas_importadas} linha(s) processada(s).\n\n"
                mensagem += "Nota: Vendas e itens duplicados foram automaticamente ignorados."
                QMessageBox.information(
                    self, "Sucesso", mensagem
                )
                self.carregar_dados()
                self.atualizar_metricas()
            except FileNotFoundError as e:
                QMessageBox.critical(
                    self, "Erro", f"Arquivo não encontrado:\n{str(e)}"
                )
            except ValueError as e:
                mensagem_erro = str(e)
                # Verificar se é erro de duplicatas
                if "Todas as linhas já existem" in mensagem_erro or "Nenhum dado novo" in mensagem_erro:
                    QMessageBox.information(
                        self, "Informação", 
                        f"{mensagem_erro}\n\n"
                        "O arquivo foi processado, mas não há dados novos para importar.\n"
                        "Todas as vendas já estão cadastradas no sistema."
                    )
                else:
                    QMessageBox.critical(
                        self, "Erro de Validação", 
                        f"Erro ao validar o arquivo Excel:\n\n{mensagem_erro}\n\n"
                        "Verifique se o arquivo está no formato correto:\n"
                        "- Colunas obrigatórias: Número da venda, Data da venda, Nome do comprador, Título do item\n"
                        "- Os dados devem começar na linha 5 do Excel"
                    )
            except Exception as e:
                QMessageBox.critical(
                    self, "Erro", 
                    f"Erro ao importar o arquivo:\n\n{str(e)}\n\n"
                    "Verifique se:\n"
                    "- O arquivo não está aberto em outro programa\n"
                    "- O arquivo não está corrompido\n"
                    "- O formato do arquivo está correto"
                )

    def carregar_dados(self):
        try:
            conn = sqlite3.connect("vendas.db")
            cursor = conn.cursor()

            query = """
                SELECT v.id, v.cliente_nome, i.produto, i.quantidade, 
                       v.valor_total, v.data_venda, v.numero_venda
                FROM vendas v
                JOIN itens_venda i ON i.venda_id = v.id
                WHERE 1=1
            """

            params = []

            # Filtro de data início
            # Aplicar filtro mas incluir registros sem data para não perder informações
            data_inicio = self.filtro_data_inicio.date()
            data_fim = self.filtro_data_fim.date()
            
            # Se ambos os filtros estão definidos, aplicar normalmente mas incluir NULLs
            if data_inicio and data_fim:
                query += " AND (v.data_venda IS NULL OR (v.data_venda >= ? AND v.data_venda <= ?))"
                params.append(data_inicio.toString("yyyy-MM-dd"))
                params.append(data_fim.toString("yyyy-MM-dd"))
            elif data_inicio:
                query += " AND (v.data_venda IS NULL OR v.data_venda >= ?)"
                params.append(data_inicio.toString("yyyy-MM-dd"))
            elif data_fim:
                query += " AND (v.data_venda IS NULL OR v.data_venda <= ?)"
                params.append(data_fim.toString("yyyy-MM-dd"))

            # Filtro de cliente
            if self.filtro_cliente.text():
                query += " AND v.cliente_nome LIKE ?"
                params.append(f"%{self.filtro_cliente.text()}%")

            # Filtro de produto
            if self.filtro_produto.text():
                query += " AND i.produto LIKE ?"
                params.append(f"%{self.filtro_produto.text()}%")

            cursor.execute(query, params)
            self.dados = cursor.fetchall()
            conn.close()

            # Configurar tabela
            self.tabela.setRowCount(len(self.dados))
            self.tabela.setColumnCount(7)  # Checkbox + 6 colunas de dados
            # Nomes mais curtos para evitar cortes no cabeçalho
            self.tabela.setHorizontalHeaderLabels(
                ["", "Cliente", "Produto", "Qtd", "Valor", "Data", "ID"]
            )
            
            # Ocultar coluna ID (usada internamente)
            self.tabela.setColumnHidden(6, True)

            # Preencher tabela
            for i, linha in enumerate(self.dados):
                # linha = [id, cliente, produto, quantidade, valor, data, numero_venda]
                venda_id = linha[0]
                
                # Checkbox na primeira coluna
                checkbox = QCheckBox()
                checkbox.setStyleSheet("QCheckBox { padding-left: 10px; }")
                self.tabela.setCellWidget(i, 0, checkbox)
                
                # Preencher colunas de dados
                colunas_dados = [
                    (1, linha[1] if len(linha) > 1 else ""),  # Cliente
                    (2, linha[2] if len(linha) > 2 else ""),  # Produto
                    (3, linha[3] if len(linha) > 3 else ""),   # Quantidade
                    (4, linha[4] if len(linha) > 4 else ""),  # Valor
                    (5, linha[5] if len(linha) > 5 else ""), # Data
                ]
                
                for col_idx, valor in colunas_dados:
                    # Formatação especial para valores
                    if col_idx == 4 and valor:  # Valor
                        try:
                            valor_formatado = f"R$ {float(valor):.2f}"
                            item = QTableWidgetItem(valor_formatado)
                            item.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                        except:
                            item = QTableWidgetItem(str(valor) if valor else "")
                    elif col_idx == 5:  # Data
                        if valor:
                            item = QTableWidgetItem(str(valor))
                        else:
                            item = QTableWidgetItem("Sem data")
                            item.setForeground(QColor("#999"))
                    else:
                        item = QTableWidgetItem(str(valor) if valor else "")
                    
                    # Cores alternadas para linhas (fundo escuro)
                    if i % 2 == 0:
                        item.setBackground(QColor("#252525"))
                    else:
                        item.setBackground(QColor("#1e1e1e"))
                    
                    self.tabela.setItem(i, col_idx, item)
                
                # Armazenar ID da venda na última coluna oculta
                id_item = QTableWidgetItem(str(venda_id))
                self.tabela.setItem(i, 6, id_item)

            # Ajustar largura das colunas de forma responsiva
            header = self.tabela.horizontalHeader()
            
            # Configurar modos de redimensionamento
            header.setSectionResizeMode(0, QHeaderView.Fixed)  # Checkbox fixo
            header.setSectionResizeMode(1, QHeaderView.Interactive)  # Cliente ajustável pelo usuário
            header.setSectionResizeMode(2, QHeaderView.Stretch)  # Produto se expande automaticamente
            header.setSectionResizeMode(3, QHeaderView.Fixed)  # Quantidade fixo
            header.setSectionResizeMode(4, QHeaderView.Fixed)  # Valor fixo
            header.setSectionResizeMode(5, QHeaderView.Fixed)  # Data fixo
            
            # Definir larguras fixas GENEROSAS para evitar qualquer corte
            self.tabela.setColumnWidth(0, 50)   # Checkbox
            self.tabela.setColumnWidth(1, 200)  # Cliente - largura generosa
            self.tabela.setColumnWidth(2, 400)  # Produto - muito espaço para textos longos
            self.tabela.setColumnWidth(3, 90)   # Quantidade - espaço suficiente
            self.tabela.setColumnWidth(4, 130)  # Valor - espaço para "R$ XXX.XX"
            self.tabela.setColumnWidth(5, 130)  # Data - espaço para "YYYY-MM-DD"
            
            # Garantir altura mínima do cabeçalho para texto não cortar
            header.setMinimumSectionSize(70)
            header.setDefaultSectionSize(100)
            
            # Desabilitar redimensionamento em cascata que pode causar problemas
            header.setCascadingSectionResizes(False)
            
            # Forçar atualização visual
            self.tabela.update()
            
            # Melhorar visualização
            self.tabela.setAlternatingRowColors(True)
            self.tabela.setSelectionBehavior(QTableWidget.SelectRows)
            self.tabela.setSelectionMode(QTableWidget.ExtendedSelection)
            
            # Habilitar scroll horizontal se necessário
            self.tabela.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
            self.tabela.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)

        except sqlite3.OperationalError as e:
            if "no such table" in str(e):
                QMessageBox.warning(
                    self, 
                    "Aviso", 
                    "Banco de dados ainda não foi criado. Importe um arquivo Excel primeiro."
                )
            else:
                QMessageBox.critical(self, "Erro", f"Erro ao carregar dados: {str(e)}")
        except Exception as e:
            QMessageBox.critical(self, "Erro", f"Erro inesperado: {str(e)}")

    def deletar_selecionadas(self):
        """Deleta as vendas selecionadas via checkbox"""
        linhas_selecionadas = []
        
        for i in range(self.tabela.rowCount()):
            checkbox = self.tabela.cellWidget(i, 0)
            if checkbox and checkbox.isChecked():
                # Pegar o ID da venda (coluna oculta 6)
                id_item = self.tabela.item(i, 6)
                if id_item:
                    linhas_selecionadas.append(int(id_item.text()))
        
        if not linhas_selecionadas:
            QMessageBox.warning(
                self, "Aviso", "Selecione pelo menos uma venda para deletar."
            )
            return
        
        # Confirmar exclusão
        resposta = QMessageBox.question(
            self, "Confirmar Exclusão",
            f"Deseja realmente deletar {len(linhas_selecionadas)} venda(s) selecionada(s)?\n\n"
            "Esta ação não pode ser desfeita!",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        
        if resposta == QMessageBox.Yes:
            try:
                conn = sqlite3.connect("vendas.db")
                cursor = conn.cursor()
                
                # Deletar itens primeiro (foreign key)
                for venda_id in linhas_selecionadas:
                    cursor.execute("DELETE FROM itens_venda WHERE venda_id = ?", (venda_id,))
                    cursor.execute("DELETE FROM vendas WHERE id = ?", (venda_id,))
                
                conn.commit()
                conn.close()
                
                QMessageBox.information(
                    self, "Sucesso", 
                    f"{len(linhas_selecionadas)} venda(s) deletada(s) com sucesso!"
                )
                
                # Recarregar dados e métricas
                self.carregar_dados()
                self.atualizar_metricas()
                self.atualizar_metricas()
                
            except Exception as e:
                QMessageBox.critical(
                    self, "Erro", f"Erro ao deletar vendas: {str(e)}"
                )

    def exportar_pdf(self):
        if not self.dados:
            QMessageBox.warning(
                self, "Aviso", "Não há dados para exportar. Aplique filtros primeiro."
            )
            return

        arquivo, _ = QFileDialog.getSaveFileName(
            self, "Salvar PDF", "relatorio_vendas.pdf", "*.pdf"
        )
        if arquivo:
            try:
                gerar_pdf(self.dados, arquivo)
                QMessageBox.information(
                    self, "Sucesso", f"PDF gerado com sucesso: {arquivo}"
                )
            except Exception as e:
                QMessageBox.critical(
                    self, "Erro", f"Erro ao gerar PDF: {str(e)}"
                )

