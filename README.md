# Sistema de Controle de Vendas - Desktop

Sistema desktop completo para controle e análise de vendas, desenvolvido em Python com PySide6.

## 🚀 Funcionalidades

- ✅ **Importação de Excel**: Importe dados de vendas diretamente de arquivos Excel
- ✅ **Dashboard Interativo**: Visualize vendas em tabela com filtros avançados
- ✅ **Filtros Avançados**: Filtre por data, cliente e produto
- ✅ **Geração de PDF**: Exporte relatórios profissionais em PDF
- ✅ **📊 Gráficos e Visualizações**: Visualize dados com gráficos interativos (linha, barras, pizza)
- ✅ **📤 Exportação para Excel**: Exporte dados filtrados para Excel com múltiplas abas e formatação profissional
- ✅ **📄 Relatórios Personalizados**: Gere PDFs com opções personalizadas (completo, resumo, por cliente, por produto)
- ✅ **📈 Dashboard com Métricas em Tempo Real**: Cards com métricas atualizadas automaticamente a cada 30 segundos
- ✅ **Banco SQLite**: Armazenamento local e eficiente
- ✅ **Interface Moderna**: Interface desktop intuitiva e responsiva com tema escuro

## 📋 Requisitos

- Python 3.8 ou superior
- pip (gerenciador de pacotes Python)

## 🔧 Instalação

1. Clone ou baixe este repositório

2. Instale as dependências:
```bash
pip install -r requirements.txt
```

## ▶️ Como Usar

1. Execute o aplicativo:
```bash
python app.py
```

2. **Importar dados do Excel**:
   - Clique em "Importar Excel"
   - Selecione o arquivo Excel com os dados de vendas
   - O sistema criará automaticamente o banco de dados e importará os dados

3. **Filtrar dados**:
   - Use os campos de filtro (Data, Cliente, Produto)
   - Clique em "Filtrar" para aplicar os filtros
   - Os resultados aparecerão na tabela

4. **Gerar PDF**:
   - Aplique os filtros desejados
   - Clique em "Gerar PDF" para relatório simples
   - Ou clique em "PDF Personalizado" para escolher tipo de relatório e opções

5. **Exportar para Excel**:
   - Aplique os filtros desejados
   - Clique em "Exportar Excel"
   - O arquivo será gerado com múltiplas abas (Vendas Detalhadas, Resumo por Cliente, Resumo por Produto)

6. **Visualizar Gráficos**:
   - Vá para a aba "Gráficos e Visualizações"
   - Clique nos botões para visualizar diferentes tipos de gráficos:
     - **Vendas por Data**: Gráfico de linha mostrando evolução temporal
     - **Top Produtos**: Gráfico de barras horizontais com os produtos mais vendidos
     - **Top Clientes**: Gráfico de barras com os principais clientes
     - **Distribuição**: Gráfico de pizza com distribuição de vendas por produto

7. **Métricas em Tempo Real**:
   - As métricas no topo da tela são atualizadas automaticamente a cada 30 segundos
   - Elas refletem os dados filtrados atualmente
   - Mostram: Total Vendido, Total de Vendas, Clientes Únicos e Produtos Únicos

## 📁 Estrutura do Projeto

```
sistema_vendas/
│
├── app.py              # Ponto de entrada da aplicação
├── database.py         # Configuração do banco de dados
├── models.py           # Modelos de tabelas (SQLAlchemy)
├── import_excel.py     # Função de importação de Excel
├── dashboard.py        # Interface gráfica principal com métricas e gráficos
├── pdf_report.py       # Geração de relatórios PDF (simples e personalizados)
├── export_excel.py     # Exportação de dados para Excel
├── charts.py           # Módulo de gráficos e visualizações
├── vendas.db           # Banco de dados SQLite (criado automaticamente)
└── requirements.txt    # Dependências do projeto
```

## 📊 Formato do Excel

O arquivo Excel deve ter as seguintes colunas (a partir da linha 5):
- Número da venda
- Data da venda
- Nome do comprador
- Cidade
- Estado
- Status da venda
- Título do item
- Quantidade
- Preço unitário
- Valor total

## 🎯 Gerar Executável (.exe)

Para criar um executável Windows:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "SistemaVendas" app.py
```

O arquivo `.exe` estará em `dist/SistemaVendas.exe`

## 🔄 Próximas Melhorias

- [ ] Sistema de login e permissões
- [ ] Mais tipos de gráficos (scatter, heatmap)
- [ ] Exportação de gráficos como imagem
- [ ] Filtros salvos e favoritos
- [ ] Notificações de novas vendas



Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

