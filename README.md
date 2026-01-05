# Sistema de Controle de Vendas - Desktop

Sistema desktop completo para controle e análise de vendas, desenvolvido em Python com PySide6.

## 🚀 Funcionalidades

- ✅ **Importação de Excel**: Importe dados de vendas diretamente de arquivos Excel
- ✅ **Dashboard Interativo**: Visualize vendas em tabela com filtros avançados
- ✅ **Filtros Avançados**: Filtre por data, cliente e produto
- ✅ **Geração de PDF**: Exporte relatórios profissionais em PDF
- ✅ **Banco SQLite**: Armazenamento local e eficiente
- ✅ **Interface Moderna**: Interface desktop intuitiva e responsiva

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
   - Clique em "Gerar PDF"
   - Escolha onde salvar o arquivo

## 📁 Estrutura do Projeto

```
sistema_vendas/
│
├── app.py              # Ponto de entrada da aplicação
├── database.py         # Configuração do banco de dados
├── models.py           # Modelos de tabelas (SQLAlchemy)
├── import_excel.py     # Função de importação de Excel
├── dashboard.py        # Interface gráfica principal
├── pdf_report.py       # Geração de relatórios PDF
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

- [ ] Gráficos e visualizações
- [ ] Sistema de login e permissões
- [ ] Exportação para Excel
- [ ] Relatórios personalizados
- [ ] Dashboard com métricas em tempo real



Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.

