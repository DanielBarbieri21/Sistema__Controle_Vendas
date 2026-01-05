from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from datetime import datetime

def gerar_pdf(dados, caminho="relatorio_vendas.pdf"):
    # Criar documento com margens maiores para evitar cortes
    doc = SimpleDocTemplate(caminho, pagesize=A4,
                            rightMargin=1.5*cm, leftMargin=1.5*cm,
                            topMargin=2*cm, bottomMargin=2.5*cm)
    
    # Container para elementos
    elementos = []
    estilos = getSampleStyleSheet()
    
    # Estilo para título
    estilo_titulo = ParagraphStyle(
        'TituloCustom',
        parent=estilos['Heading1'],
        fontSize=18,
        textColor=colors.HexColor('#1565C0'),
        spaceAfter=12,
        alignment=1  # Centralizado
    )
    
    # Estilo para data
    estilo_data = ParagraphStyle(
        'DataCustom',
        parent=estilos['Normal'],
        fontSize=10,
        textColor=colors.HexColor('#666666'),
        spaceAfter=20,
        alignment=1  # Centralizado
    )
    
    # Título
    titulo = Paragraph("Relatório de Vendas", estilo_titulo)
    elementos.append(titulo)
    
    # Data de geração
    data_geracao = Paragraph(
        f"Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
        estilo_data
    )
    elementos.append(data_geracao)
    elementos.append(Spacer(1, 0.5*cm))
    
    # Preparar dados da tabela
    dados_tabela = []
    total_geral = 0
    
    # Cabeçalho
    cabecalho = [
        Paragraph("<b>Cliente</b>", estilos['Normal']),
        Paragraph("<b>Produto</b>", estilos['Normal']),
        Paragraph("<b>Qtde</b>", estilos['Normal']),
        Paragraph("<b>Valor</b>", estilos['Normal']),
        Paragraph("<b>Data</b>", estilos['Normal'])
    ]
    dados_tabela.append(cabecalho)
    
    # Dados
    for linha in dados:
        # linha = [id, cliente, produto, quantidade, valor_total, data_venda, numero_venda]
        # Índices: 0=id, 1=cliente, 2=produto, 3=quantidade, 4=valor_total, 5=data_venda, 6=numero_venda
        cliente = str(linha[1]) if len(linha) > 1 and linha[1] else ""
        produto = str(linha[2]) if len(linha) > 2 and linha[2] else ""
        quantidade = str(linha[3]) if len(linha) > 3 and linha[3] else "0"
        
        # Valor
        try:
            valor_num = float(linha[4]) if len(linha) > 4 and linha[4] else 0
            valor_str = f"R$ {valor_num:.2f}"
            total_geral += valor_num
        except:
            valor_str = "R$ 0.00"
        
        # Data
        data = str(linha[5])[:10] if len(linha) > 5 and linha[5] else "Sem data"
        
        # Limitar tamanho dos textos longos para evitar cortes
        # Ajustar conforme a largura disponível da coluna
        cliente = cliente[:35] + "..." if len(cliente) > 35 else cliente
        produto = produto[:45] + "..." if len(produto) > 45 else produto
        
        linha_tabela = [
            Paragraph(cliente, estilos['Normal']),
            Paragraph(produto, estilos['Normal']),
            Paragraph(quantidade, estilos['Normal']),
            Paragraph(valor_str, estilos['Normal']),
            Paragraph(data, estilos['Normal'])
        ]
        dados_tabela.append(linha_tabela)
    
    # Adicionar linha de total ANTES de criar a tabela
    linha_total = [
        Paragraph("<b>Total Geral</b>", estilos['Normal']),
        Paragraph("", estilos['Normal']),
        Paragraph("", estilos['Normal']),
        Paragraph(f"<b>R$ {total_geral:.2f}</b>", estilos['Normal']),
        Paragraph("", estilos['Normal'])
    ]
    dados_tabela.append(linha_total)
    
    # Calcular largura disponível (A4 width - margens)
    # A4 width = 21cm, margens = 1.5cm cada lado = 3cm total
    largura_disponivel = 21*cm - (1.5*cm * 2)  # 18cm disponível
    
    # Distribuir larguras das colunas proporcionalmente para caber na página
    # Cliente: 28%, Produto: 42%, Qtde: 10%, Valor: 10%, Data: 10%
    col_widths = [
        largura_disponivel * 0.28,  # Cliente (~5cm)
        largura_disponivel * 0.42,  # Produto (~7.5cm)
        largura_disponivel * 0.10,  # Qtde (~1.8cm)
        largura_disponivel * 0.10,  # Valor (~1.8cm)
        largura_disponivel * 0.10   # Data (~1.8cm)
    ]
    
    # Criar tabela com todas as linhas incluindo o total
    tabela = Table(dados_tabela, colWidths=col_widths, repeatRows=1)
    
    # Estilo da tabela
    estilo_tabela = TableStyle([
        # Cabeçalho
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1565C0')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('ALIGN', (2, 0), (2, -1), 'CENTER'),  # Quantidade centralizada
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),   # Valor alinhado à direita
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        
        # Linhas alternadas (exceto cabeçalho e total)
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.HexColor('#f5f5f5')]),
        
        # Bordas
        ('GRID', (0, 0), (-1, -1), 1, colors.grey),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        
        # Padding das células (reduzido para evitar cortes)
        ('LEFTPADDING', (0, 0), (-1, -1), 4),
        ('RIGHTPADDING', (0, 0), (-1, -1), 4),
        ('TOPPADDING', (0, 1), (-1, -2), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -2), 6),
        
        # Estilo para linha de total (última linha)
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#E3F2FD')),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, -1), (-1, -1), 11),
        ('ALIGN', (3, -1), (3, -1), 'RIGHT'),
        ('TOPPADDING', (0, -1), (-1, -1), 12),
        ('BOTTOMPADDING', (0, -1), (-1, -1), 12),
    ])
    
    tabela.setStyle(estilo_tabela)
    elementos.append(tabela)
    
    # Construir PDF
    doc.build(elementos)

