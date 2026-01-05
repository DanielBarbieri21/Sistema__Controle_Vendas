from sqlalchemy import (
    Table, Column, Integer, String, Float, Date, ForeignKey, MetaData
)
from database import engine

metadata = MetaData()

vendas = Table(
    "vendas", metadata,
    Column("id", Integer, primary_key=True),
    Column("numero_venda", String),
    Column("data_venda", Date),
    Column("cliente_nome", String),
    Column("cliente_cidade", String),
    Column("cliente_estado", String),
    Column("status_venda", String),
    Column("valor_total", Float)
)

itens_venda = Table(
    "itens_venda", metadata,
    Column("id", Integer, primary_key=True),
    Column("venda_id", Integer, ForeignKey("vendas.id")),
    Column("produto", String),
    Column("quantidade", Integer),
    Column("preco_unitario", Float)
)

metadata.create_all(engine)

