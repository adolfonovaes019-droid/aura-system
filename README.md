# AURA SYSTEM — PDV

Sistema de ponto de venda (PDV) para lanchonetes e restaurantes, desenvolvido em Python com interface gráfica nativa.

## Funcionalidades

- **Vendas** — grade de produtos por categoria, carrinho com adicionais, observações por item
- **Clientes** — cadastro com endereço, bairro, bloco/apto e taxa de entrega por cliente
- **Pedidos** — histórico do dia com filtro por data, edição e reimpressão de comandas
- **Financeiro** — dashboard com totais por forma de pagamento (Dinheiro, Cartão, PIX) e impressão de relatório de caixa
- **Produtos** — CRUD completo do cardápio com categorias (Lanches, Porções, Refrigerantes, Outros)
- **Adicionais** — gerenciamento dos extras que aparecem ao montar um pedido
- **Impressão térmica** — suporte a impressoras RAW (80mm/60mm) com comandos ESC/POS
- **Temas visuais** — 4 paletas de cores configuráveis
- **Imagem de fundo** — personalize a tela de vendas com uma imagem

## Pré-requisitos

- Python 3.8+
- Windows (a impressão térmica usa `pywin32`)

## Instalação

```bash
pip install tkcalendar pillow pywin32
```

> `pywin32` é opcional. Sem ele o sistema roda em modo simulação (sem impressão física).

## Como usar

```bash
python main.py
```

Na primeira execução, informe o nome da loja e uma senha para o módulo financeiro. Todos os dados ficam salvos em:

```
C:\Users\<seu_usuario>\AppData\Local\AURA SYSTEM\
```

## Estrutura dos dados

O sistema usa SQLite (`aura_system.db`) com as seguintes tabelas:

| Tabela | Conteúdo |
|--------|----------|
| `clientes` | Cadastro de clientes |
| `produtos` | Cardápio |
| `vendas` | Histórico de pedidos |
| `adicionais` | Extras dos lanches |
| `config_loja` | Nome, senha, CNPJ, etc. |

## Dependências

| Pacote | Uso |
|--------|-----|
| `tkinter` | Interface gráfica (incluso no Python) |
| `tkcalendar` | Seletor de data nos filtros |
| `Pillow` | Imagem de fundo na tela de vendas |
| `pywin32` | Comunicação com impressora térmica |
| `sqlite3` | Banco de dados (incluso no Python) |

## Desenvolvido por

Adolfo Rodrigues Novaes
