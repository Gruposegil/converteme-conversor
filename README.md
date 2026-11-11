# Converteme — servidor de conversão

Servidor pequeno que o site [converteme.com.br](https://converteme.com.br) usa nas
ferramentas que não dá para fazer no navegador:

| Ferramenta | Como |
| --- | --- |
| PDF para Word (.docx) | [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) (MIT) + [PyMuPDF](https://github.com/pymupdf/PyMuPDF) (AGPL-3.0); relatórios em colunas com `layout.py` (próprio) |
| PDF para Excel (.xlsx) | `planilha.py` (próprio) com PyMuPDF e [openpyxl](https://foss.heptapod.net/openpyxl/openpyxl) (MIT) |
| PDF para PowerPoint (.pptx) | [LibreOffice](https://www.libreoffice.org/) (MPL-2.0), importação de PDF do Impress: um slide por página, texto em caixas editáveis |
| PDF para PDF/A (1b ou 2b) | [Ghostscript](https://www.ghostscript.com/) (AGPL-3.0), com perfil de cor sRGB (OutputIntent); antes, o PyMuPDF tira JavaScript e anexos |
| Excel, PowerPoint, Word (.docx e .doc), OpenDocument para PDF | [LibreOffice](https://www.libreoffice.org/) (MPL-2.0), sem interface |

Roda no **Google Cloud Run, região São Paulo** (`southamerica-east1`).

## Privacidade

- O arquivo chega, vai para uma pasta temporária só daquela conversão, e a pasta é
  apagada assim que a resposta sai, com erro ou sem.
- O registro (log) guarda só o tipo da conversão, o tamanho arredondado e o tempo.
  Nunca o nome do arquivo nem o conteúdo.
- Só atende pedidos com a chave combinada (variável `CHAVE`). O site fala com este
  servidor por um Worker da Cloudflare, que guarda a chave; o navegador nunca vê a chave.

## Endereços

| Método | Caminho | O que faz |
| --- | --- | --- |
| GET | `/saude` | responde `ok` |
| POST | `/converter?para=docx` | corpo = PDF → devolve .docx |
| POST | `/converter?para=xlsx` | corpo = PDF → devolve .xlsx |
| POST | `/converter?para=pptx` | corpo = PDF → devolve .pptx |
| POST | `/converter?para=pdfa&nivel=2` | corpo = PDF → devolve PDF/A-2b (`nivel=1`: PDF/A-1b) |
| POST | `/converter?para=pdf&ext=.xlsx` | corpo = Word/Excel/PowerPoint/OpenDocument → devolve PDF |

O cabeçalho de resposta `X-Modo` diz como foi feito: `completo`, `posicao` ou `simples` (PDF para Word);
`pdfa-2` ou `pdfa-1` (PDF/A), com `-imagem-N` no fim quando N páginas com transparência tiveram de
virar imagem no PDF/A-1b (o PDF/A-1 não aceita transparência).

Cabeçalho obrigatório: `X-Chave: <a chave>`. Erros voltam em JSON
(`{"erro": "...", "mensagem": "..."}`, em português).

## Configuração (variáveis do Cloud Run)

| Variável | Padrão | Para quê |
| --- | --- | --- |
| `CHAVE` | (vazia = recusa tudo) | chave combinada com o Worker. **Nunca colocar no código.** |
| `MAX_MB` | 30 | tamanho máximo do arquivo (o Cloud Run aceita até 32 MiB por pedido) |
| `MAX_PAGINAS` | 300 | páginas máximas de um PDF para Word |
| `PRAZO_SEGUNDOS` | 85 | tempo máximo de uma conversão do LibreOffice (Office → PDF) |
| `PRAZO_DOCX_SEGUNDOS` | 75 | tempo máximo da conversão detalhada de PDF para Word; passou disso, sai a conversão simples |
| `MEMORIA_DOCX_MB` | 900 | teto de memória de cada processo da conversão detalhada (0 = sem teto); estourou, sai a "no lugar" |
| `PRAZO_XLSX_SEGUNDOS` | 85 | tempo máximo do PDF para Excel |
| `PROCESSOS_DOCX` | 2 | processos da conversão detalhada (só usa mais de 1 se houver mais de 1 CPU) |

Os prazos somados ficam abaixo de 100 s, que é quanto a Cloudflare espera uma resposta.

### PDF para Word: qual conversão é usada
1. **Relatório em colunas** (de sistema de RH, estoque, contas; sem grade de tabela):
   vai direto para a conversão **"no lugar"** (`layout.py`). Cada linha vira um parágrafo na
   mesma altura e cada coluna fica na mesma posição, por tabulação. Fica igual ao PDF e
   sai em segundos.
2. **Os outros** vão para a conversão **detalhada** (pdf2docx), que junta as linhas em parágrafos
   e faz tabela de verdade quando a tabela tem bordas. Ela roda num processo separado, com
   prazo e teto de memória, para um PDF complicado não derrubar o servidor.
3. Se a detalhada passar do prazo ou falhar, sai a **"no lugar"**. Se até ela falhar, sai a
   **simples** (texto corrido).

O cabeçalho `X-Modo` da resposta diz qual foi usada: `completo`, `posicao` ou `simples`
(no PDF para Excel, `planilha`).

### PDF para Excel
- Tabela com grade (linhas desenhadas): `page.find_tables()` do PyMuPDF; cada célula vira célula.
- Relatório sem grade: cada linha do PDF vira uma linha da planilha. As colunas saem do
  alinhamento: texto que começa no mesmo x, ou número que termina no mesmo x, fica na mesma
  coluna em todas as páginas. Colunas vizinhas que nunca aparecem juntas na mesma linha viram
  uma só. Dois espaços ou mais dentro de um texto também separam colunas.
- Valores brasileiros viram valor de verdade: `1.234,56` e `R$ 10,00` (número), `12,5%`
  (porcentagem), `01/10/2026` (data). Código com zero à esquerda, CPF, CNPJ e números com 11
  dígitos ou mais continuam texto.
- Todas as páginas na mesma aba, uma depois da outra.

## Testes

```
pip install -r requirements.txt python-docx python-pptx pikepdf
python3 -m unittest discover -s testes -v
```
(precisa do LibreOffice instalado, com o Draw e o Impress: `soffice`; e do Ghostscript: `gs`)

O PDF/A é conferido nos testes pela identificação (XMP), pelo perfil de cor e pelas fontes
embutidas. Não foi conferido com um validador completo (veraPDF).

## Licença

AGPL-3.0 (arquivo `LICENSE`), porque usa o PyMuPDF e o Ghostscript. Quem usa o servidor pela internet
tem direito ao código-fonte: é este repositório.
