# Converteme — servidor de conversão

Servidor pequeno que o site [converteme.com.br](https://converteme.com.br) usa nas
ferramentas que não dá para fazer no navegador:

| Ferramenta | Como |
| --- | --- |
| PDF para Word (.docx) | [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) (MIT) + [PyMuPDF](https://github.com/pymupdf/PyMuPDF) (AGPL-3.0) |
| Excel, PowerPoint, Word, OpenDocument para PDF | [LibreOffice](https://www.libreoffice.org/) (MPL-2.0), sem interface |

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
| POST | `/converter?para=pdf&ext=.xlsx` | corpo = Word/Excel/PowerPoint/OpenDocument → devolve PDF |

Cabeçalho obrigatório: `X-Chave: <a chave>`. Erros voltam em JSON
(`{"erro": "...", "mensagem": "..."}`, em português).

## Configuração (variáveis do Cloud Run)

| Variável | Padrão | Para quê |
| --- | --- | --- |
| `CHAVE` | (vazia = recusa tudo) | chave combinada com o Worker. **Nunca colocar no código.** |
| `MAX_MB` | 30 | tamanho máximo do arquivo (o Cloud Run aceita até 32 MiB por pedido) |
| `MAX_PAGINAS` | 300 | páginas máximas de um PDF para Word |
| `PRAZO_SEGUNDOS` | 150 | tempo máximo de uma conversão do LibreOffice |

Recomendado no Cloud Run: 1 CPU, 2 GiB de memória, até 2 pedidos ao mesmo tempo por
cópia, no máximo 3 cópias, mínimo 0 (só paga quando usa), prazo de 300 s.

## Testes

```
pip install -r requirements.txt python-docx
python3 -m unittest discover -s testes -v
```
(precisa do LibreOffice instalado: `soffice`)

## Licença

AGPL-3.0 (arquivo `LICENSE`), porque usa o PyMuPDF. Quem usa o servidor pela internet
tem direito ao código-fonte: é este repositório.
