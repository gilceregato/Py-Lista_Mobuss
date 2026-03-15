# Py-Lista_Mobuss
Aplicação em Python para geração automática de uma Lista Mestra de Documentos padronizada a partir de um arquivo Excel exportado da plataforma Mobuss.
O objetivo do projeto é automatizar a transformação da listagem de arquivos da plataforma em uma estrutura documental consistente, adequada para controle de documentos, auditoria e gestão de projetos.

## Contexto
Plataformas de gestão de arquivos como o Mobuss permitem exportar listagens de documentos em formato Excel. Entretanto, essas exportações normalmente:
- Não seguem uma estrutura de lista mestra documental
- Não classificam adequadamente os documentos por disciplina ou tipo

Este projeto resolve esse problema ao processar automaticamente o Excel exportado, aplicando regras de padronização e gerando uma planilha estruturada de controle documental.

## Principais Funcionalidades
1. Importação de Dados
- Leitura de arquivos .xlsx
- Identificação automática de colunas relevantes
- Extração de metadados dos documentos

2. Normalização
- Aplicação de regras de padronização
- Classificação por disciplina
- Identificação de revisões

3. Geração da Lista Mestra
Produção de uma planilha estruturada

A aplicação segue uma arquitetura simples baseada em pipeline de processamento de dados.

Excel Export (Mobuss)
        │
        ▼
   Parser (Leitura)
        │
        ▼
 Normalizer (Regras)
        │
        ▼
 Generator (Lista Mestra)
        │
        ▼
 Excel Final Padronizado
 
## Estrutura do Repositório (em revisão)
project/
│
├── input/
│   └── mobuss_export.xlsx
│
├── output/
│   └── lista_mestra.xlsx
│
├── rules/
│   └── normalization_rules.yaml
│
├── src/
│   ├── parser.py
│   ├── normalizer.py
│   ├── generator.py
│   └── utils.py
│
├── main.py
├── requirements.txt
└── README.md

## Tecnologias Utilizadas
- Python 3.10+
- pandas — manipulação de dados
- openpyxl — leitura/escrita de Excel

# Instalação (em desenvolvimento)

# Uso  (em desenvolvimento)

# Roadmap  (em desenvolvimento)
Melhorias planejadas:
1. interface CLI avançada
2. validação automática de nomenclatura
3. integração direta com API do Mobuss
4. exportação para banco de dados

# Contribuição
Contribuições são bem-vindas.

# Licença
Este projeto está licenciado sob a licença MIT.
