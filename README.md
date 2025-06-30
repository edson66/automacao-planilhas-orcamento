# Projeto de Automação de Orçamentos com Python

Este projeto foi desenvolvido com o objetivo de **automatizar a criação de orçamentos comerciais e documentos auxiliares** a partir de planilhas do Excel, utilizando Python e bibliotecas como `openpyxl` e `python-docx`.

---

## 📌 Objetivo

O sistema foi pensado para auxiliar micro e pequenas empresas que lidam com muitos orçamentos baseados em planilhas.  
Ele permite gerar automaticamente diferentes versões de orçamento com **simulação de preços entre fornecedores**, consolidar valores e criar recibos personalizados.

---

## ⚙️ Funcionalidades

- Leitura de dados de uma planilha “doadora”
- Preenchimento automático de até três modelos de orçamento em Excel
- Aplicação automática de variações de preços para simular diferentes fornecedores (entre 15% e 25%)
- Inserção de imagens conforme layout de cada modelo
- Geração opcional de recibo preenchido em `.docx`
- Geração opcional de um documento de consolidação de preços

---

## 🧪 Justificativa para simulação de preços

O sistema aplica variações de preço de forma aleatória dentro de um intervalo entre 15% e 25%.  
Essa prática é usada para **simular cenários reais** onde diferentes fornecedores oferecem valores distintos por produto, permitindo comparar orçamentos de forma automatizada.  
Nenhum dado real de preço ou empresa é utilizado.

---

## 📂 Como usar

1. Prepare sua planilha “doadora” com os produtos, quantidades e preços-base.
2. Coloque os modelos `.xlsx` e `.docx` nas pastas corretas.
3. Execute o script `projeto-planilhas.py`.
4. Responda às perguntas do sistema (CNPJ, data, nota, etc.).
5. Os arquivos gerados serão salvos na pasta `arquivos/`.

---

## 📦 Bibliotecas utilizadas

- `openpyxl` — manipulação de arquivos Excel
- `python-docx` — geração de documentos Word
- `random` — simulação de variação de preços
- `num2words` — conversão de valores numéricos para texto por extenso
- `sys` — controle de execução e encerramento do programa

---

## 👨‍💻 Autor

**Edson Ulisses de Melo Sobrinho**  
Estudante de Sistemas de Informação na UFS  
[LinkedIn](https://www.linkedin.com/in/edson-ulisses-103657372/) | [GitHub](https://github.com/edson66)

---

Este projeto faz parte da minha jornada de aprendizado em Python com foco em automação de tarefas administrativas.
