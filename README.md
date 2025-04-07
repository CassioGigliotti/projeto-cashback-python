# 🧾 Cálculo de Cashback por Subcategoria - Automação com Python e SQL

Este projeto realiza a análise de vendas dos últimos 90 dias utilizando dados do banco `AdventureWorksDW2022`, confronta essas informações com regras de cashback definidas em um arquivo Excel e gera um relatório com os clientes que têm direito ao cashback.

## 🔧 Tecnologias utilizadas

- Python (Pandas, SQLAlchemy, Datetime)
- Jupyter Notebook
- SQL Server (AdventureWorksDW2022)
- Excel

## 🧠 Lógica do projeto

1. Conexão com o banco de dados SQL Server
2. Extração das vendas dos últimos 90 dias por subcategoria e cliente
3. Leitura de um Excel com as regras de cashback (valor mínimo e percentual por subcategoria)
4. Verificação de quais clientes atingiram o valor mínimo
5. Cálculo do valor de cashback
6. Geração de um novo Excel com os valores devidos por cliente

## 📁 Estrutura do projeto

```plaintext
projeto-cashback-python/
│── Cashback.ipynb # Notebook principal com toda a automação
│── Cashback.xlsx # Arquivo com regras de cashback
│── Clientes e Cashback.xlsx # Arquivo gerado apos automação
│── README.md # Este arquivo
```

## 📊 Exemplo de regra no Excel  

| Subcategoria | Valor Mínimo | % Cashback |  
|-------------|--------------|------------|  
| Bikes       | 10.000       | 0.05       |  
| Accessories | 5.000        | 0.03       |  

## 📤 Resultado gerado  

Um arquivo Excel com o nome do cliente, subcategoria e valor de cashback a receber.  

## 📘 Acesse o notebook principal  

🔗 [Clique aqui para abrir o notebook no GitHub](Cashback.ipynb)  

## 🙋‍♂️ Autor  

Cássio Gigliotti  

🔗 [Meu portfólio de Power BI](https://app.xperiun.com/in/cassio-gigliotti)  
🔗 [LinkedIn](https://www.linkedin.com/in/cassio-gigliotti/)  
