# 🤖 Automação de Contrato de Trabalho 🤖

Imagina que você precisa fazer vários contratos de trabalho dos funcionários novos que entrarão na sua empresa. 
Para uma pessoa será bem demorado abrir um contrato de trabalho modelo, adicionar os dados de casa pessoa, salvar e ir para outro.

Com este programa, só necessita de um banco de dados e ele irá automatizar para você, em poucos segundos crie centenas de contratos de trabalho

---

# ⚙️Como está funcionando a automação? ⚙️

Da linha 30 até a linha 68 estou criando um banco de dados com dados "fakes" para usar de exemplo ao gerar os contratos.
Logo após temos o contrato de trabalho modelo, onde dentro tem as especificações de onde serão recolocados as informações de cada funcionário.
O programa abrirá o banco de dados na linha 89
```wb = load_workbook('banco_de_dados_fake.xlsx')```
Aqui você dará o nome, no nosso caso banco_de_dados_fake.

Ele vai pegar as informações de cada funcionário, passará para a classe "Employee" criada na linha 10, vai abrir o contrato de trabalho inserir as informações da classe no documento, e vai salvar dentro de uma pasta
chamada "CONTRATOS GERADOS", com a função e nome do funcionário.

---

# 🔨 Tem uma empresa e precisa usar o programa para automatizar sua vida? Veja como utilizar: 🔨

- **_💾 1° CONTRATO MODELO:_**     
      Deixe um contrato modelo na pasta onde está o arquivo "main.py" com nome 'Modelo Contrato de Trabalho'
  
- **_💾 2° DADOS DO EMPREGADO:_**    
      Na classe "EMPLOYEE" adicione ou remova os dados necessários de cada funcionario que seu contrato de trabalho está solicitando, e na linha 114 pra baixo, adicione ou remova os dados de acordo
      com seu banco de dados, por exemplo:
  
| RG | CPF | NOME |
|------------|-------------|------------|
| 1234567    | 333.222.111 | ANA CLARA  |
| 1234567    | 333.222.111 | ANA CLARA  |    

    Deixe somente RG, CPF e NOME na classe employee.    
- **_💾 3° MODELE O JEITO QUE IRÁ COLETAR OS DADOS:_**
      Nas linhas 103 à 112 vemos que estamos coletando os dados, por exemplo ```employee.name = ws[f'A{i}'].value``` significa que estamos pegando na coluna "A", linha de numero "i" onde i vai de 2 (primeira linha abaixo da coluna nome até o ultimo funcionário, se for 20 funcionários vai até a linha 21 e estamos dando o valor que pegamos da lacuna do excel e inserimos para a variavel "employee.name", para que logo abaixo no documento (Contrato Modelo) ele irá procurar onde está escrito "employee.name" e irá colocar o valor da lacuna do excel ```paragraph.text = paragraph.text.replace("employee.name", employee.name)```

---
# 📷 Screenshots do funcionamento 📷

Veja as fotos abaixo: 

a primeira mostra como ficou o banco de dados criado, 

![image](https://github.com/xBigJhow/Automacao-Contrato-de-Trabalho/assets/103526432/00909a91-788d-47bb-854b-c99900f6f40c)


e logo após os contratos já criados,

![image](https://github.com/xBigJhow/Automacao-Contrato-de-Trabalho/assets/103526432/0ad2b8a1-5b9c-4e8b-9e60-c683440640e4)


e por último como ficou os dados inseridos automáticamente.

![image](https://github.com/xBigJhow/Automacao-Contrato-de-Trabalho/assets/103526432/19b3ccf3-791d-41ca-8290-501188bb9c1d)


---
# 📖🤓 AUTOR 📖🤓

Feitor por **BIGJHOW**

#### Caso deseja alguma explicação ou alinhamento para seu uso, entre em contato, serei grato em poder ajudar de alguma forma.
---
