[![en](https://img.shields.io/badge/lang-en-red.svg)](https://github.com/allanwk/RRGA_scheduler/blob/master/README.md)

Algoritmo de Simulated Annealing para geração de horários escolares
==================

Scripts implementando um algoritmo de Simulated Annealing capaz de gerar horários válidos
para uma escola, dado o número de aulas ministradas por cada professor em cada sala
e restrições de horários.

Método de Geração
------------

`scheduler.cpp` implementa um algoritmo de Simulated Annealing:

- Um horário aleatório é gerado com o número correto de aulas por professor
- A qualidade do horário é determinada com base em vários fatores, como
  conflitos de aulas, agrupamento de aulas, excesso de aulas e restrições
  definidas pelo usuário
- Professores são trocados aleatoriamente utilizando o método de Simulated Annealing,
  através do qual no começo, todas as trocas são aceitas, mas com o passar do tempo,
  apenas trocas que melhoram a qualide do horário são aceitas
- Dessa forma, o algoritmo maximiza a qualidade do horário sendo gerado

Functionalidade
----------------------

- `control.py` gera arquivos csv de acordo com as configurações nas planilhas
- `scheduler.cpp` implementa o algoritmo, salvando os horários gerados em arquivos de texto
- `converter.py` então converte a saída do programa anterior em uma planilha
  contendo o horário gerado

Entradas e Saídas
----------------------

- Exempo de horário produzido pelo algoritmo:

![schematic](images/horario.png)

- [Configuração do número de aulas por professor](https://github.com/allanwk/RRGA_scheduler/blob/master/images/n_aulas_por_professor.png)
- [Configuração do número de matérias por professor](https://github.com/allanwk/RRGA_scheduler/blob/master/images/materias_por_professor.png)
- [Exemplo de restrição complexa de horário para um professor](https://github.com/allanwk/RRGA_scheduler/blob/master/images/restricoes.png)
- [Exemplo de restrição de horário para um professor](https://github.com/allanwk/RRGA_scheduler/blob/master/images/restricoes_simples.png)

Features
----------------------

- Definição de restrições de horário com pesos customizáveis para cada professor
- Tipo de restrição especial - professor só pode dar a última aula
- Possibilidade de definir aulas constantes - todos os horários gerados terão
  aquelas aulas nos locais exatos como foi definido