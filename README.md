# Sistema-de-Inventario-Ciclico
Sistema de inventário cíclico.


### Objetivo  📋
Realizar a amostragem de materiais a serem inventariados periodicamente, considerando o tempo do período entre cada inventário para cada depósito, critério de não repetição de itens e critério de auditoria completa de materiais dentro do exercício (ano). </br>

### Linguagem Utilizada 📚
Visual Basic for Application (VBA) </br>
## Descrição Detalhada do Sistema
O sistema irá possuir dois modos de uso, sendo eles o modo manual e automático. </br>
### Modo Manual ⚙️
O modo manual é utilizado para momentos em que apenas um depósito precise ser inventariado. Uma lista está disponível para que o depósito desejado seja selecionado antes de inciciar o processo. </br>
![image](https://github.com/GabrielHirt/Sistema-de-Inventario-Ciclico/assets/98654562/9025c4fe-82af-4675-87a1-0e3d6477b9e0)

### Modo Automático ⚙️
O modo automático é utilizado para que todos os depósitos sejam inventariados de uma única vez. </br>
![image](https://github.com/GabrielHirt/Sistema-de-Inventario-Ciclico/assets/98654562/64d86459-1b0c-429f-9ff4-b3eca75abe35)

### Lógica do Sistema 🛠️
O sistema foi feito com o objetivo de realizar um inventário semanal para cada depósito existente. Para cada depósito dependendo da quantidade de itens presentes, irá selecionar de forma aleatória um número de itens, de forma que, ao final do mês ou ano, aquele estoque esteja zerado.
Estoques neste caso poderão ser inventáriados por:
- Dia.
- Por mês.
- Ano.
A quantidade de itens em cada depósito ou a necessidade do negócio irá determinar o intervalo de tempo e necessidade para cada depósito.

Ao final dos prazos, sendo eles por dia, mês ou ano, todos os itens terão sidos inventáriados.

A contagem dos itens deve ocorrer semanal.

Para cada "rodada" (vez que o inventário ser executado), o código irá verificar se itens novos tiveram entrada.

Para cada depósito ocorra uma reposição de seus itens ao final de um cíclo. Sendo cada cíclo determinada pela lógica de tempo que será levado para a contagem total do depósito, como mencionado para este caso, teremos cíclos que terão seu valor de contagem total contado dentro de um dia, outros 1 mês e ainda 1 ano.

## Demonstração Em Vídeo
No vídeo abaixo, o modo de inventário automático está sendo usado para inventáriar todos os depósitos desejados de uma única vez.
Todo o processo de manipulação dos dados não é aparente no vídeo, pois é utilizado um comando em VBA para que as atualizações de tela sejam desligadas até que o processo seja concluído.


https://github.com/GabrielHirt/Sistema-de-Inventario-Ciclico/assets/98654562/8a6a1730-e47f-4caf-9890-fe6fce240cb5






