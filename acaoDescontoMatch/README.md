# Escopo 
## 1 - Contexto

Se faz necessário uma ação de desconto de 40% em determinados produtos, neste caso produtos da categoria MATCH. Os dados que foram enviados foram a base de tarefas que estavam disponiveis para cada ponto de venda e os respectivos skus que validavam a tarefa separados por vírgula 

## 2 - Lógica para Clusterização da tarefas

Antes de começar os trabalhos de análise, foi realizado uma limpeza da base para um melhor entendimento da situação. 

- havia duas ou mais tarefas para o mesmo pdv. Deixou-se apenas a tarefa mais abrangente para evitar duplicidade desnecessária no arquivo

-  validação de quais eram as tarefas únicas existentes e as mesmas foram colocadas em clusters
   ![image](https://github.com/user-attachments/assets/d54d09cb-b97b-440b-8216-abd45cc6aec6)
  
- A justifica de colocar as tarefas que pediam SKUs distintos no cluster de MATCH TOTAL foi devido ao formato do arquivo para subir o desconto, no qual não há essa possibilidade de contabilizar SKUs distintos. Então liberou para a base total, deixando a cargo do Representante de Negócios realizar a validação da melhor forma

- Neste momento também foi percebido uma dificuldade e dilema em relação a tarefas nos clusters de volume. Perguntas como: Como colocar essa quantidade esperada de caixas na ação visto que ação só permite comprar uma vez e é pouco provável que o cliente compre uma quantidade alta de uma vez

## 3 - Breve explicação dos códigos 

- De início, o foco inicial foi realizar a separação dos skus uma tentativa de capturar a quantidade de caixas pedidas para cada tarefa e ligar isto ao sku solicitado. Retornando assim, em uma outra planilha a seguinte estrutura para cada coluna : 

   -    produto | pdv | quantidade

  Foi realizado desta forma, pois o template solicitado para subir era :

   -   agrupador | operação | pdv

  *o agrupador é um código que precisei para lincar o sku. 

  *operação é o código que indica em qual operação ou Revenda a ação será submetida

Para isto foi utilizado o código 1 e 2. A diferença entre eles é que um tem o filtro de Cluster e o outro não, deixando o processo mais automatizado. 

## 4 - Lições para o futuro

   - eu não preciso criar um agrupador para lincar ao sku, posso usar o próprio numero do sku como agrupador e seguir a vida, pois não tem restrição de ter que começar no 1,2,3,4, ....
     
   - não focar energia no que é minímo. Foquei tanto esforço na questão de saber o volume da caixa e de como eu colocaria no template solicitado que perdi tempo e uns neurônios a toa pensando em como atuar, sendo que já poderia ter feito o que era o mais básico antes e deixar rodando.
