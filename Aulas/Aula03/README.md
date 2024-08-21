## Como dados são coletados e quem gera os dados coletados?

Em Excel e Google Sheets, os dados são coletados de diversas maneiras, e quem gera esses dados pode variar.

### **1. Como os dados são coletados?**

**Manual:**

- **Usuário:** A pessoa que está utilizando o Excel ou Google Sheets pode inserir os dados manualmente. Por exemplo, digitar nomes, números, datas, etc.
- **Copiar/Colar:** O usuário pode copiar dados de outro lugar (como um site, outro documento, ou um e-mail) e colar no Excel ou Google Sheets.

**Importação de arquivos:**

- **Arquivos CSV/Excel:** Você pode importar dados que já estão em outro formato, como um arquivo CSV (que é basicamente um arquivo de texto com dados separados por vírgulas) ou um arquivo Excel de outro sistema.

**Integrações Automáticas:**

- **Formulários Google:** No Google Sheets, por exemplo, você pode criar um formulário e as respostas das pessoas que preencherem o formulário são automaticamente coletadas na planilha.
- **Conexões com bancos de dados ou sistemas:** Algumas empresas conectam planilhas com seus sistemas internos, e os dados são atualizados automaticamente.

### **2. Quem gera os dados coletados?**

**Usuários:**

- Muitas vezes, são as pessoas que trabalham com as planilhas que geram os dados ao digitarem ou copiarem informações.

**Sistemas Automatizados:**

- Sistemas podem ser programados para enviar dados para o Excel ou Google Sheets automaticamente, sem a necessidade de intervenção humana.

**Outras Ferramentas:**

- **API (Interface de Programação de Aplicações):** Ferramentas podem ser integradas para enviar dados para as planilhas de forma automatizada.

### **Exemplo Simplificado:**

Imagine que você está gerenciando uma lista de convidados para uma festa:

- **Você** (usuário) cria uma planilha no Google Sheets e digita o nome de cada convidado.
- **Outros amigos** (usuários) que também têm acesso à planilha podem adicionar os nomes de outros convidados.
- Você pode criar um formulário Google para que as pessoas confirmem presença, e as respostas são **automaticamente** coletadas na planilha.

Aqui, **você e seus amigos geram os dados** (nomes dos convidados), e o **Google Forms coleta as respostas automaticamente** na planilha.

Isso é o básico de como os dados são coletados e quem gera esses dados no Excel e Google Sheets!

## **Buscando erros**

Buscar erros em planilhas como Excel ou Google Sheets é uma prática comum para garantir que os dados estejam corretos e que as fórmulas funcionem conforme o esperado. Aqui está uma explicação simples de como fazer isso:

### **1. Erros de Entrada de Dados**

- **Verificação Visual:** Olhe para os dados inseridos e veja se há algo que não parece certo, como números fora do lugar ou palavras digitadas incorretamente.
- **Ordenação e Filtros:** Ordenar os dados ou usar filtros pode ajudar a identificar entradas duplicadas ou anômalas.

### **2. Erros de Fórmulas**

- **Referências Inválidas:** Se uma fórmula se refere a uma célula ou intervalo que não existe, Excel e Google Sheets geralmente mostrarão um erro como `#REF!`.
- **Erros Comuns:** 
  - `#DIV/0!`: Tentar dividir por zero.
  - `#VALUE!`: Quando a fórmula espera um tipo de dado (como um número) e recebe outro (como texto).
  - `#NAME?`: Quando você usa o nome de uma função ou referência de célula que não é reconhecida.

### **3. Ferramentas para Identificar Erros**

- **Excel:**
  - **Revisar Fórmulas:** Excel tem uma ferramenta chamada "Auditoria de Fórmulas" que permite ver as dependências das células e identificar onde uma fórmula pode estar falhando.
  - **Verificação de Erros:** No menu de "Fórmulas", você pode clicar em "Verificação de Erros" para que o Excel procure automaticamente erros comuns.

- **Google Sheets:**
  - **Mostrar Fórmulas:** Assim como no Excel, você pode alternar para mostrar as fórmulas em vez dos resultados, o que facilita a identificação de onde algo pode estar errado.
  - **Assistente de Fórmulas:** O Google Sheets sugere correções automáticas para erros comuns quando detecta que uma fórmula pode estar errada.

### **4. Uso de Funções para Verificar Erros**

- **`IFERROR`:** Essa função permite que você "capture" erros e substitua por algo mais amigável. Por exemplo, `=IFERROR(A1/B1, "Erro")` evita que o erro `#DIV/0!` apareça, substituindo-o por "Erro".
- **`ISNUMBER`, `ISTEXT`:** Estas funções podem ser usadas para verificar se uma célula contém o tipo de dado esperado.

### **5. Validação de Dados**

- **Regras de Validação:** Você pode configurar regras para garantir que apenas certos tipos de dados sejam inseridos, como limitar entradas a números positivos ou restringir texto a um certo formato.
- **Mensagens de Erro:** Se alguém tentar inserir algo que não cumpre as regras, uma mensagem de erro pode ser exibida, ajudando a evitar problemas futuros.

### **6. Teste com Dados Conhecidos**

- **Comparação:** Insira alguns dados cujo resultado você já conhece e veja se as fórmulas ou operações retornam o valor esperado.

Esses métodos ajudam a garantir que as planilhas sejam precisas e que qualquer erro seja identificado e corrigido rapidamente.

## **Função UNIQUE**

A função `UNIQUE` em Excel e Google Sheets é usada para extrair valores únicos de uma lista ou intervalo de dados. Em outras palavras, ela remove duplicatas e retorna apenas os valores que aparecem uma vez.

### **Como Usar a Função `UNIQUE`**

#### **1. Em Google Sheets:**

A sintaxe básica é:

```excel
=UNIQUE(intervalo)
```

- **`intervalo`**: É o intervalo de células do qual você deseja extrair valores únicos.

**Exemplo:**

Se você tem uma lista de nomes na coluna A (de A1 a A10), e alguns nomes se repetem, você pode usar:

```excel
=UNIQUE(A1:A10)
```

Isso retornará uma lista de nomes onde cada nome aparece apenas uma vez, sem duplicatas.

#### **2. Em Excel:**

No Excel, a função `UNIQUE` funciona de maneira similar, mas está disponível apenas para usuários do Microsoft 365 ou Excel 2021 e posteriores. A sintaxe também é:

```excel
=UNIQUE(intervalo)
```

**Exemplo:**

Se você tem uma lista de números de B1 a B15 e deseja extrair apenas os números únicos, você usaria:

```excel
=UNIQUE(B1:B15)
```

### **Opções Adicionais (Google Sheets):**

A função `UNIQUE` no Google Sheets também pode ter opções adicionais para especificar se deseja verificar as colunas ou as linhas por duplicatas:

```excel
=UNIQUE(intervalo, [by_column])
```

- **`intervalo`**: O intervalo de dados.
- **`by_column`** (opcional): Se for TRUE, a função considera duplicados por coluna. Se for FALSE ou omitido, considera por linha.

### **O que a Função `UNIQUE` Retorna?**

- **Valores Únicos:** A função retorna uma matriz (ou lista) de valores únicos.
- **Matriz Dinâmica:** Se você usa `UNIQUE` em uma célula, o Excel ou Google Sheets automaticamente expande os resultados para as células ao lado, conforme necessário.

### **Exemplos Práticos:**

1. **Lista de Clientes Sem Repetição:**
   Se você tem uma lista de clientes com alguns nomes repetidos e deseja ver apenas os nomes únicos:

   ```excel
   =UNIQUE(C1:C100)
   ```

2. **Filtrar Produtos Diferentes:**
   Se você tem uma lista de produtos vendidos e quer ver quais são os diferentes tipos de produtos vendidos:

   ```excel
   =UNIQUE(D1:D50)
   ```

A função `UNIQUE` é bastante útil para limpar e organizar dados, removendo duplicatas e facilitando a análise de informações.

A classificação ou ordenação alfabética em Excel e Google Sheets permite organizar seus dados em ordem crescente (A a Z) ou decrescente (Z a A). Isso é útil para ordenar listas de nomes, produtos, categorias, entre outros.

### **1. Como Classificar em Excel**

#### **Classificação Simples:**

1. **Selecione os Dados:** Clique e arraste para selecionar a coluna ou intervalo de dados que você deseja ordenar.
2. Vá para a Aba "Dados":
   - No Excel, clique na aba **"Dados"** no menu superior.
3. Classificar de A a Z ou Z a A:
   - Para ordem crescente, clique em **"Classificar de A a Z"** (ícone com A acima do Z).
   - Para ordem decrescente, clique em **"Classificar de Z a A"** (ícone com Z acima do A).

#### **Classificação Personalizada:**

1. Selecione a Tabela:
   - Se os seus dados têm cabeçalhos, selecione a tabela inteira, incluindo os cabeçalhos.
2. Opção de Classificação:
   - Na aba **"Dados"**, clique em **"Classificação Personalizada..."**.
3. Escolha a Coluna:
   - Na janela que aparece, escolha a coluna pela qual você deseja ordenar.
   - Selecione se deseja ordenar em ordem crescente ou decrescente.
4. **Classificar:** Clique em **"OK"**.

### **2. Como Classificar em Google Sheets**

#### **Classificação Simples:**

1. **Selecione os Dados:** Clique e arraste para selecionar a coluna ou intervalo de dados que você deseja ordenar.
2. Menu de Dados:
   - No Google Sheets, vá para o menu **"Dados"** no topo.
3. Ordenar de A a Z ou Z a A:
   - Clique em **"Classificar intervalo de A a Z"** para uma ordem crescente.
   - Clique em **"Classificar intervalo de Z a A"** para uma ordem decrescente.

#### **Classificação com Cabeçalhos:**

1. Selecione a Tabela:
   - Selecione toda a tabela, incluindo os cabeçalhos.
2. Menu de Dados:
   - No menu **"Dados"**, escolha **"Classificar intervalo..."**.
3. Selecionar a Coluna:
   - Na janela que aparece, marque a opção **"Os dados têm linha de cabeçalho"**.
   - Escolha a coluna pela qual você deseja ordenar e a ordem (A a Z ou Z a A).
4. **Classificar:** Clique em **"Ordenar"**.

### **Exemplo Prático:**

Imagine que você tem uma lista de nomes na coluna A, e quer ordenar essa lista alfabeticamente:

- **Em Excel:** Selecione a coluna A, vá em "Dados", e clique em "Classificar de A a Z".
- **Em Google Sheets:** Selecione a coluna A, vá em "Dados", e clique em "Classificar intervalo de A a Z".

### **Dicas Adicionais:**

- **Ordenar Várias Colunas:** Se você tiver uma tabela com várias colunas (como nomes e datas de nascimento), pode ordenar primeiro por uma coluna (como nomes) e depois por outra (como datas).
- **Cuidado com as Linhas:** Quando ordenar dados, certifique-se de que todas as colunas correspondentes sejam ordenadas em conjunto, para não misturar os dados.

Ordenar dados de forma alfabética ajuda a organizar a informação de maneira que seja mais fácil de ler e analisar.

## **Filtros**

Filtros em Excel e Google Sheets são ferramentas poderosas que permitem visualizar, analisar e trabalhar com subconjuntos específicos de seus dados, ocultando temporariamente as informações que você não precisa. Eles são úteis quando você tem grandes volumes de dados e deseja focar em informações específicas.

### **1. Como Aplicar Filtros em Excel**

#### **Aplicar Filtros Simples:**

1. **Selecione a Tabela:**
   - Clique em qualquer célula dentro do intervalo de dados que você deseja filtrar. Se você quiser aplicar filtros a toda a tabela, certifique-se de que todas as colunas tenham cabeçalhos (títulos).

2. **Vá para a Aba "Dados":**
   - Na aba **"Dados"** do Excel, clique em **"Filtro"**. Isso adicionará pequenas setas de filtro ao lado de cada cabeçalho de coluna.

3. **Aplicar o Filtro:**
   - Clique na seta ao lado do cabeçalho da coluna que você deseja filtrar.
   - Uma lista suspensa aparecerá mostrando todos os valores únicos dessa coluna. Desmarque os valores que você não deseja ver ou use as caixas de seleção para escolher os valores específicos.
   - Clique em **"OK"** para aplicar o filtro. Apenas as linhas que correspondem aos critérios selecionados serão exibidas.

#### **Filtros Avançados:**

1. **Filtro por Condições:**
   - Na lista suspensa de filtro, você também pode filtrar por condições específicas, como números maiores que, textos que contêm uma palavra específica, datas antes ou depois de certo dia, etc.

2. **Filtro de Pesquisa:**
   - Se você está procurando algo específico, pode digitar o valor que deseja na caixa de pesquisa que aparece na parte superior da lista suspensa.

3. **Remover Filtros:**
   - Para remover os filtros e mostrar todos os dados novamente, clique novamente no botão **"Filtro"** na aba **"Dados"**.

### **2. Como Aplicar Filtros em Google Sheets**

#### **Aplicar Filtros Simples:**

1. **Selecione a Tabela:**
   - Clique em qualquer célula dentro do intervalo de dados que você deseja filtrar.

2. **Adicionar um Filtro:**
   - Vá para o menu **"Dados"** no topo da página e selecione **"Criar um filtro"**. Isso adicionará setas de filtro ao lado de cada cabeçalho de coluna.

3. **Aplicar o Filtro:**
   - Clique na seta ao lado do cabeçalho da coluna que você deseja filtrar.
   - Uma lista suspensa aparecerá, onde você pode selecionar ou desmarcar valores, semelhante ao Excel.
   - Você também pode usar a barra de pesquisa para encontrar valores específicos.

4. **Filtro por Condições:**
   - No Google Sheets, você pode filtrar por condições, como textos que contêm uma palavra específica, números maiores que um certo valor, ou datas em um intervalo específico. Isso pode ser feito clicando em "Filtrar por condição" dentro da lista suspensa.

#### **Remover ou Desativar Filtros:**

- Para remover um filtro, vá para o menu **"Dados"** e selecione **"Remover filtro"**. Isso restaurará a visualização completa dos dados.

### **Exemplos Práticos:**

- **Filtrar Produtos por Categoria:**
  Se você tem uma tabela de produtos e deseja ver apenas os produtos de uma categoria específica, aplique um filtro na coluna "Categoria" e selecione apenas a categoria desejada.

- **Filtrar Transações por Data:**
  Se você tem uma lista de transações e quer ver apenas as que ocorreram em um determinado mês, aplique um filtro na coluna "Data" e use a condição de filtro para escolher o intervalo de datas.

### **Dicas Adicionais:**

- **Combinação de Filtros:** Você pode aplicar filtros em várias colunas ao mesmo tempo para refinar ainda mais os dados que você vê.
- **Filtros e Classificação:** Filtros e classificação podem ser usados juntos. Por exemplo, você pode filtrar para mostrar apenas certos dados e, em seguida, classificar esses dados filtrados em ordem alfabética ou numérica.
- **Proteção de Dados:** Lembre-se de que ao filtrar os dados, você está apenas ocultando as informações que não atendem aos critérios, e não excluindo-as.

Filtros são uma excelente maneira de simplificar e focar nos dados que são realmente importantes para a análise que você está fazendo.

## **Como validamos as coisas no nosso dia a dia?**

A validação no nosso dia a dia envolve verificar, confirmar ou testar algo para garantir que seja verdadeiro, correto ou apropriado. Esse processo de validação acontece de diversas formas, e muitas vezes, fazemos isso de forma automática ou inconsciente. Aqui estão alguns exemplos de como validamos as coisas no cotidiano:

### **1. Validação de Informação**

- **Checar Fontes:** Quando ouvimos uma notícia, frequentemente verificamos em várias fontes para garantir que seja verdadeira. Por exemplo, se alguém nos diz que haverá um evento, podemos procurar online, checar com outras pessoas ou consultar o organizador.
- **Conferir Fatos:** Se você recebe uma informação que parece duvidosa, como um boato ou uma notícia em redes sociais, você pode pesquisar para ver se outros meios de comunicação confirmam a mesma coisa.

### **2. Validação de Decisões**

- **Consulta com Outros:** Antes de tomar uma decisão importante, como comprar um carro ou escolher um curso, é comum conversar com amigos, familiares ou especialistas para validar se a escolha é boa.
- **Experiência Anterior:** Muitas vezes validamos uma decisão com base em experiências passadas. Por exemplo, se você já usou um produto de uma marca e teve uma boa experiência, você valida que comprar da mesma marca novamente pode ser uma boa decisão.

### **3. Validação de Processos**

- **Revisão e Conferência:** Antes de enviar um e-mail importante ou concluir uma tarefa, normalmente revisamos o conteúdo para garantir que esteja correto. Isso pode incluir checar a ortografia, verificar se os anexos estão incluídos, ou se as informações estão precisas.
- **Testar Antes de Usar:** Antes de usar um equipamento novo, como um eletrodoméstico, muitas vezes fazemos um teste para garantir que funcione corretamente. Por exemplo, ao comprar um celular, você pode verificar se todos os aplicativos funcionam como esperado.

### **4. Validação de Relacionamentos**

- **Observação de Comportamentos:** Em relações pessoais ou profissionais, validamos a confiança e integridade das pessoas observando seus comportamentos ao longo do tempo. Por exemplo, se um colega sempre cumpre prazos e é confiável, você valida que pode confiar nele em projetos futuros.
- **Comunicação:** Pedimos feedback ou confirmamos com as pessoas se entenderam corretamente o que foi dito ou solicitado, o que ajuda a evitar mal-entendidos.

### **5. Validação Financeira**

- **Checagem de Saldo:** Antes de fazer uma compra, você pode verificar o saldo da sua conta bancária para validar se tem dinheiro suficiente.
- **Comparação de Preços:** Ao comprar algo, muitas vezes validamos o valor comparando preços em diferentes lojas para garantir que estamos fazendo um bom negócio.

### **6. Validação de Sentimentos**

- **Reflexão Interna:** Muitas vezes validamos nossos sentimentos ou reações refletindo sobre eles. Por exemplo, se você se sente desconfortável em uma situação, pode se perguntar por que se sente assim e se a reação é justificada.
- **Feedback Externo:** Podemos validar como nos sentimos em relação a algo conversando com amigos ou familiares para entender se nossos sentimentos são comuns ou apropriados.

### **7. Validação de Segurança**

- **Verificação de Portas e Janelas:** Antes de sair de casa ou dormir, é comum validar a segurança verificando se portas e janelas estão trancadas.
- **Cuidado em Situações Desconhecidas:** Quando estamos em um lugar desconhecido, validamos a segurança observando o ambiente, verificando se é seguro e, se necessário, pedindo informações.

### **8. Validação de Resultados**

- **Comparação com Expectativas:** Depois de completar uma tarefa ou projeto, muitas vezes comparamos os resultados com o que esperávamos para validar se tudo saiu conforme o planejado.
- **Avaliação de Desempenho:** No trabalho ou nos estudos, validamos nosso desempenho através de avaliações, feedbacks, ou comparando com metas e objetivos estabelecidos.

### **Conclusão**

A validação no dia a dia é um processo contínuo de checar, confirmar e testar para garantir que estamos tomando as decisões corretas, lidando com informações precisas e interagindo de maneira apropriada. Isso ajuda a minimizar erros, evitar mal-entendidos e garantir que nossas ações estejam alinhadas com nossos objetivos e valores.

## **Validação de dados**

A validação de dados é um processo importante tanto em ferramentas como Excel e Google Sheets quanto em sistemas mais complexos, e tem como objetivo garantir que os dados inseridos em uma planilha ou banco de dados estejam corretos, coerentes e dentro de certos critérios predefinidos. Esse processo ajuda a evitar erros e garantir que os dados sejam úteis e confiáveis para análises e decisões.

### **Por que a Validação de Dados é Importante?**

- **Prevenção de Erros:** A validação de dados ajuda a evitar a entrada de informações incorretas ou incoerentes, como números fora de um intervalo aceitável, datas inválidas, ou texto em vez de números.
- **Garantia de Consistência:** Assegura que os dados sejam consistentes e sigam um formato padrão, o que facilita a análise e o processamento posterior.
- **Aumento da Confiabilidade:** Dados validados são mais confiáveis, o que melhora a qualidade das análises e decisões baseadas nesses dados.

### **Como Fazer Validação de Dados em Excel**

1. **Selecione as Células para Validação:**
   - Selecione as células ou intervalo de células onde você deseja aplicar a validação de dados.

2. **Acesse a Ferramenta de Validação:**
   - Vá até a aba **"Dados"** e clique em **"Validação de Dados"**. Uma janela com opções de validação será aberta.

3. **Defina os Critérios de Validação:**
   - **Permitir:** Escolha o tipo de dado que será permitido, como números, datas, texto, listas, etc.
   - **Critérios:** Dependendo do tipo de dado selecionado, você poderá definir critérios específicos. Por exemplo:
     - Para **números**, você pode definir um intervalo mínimo e máximo.
     - Para **datas**, você pode restringir a entrada a um determinado intervalo de datas.
     - Para **listas**, você pode fornecer uma lista de opções pré-definidas que o usuário deve escolher.

4. **Mensagens de Erro:**
   - **Mensagem de Entrada:** Você pode configurar uma mensagem que aparecerá quando o usuário selecionar a célula, explicando as regras de validação.
   - **Mensagem de Erro:** Configure uma mensagem de erro personalizada que aparecerá se o usuário tentar inserir dados que não atendem aos critérios.

5. **Aplicar a Validação:**
   - Clique em **"OK"** para aplicar as regras de validação. Agora, qualquer dado inserido nas células selecionadas será verificado de acordo com os critérios estabelecidos.

### **Como Fazer Validação de Dados em Google Sheets**

1. **Selecione as Células para Validação:**
   - Selecione as células ou intervalo de células onde você deseja aplicar a validação de dados.

2. **Acesse a Ferramenta de Validação:**
   - Vá para o menu **"Dados"** e selecione **"Validação de Dados"**.

3. **Defina os Critérios de Validação:**
   - **Critérios:** Escolha os critérios de validação, como um intervalo de números, uma lista de itens, ou um texto específico.
   - **Intervalo de Células:** Você pode definir critérios baseados em intervalos de células, como permitir apenas valores que estão em uma lista predefinida em outro lugar na planilha.

4. **Opções de Validação:**
   - **Mostrar Lista suspensa:** Se você escolher validar com uma lista, marque a opção para mostrar uma lista suspensa, facilitando a escolha para o usuário.
   - **Aviso ou Rejeitar:** Escolha se o sistema apenas dará um aviso ou se rejeitará a entrada que não cumprir os critérios.
   - **Mensagem de Ajuda:** Adicione uma mensagem personalizada para ajudar o usuário a entender o que é permitido.

5. **Aplicar a Validação:**
   - Clique em **"Salvar"** para aplicar as regras de validação. Agora, as células seguirão as regras definidas ao inserir dados.

### **Exemplos de Uso de Validação de Dados:**

- **Números em um Intervalo:** Garantir que a nota de um aluno esteja entre 0 e 10.
- **Datas:** Garantir que uma data de vencimento esteja no futuro.
- **Listas:** Restringir a entrada a determinados valores, como categorias de produtos.
- **Comprimento do Texto:** Validar que um código de produto tenha exatamente 8 caracteres.

### **Conclusão**

A validação de dados é uma prática essencial para garantir a qualidade dos dados em planilhas e sistemas. Ela ajuda a prevenir erros e garante que as informações inseridas sejam consistentes, precisas e dentro das expectativas. Isso facilita a análise posterior, aumenta a confiabilidade dos dados e apoia a tomada de decisões mais informada.

## **Mais recursos e formatos condicionais**

O uso de texto em planilhas como Excel e Google Sheets é muito comum para armazenar e organizar informações não numéricas, como nomes, endereços, descrições de produtos, e muito mais. Existem várias funções e técnicas específicas para trabalhar com texto em planilhas, que ajudam a manipular, formatar e extrair informações úteis.

### **1. Armazenamento de Texto**

- **Células de Texto:** Em planilhas, você pode simplesmente digitar texto em qualquer célula. O texto pode ser de qualquer tamanho e pode incluir letras, números e símbolos.
- **Formatação de Texto:** Você pode formatar o texto usando opções como negrito, itálico, sublinhado, alterar a cor da fonte, ajustar o tamanho da fonte e alinhar o texto à esquerda, centro ou direita da célula.

### **2. Funções para Manipular Texto**

#### **a. CONCATENAR/CONCAT**

- **Descrição:** Junta dois ou mais pedaços de texto de diferentes células em uma única célula.
- **Exemplo:**
  - Se você tem "João" na célula A1 e "Silva" na célula B1, você pode usar `=CONCATENAR(A1, " ", B1)` para obter "João Silva".
  - No Google Sheets e Excel mais recentes, você pode usar `=CONCAT(A1, " ", B1)`.

#### **b. ESQUERDA, DIREITA, e EXT.TEXTO**

- **Descrição:** Extraem partes específicas de um texto.
  - **ESQUERDA:** Extrai um número específico de caracteres do início de um texto.
    - Exemplo: `=ESQUERDA(A1, 3)` extrairá os três primeiros caracteres do texto em A1.
  - **DIREITA:** Extrai um número específico de caracteres do final de um texto.
    - Exemplo: `=DIREITA(A1, 4)` extrairá os quatro últimos caracteres do texto em A1.
  - **EXT.TEXTO:** Extrai um número específico de caracteres a partir de uma posição específica no texto.
    - Exemplo: `=EXT.TEXTO(A1, 2, 5)` começará no segundo caractere de A1 e extrairá cinco caracteres.

#### **c. LOCALIZAR/PROCURAR**

- **Descrição:** Encontram a posição de um texto dentro de outro texto.
  - **LOCALIZAR:** Sensível a maiúsculas e minúsculas.
  - **PROCURAR:** Não sensível a maiúsculas e minúsculas.
  - Exemplo: `=LOCALIZAR("Silva", A1)` retorna a posição onde "Silva" começa no texto em A1.

#### **d. MAIÚSCULA, MINÚSCULA, e PRI.MAIÚSCULA**

- **Descrição:** Alteram o caso das letras em um texto.
  - **MAIÚSCULA:** Converte todo o texto para letras maiúsculas.
    - Exemplo: `=MAIÚSCULA(A1)` transforma "joão silva" em "JOÃO SILVA".
  - **MINÚSCULA:** Converte todo o texto para letras minúsculas.
    - Exemplo: `=MINÚSCULA(A1)` transforma "JOÃO SILVA" em "joão silva".
  - **PRI.MAIÚSCULA:** Converte a primeira letra de cada palavra para maiúscula.
    - Exemplo: `=PRI.MAIÚSCULA(A1)` transforma "joão silva" em "João Silva".

#### **e. SUBSTITUIR e TROCAR**

- **Descrição:** Substituem partes específicas de um texto.
  - **SUBSTITUIR:** Substitui texto com base em um termo exato.
    - Exemplo: `=SUBSTITUIR(A1, "Silva", "Oliveira")` substitui "Silva" por "Oliveira" em A1.
  - **TROCAR:** Substitui texto, mas é mais flexível ao permitir a substituição parcial e repetida.
    - Exemplo: `=TROCAR(A1, "a", "o")` substitui todas as ocorrências de "a" por "o" em A1.

### **3. Funções Específicas para Google Sheets**

#### **a. SPLIT**

- **Descrição:** Divide o texto em várias células com base em um delimitador.
  - Exemplo: `=SPLIT(A1, " ")` dividirá o texto em A1 em células separadas onde houver um espaço.

#### **b. JOIN**

- **Descrição:** Junta um intervalo de células em uma única célula, separando cada valor por um delimitador.
  - Exemplo: `=JOIN(", ", A1:A5)` combinará o texto das células de A1 a A5, separando cada valor por uma vírgula e um espaço.

### **4. Uso Prático de Texto nas Planilhas**

- **Gerenciamento de Listas:** Você pode organizar listas de nomes, endereços, produtos, etc., e utilizar funções como CONCAT ou JOIN para combinar informações.
- **Extração de Dados:** Usar ESQUERDA, DIREITA, ou EXT.TEXTO para extrair partes de códigos ou números de série.
- **Formatação e Normalização:** Usar MAIÚSCULA, MINÚSCULA, e PRI.MAIÚSCULA para padronizar a formatação de texto.
- **Pesquisa e Substituição:** Utilizar LOCALIZAR, PROCURAR, SUBSTITUIR e TROCAR para pesquisar e substituir informações rapidamente.

### **Conclusão**

Trabalhar com texto em Excel e Google Sheets oferece uma vasta gama de possibilidades para manipular e organizar dados textuais de maneira eficiente. Usando as funções e técnicas descritas, você pode transformar, analisar e apresentar suas informações textuais de forma que seja mais útil e acessível.

## **Funções condicionais**

Funções condicionais em Excel e Google Sheets permitem que você realize cálculos ou execute ações com base em determinadas condições. Essas funções são extremamente úteis para análise de dados, tomadas de decisão e automação de processos dentro das planilhas. Aqui estão as funções condicionais mais comuns e como usá-las:

### **1. Função `SE` (IF)**

#### **Descrição:**

A função `SE` (em inglês, `IF`) é a função condicional mais básica. Ela verifica uma condição e retorna um valor se a condição for verdadeira e outro valor se for falsa.

#### **Sintaxe:**

```excel
=SE(teste_lógico, valor_se_verdadeiro, valor_se_falso)
```

- **`teste_lógico`:** A condição que você quer verificar (por exemplo, `A1 > 10`).
- **`valor_se_verdadeiro`:** O valor que será retornado se o teste for verdadeiro.
- **`valor_se_falso`:** O valor que será retornado se o teste for falso.

#### **Exemplo:**

```excel
=SE(A1 > 10, "Maior que 10", "Menor ou igual a 10")
```

- Se o valor em A1 for maior que 10, a função retorna "Maior que 10". Caso contrário, retorna "Menor ou igual a 10".

### **2. Função `SEERRO` (IFERROR)**

#### **Descrição:**

A função `SEERRO` (em inglês, `IFERROR`) é usada para capturar e gerenciar erros em fórmulas. Ela permite que você substitua o erro por um valor ou mensagem personalizada.

#### **Sintaxe:**

```excel
=SEERRO(fórmula, valor_se_erro)
```

- **`fórmula`:** A fórmula que você deseja avaliar.
- **`valor_se_erro`:** O valor que será retornado se a fórmula resultar em um erro.

#### **Exemplo:**

```excel
=SEERRO(A1/B1, "Divisão por zero")
```

- Se B1 for 0, a fórmula evita o erro `#DIV/0!` e retorna "Divisão por zero".

### **3. Função `SOMASE` (SUMIF)**

#### **Descrição:**

A função `SOMASE` (em inglês, `SUMIF`) soma os valores em um intervalo que atendem a uma condição específica.

#### **Sintaxe:**

```excel
=SOMASE(intervalo, critério, [intervalo_soma])
```

- **`intervalo`:** O intervalo de células que você quer avaliar.
- **`critério`:** A condição que deve ser atendida para que os valores sejam somados.
- **`intervalo_soma` (opcional):** O intervalo de células cujos valores serão somados se a condição for atendida.

#### **Exemplo:**

```excel
=SOMASE(A1:A10, ">10")
```

- Soma todos os valores em A1:A10 que são maiores que 10.

### **4. Função `CONT.SE` (COUNTIF)**

#### **Descrição:**

A função `CONT.SE` (em inglês, `COUNTIF`) conta o número de células em um intervalo que atendem a uma condição específica.

#### **Sintaxe:**

```excel
=CONT.SE(intervalo, critério)
```

- **`intervalo`:** O intervalo de células que você quer avaliar.
- **`critério`:** A condição que as células devem atender para serem contadas.

#### **Exemplo:**

```excel
=CONT.SE(B1:B10, "Aprovado")
```

- Conta quantas vezes "Aprovado" aparece no intervalo B1:B10.

### **5. Função `SE` Aninhado (Nested IF)**

#### **Descrição:**

Você pode aninhar várias funções `SE` dentro de uma fórmula para lidar com múltiplas condições.

#### **Sintaxe:**

```excel
=SE(teste_lógico1, valor_se_verdadeiro1, SE(teste_lógico2, valor_se_verdadeiro2, valor_se_falso))
```

#### **Exemplo:**

```excel
=SE(A1 > 90, "Excelente", SE(A1 > 75, "Bom", "Regular"))
```

- Se A1 for maior que 90, retorna "Excelente". Se for maior que 75 mas menor ou igual a 90, retorna "Bom". Caso contrário, retorna "Regular".

### **6. Função `MÍNIMOSE` e `MÁXIMOSE` (MINIFS e MAXIFS)**

#### **Descrição:**

As funções `MÍNIMOSE` e `MÁXIMOSE` (em inglês, `MINIFS` e `MAXIFS`) retornam o menor ou maior valor em um intervalo que atende a uma ou mais condições.

#### **Sintaxe:**

```excel
=MÍNIMOSE(intervalo_mínimo, intervalo_critério1, critério1, [intervalo_critério2, critério2, ...])
=MÁXIMOSE(intervalo_máximo, intervalo_critério1, critério1, [intervalo_critério2, critério2, ...])
```

#### **Exemplo:**

```excel
=MÍNIMOSE(B2:B10, A2:A10, ">=2023")
```

- Retorna o menor valor no intervalo B2:B10 onde o valor correspondente em A2:A10 é maior ou igual a 2023.

```excel
=MÁXIMOSE(C2:C10, A2:A10, "São Paulo", B2:B10, ">100")
```

- Retorna o maior valor em C2:C10 onde a célula correspondente em A2:A10 é "São Paulo" e o valor em B2:B10 é maior que 100.

### **7. Função `PROCV` e `ÍNDICE/CORRESP` (VLOOKUP e INDEX/MATCH)**

#### **Descrição:**

Embora não sejam estritamente condicionais, essas funções são frequentemente usadas para criar condições complexas, como buscar valores em uma tabela com base em critérios.

#### **Exemplo PROCV (VLOOKUP):**

```excel
=PROCV(valor_procurado, tabela, número_da_coluna, [procurar_intervalo])
```

- Exemplo: `=PROCV("Produto A", A2:C10, 3, FALSO)` procura por "Produto A" na primeira coluna do intervalo A2:C10 e retorna o valor correspondente da terceira coluna.

#### **Exemplo ÍNDICE/CORRESP (INDEX/MATCH):**

```excel
=ÍNDICE(intervalo_resultado, CORRESP(valor_procurado, intervalo_procurado, 0))
```

- Exemplo: `=ÍNDICE(C2:C10, CORRESP("Produto A", A2:A10, 0))` retorna o valor da coluna C correspondente ao "Produto A" na coluna A.

### **Conclusão**

As funções condicionais são ferramentas poderosas para analisar e manipular dados com base em condições específicas. Elas ajudam a automatizar processos, tornar análises mais precisas e facilitar a tomada de decisões em planilhas. Dependendo da complexidade dos dados, você pode combinar várias funções para criar soluções sofisticadas e dinâmicas.



## Formatação Condicional

A formatação condicional em Excel e Google Sheets é uma ferramenta que permite aplicar formatações específicas (como cores de células, estilos de texto, ícones, etc.) automaticamente a células ou intervalos de células com base em determinadas condições. Isso ajuda a destacar informações importantes, identificar padrões e facilitar a análise visual dos dados.

### **1. Como Usar a Formatação Condicional em Excel**

#### **Passos Básicos para Aplicar a Formatação Condicional:**

1. **Selecione as Células:**
   - Selecione as células ou o intervalo de células em que deseja aplicar a formatação condicional.

2. **Acesse a Ferramenta de Formatação Condicional:**
   - Na aba **"Página Inicial"** do Excel, clique em **"Formatação Condicional"** no grupo **"Estilos"**.

3. **Escolha uma Regra de Formatação:**
   - O Excel oferece várias opções de regras predefinidas:
     - **Realce de Células:** Para destacar células que contêm valores específicos, como valores maiores que, menores que, entre outros.
     - **Regras de Primeiro/Último:** Para destacar os maiores ou menores valores de um intervalo.
     - **Escalas de Cor:** Aplica uma gradação de cores baseada nos valores das células.
     - **Barras de Dados:** Adiciona barras de dados dentro das células para mostrar visualmente a magnitude dos valores.
     - **Conjuntos de Ícones:** Adiciona ícones como setas, sinais de alerta, etc., para categorizar os valores.

4. **Configurar a Regra:**
   - Depois de escolher uma regra, você poderá ajustar os critérios da formatação, como o valor de comparação, o estilo de formatação (cor, tipo de ícone, etc.), e o intervalo de valores.

5. **Aplicar a Regra:**
   - Clique em **"OK"** para aplicar a formatação condicional. O Excel ajustará automaticamente a formatação das células que atendem aos critérios estabelecidos.

#### **Exemplo:**

- **Destacar valores maiores que 100:**
  1. Selecione o intervalo (por exemplo, A1:A10).
  2. Vá em **"Formatação Condicional"** > **"Realçar Regras de Células"** > **"Maior que..."**.
  3. Digite `100` e escolha uma cor de destaque.
  4. Clique em **"OK"**.

### **2. Como Usar a Formatação Condicional em Google Sheets**

#### **Passos Básicos para Aplicar a Formatação Condicional:**

1. **Selecione as Células:**
   - Selecione as células ou intervalo de células em que deseja aplicar a formatação condicional.

2. **Acesse a Ferramenta de Formatação Condicional:**
   - No menu superior, clique em **"Formatar"** e selecione **"Formatação condicional"**.

3. **Configurar as Regras de Formatação:**
   - Na barra lateral que aparecerá, você pode escolher:
     - **Intervalo de células:** Verifique se o intervalo selecionado está correto.
     - **Regras de Formatação:** Escolha a regra que deseja aplicar, como:
       - Valores maiores que, menores que, entre, etc.
       - Texto que contém, começa com, termina com, etc.
       - Data, texto vazio ou não vazio, etc.
     - **Formatação personalizada:** Caso as regras predefinidas não atendam às suas necessidades, você pode criar fórmulas personalizadas para definir as condições.

4. **Escolher o Estilo de Formatação:**
   - Defina como a formatação será aplicada, escolhendo a cor da célula, cor do texto, estilo da fonte, etc.

5. **Aplicar a Formatação:**
   - Após configurar, clique em **"Concluído"** para aplicar a formatação condicional.

#### **Exemplo:**

- **Destacar células com texto específico:**
  1. Selecione o intervalo (por exemplo, B1:B20).
  2. Vá em **"Formatar"** > **"Formatação condicional"**.
  3. Em "Regras de formatação", escolha **"O texto contém"** e digite o texto (por exemplo, "Aprovado").
  4. Escolha uma cor de preenchimento e clique em **"Concluído"**.

### **3. Regras Personalizadas com Fórmulas**

#### **Excel:**

Você pode criar regras de formatação condicional usando fórmulas personalizadas. Isso é útil quando as regras predefinidas não são suficientes.

- **Exemplo:**
  Se você quiser destacar células em uma coluna se o valor em uma célula adjacente (na mesma linha) for maior que 100:
  1. Selecione o intervalo (por exemplo, A1:A10).
  2. Vá em **"Formatação Condicional"** > **"Nova Regra..."** > **"Usar uma fórmula para determinar quais células devem ser formatadas"**.
  3. Digite a fórmula: `=B1>100`.
  4. Escolha a formatação desejada e clique em **"OK"**.

#### **Google Sheets:**

Também permite o uso de fórmulas para criar regras personalizadas.

- **Exemplo:**
  Para destacar células na coluna A se o valor na coluna B da mesma linha for "Sim":
  1. Selecione o intervalo (por exemplo, A1:A10).
  2. Vá em **"Formatar"** > **"Formatação condicional"**.
  3. Escolha "A fórmula personalizada é" e digite: `=B1="Sim"`.
  4. Escolha a formatação desejada e clique em **"Concluído"**.

### **4. Gerenciamento de Regras de Formatação Condicional**

- **Editar/Remover Regras:**
  - Em Excel, você pode gerenciar todas as regras de formatação condicional em **"Formatação Condicional"** > **"Gerenciar Regras..."**.
  - No Google Sheets, você pode ver e editar as regras na barra lateral da formatação condicional.

- **Prioridade das Regras:**
  - Em Excel, se várias regras de formatação condicional se aplicarem a uma mesma célula, a ordem das regras na lista determinará qual será aplicada primeiro.

### **Conclusão**

A formatação condicional é uma ferramenta extremamente poderosa para destacar informações relevantes em grandes conjuntos de dados. Seja para identificar rapidamente valores extremos, destacar tendências, ou simplesmente organizar visualmente seus dados, o uso adequado de formatação condicional pode transformar a maneira como você interage com suas planilhas.

## **Operadores lógicos**

Operadores lógicos são fundamentais em Excel e Google Sheets para realizar comparações e tomar decisões baseadas em condições específicas. Eles são frequentemente usados em conjunto com funções como `SE` (IF), `E` (AND), `OU` (OR), e outras, para criar fórmulas mais complexas e dinâmicas. Vamos explorar os principais operadores lógicos e como usá-los.

### **1. Operadores Lógicos Básicos**

#### **a. Igualdade (`=`)**

- **Descrição:** Verifica se dois valores são iguais.
- **Exemplo:**
  - `=A1=B1` retorna VERDADEIRO se os valores em A1 e B1 forem iguais.

#### **b. Diferente (`<>`)**

- **Descrição:** Verifica se dois valores são diferentes.
- **Exemplo:**
  - `=A1<>B1` retorna VERDADEIRO se os valores em A1 e B1 forem diferentes.

#### **c. Maior que (`>`)**

- **Descrição:** Verifica se um valor é maior que outro.
- **Exemplo:**
  - `=A1>B1` retorna VERDADEIRO se o valor em A1 for maior que o valor em B1.

#### **d. Menor que (`<`)**

- **Descrição:** Verifica se um valor é menor que outro.
- **Exemplo:**
  - `=A1<B1` retorna VERDADEIRO se o valor em A1 for menor que o valor em B1.

#### **e. Maior ou Igual a (`>=`)**

- **Descrição:** Verifica se um valor é maior ou igual a outro.
- **Exemplo:**
  - `=A1>=B1` retorna VERDADEIRO se o valor em A1 for maior ou igual ao valor em B1.

#### **f. Menor ou Igual a (`<=`)**

- **Descrição:** Verifica se um valor é menor ou igual a outro.
- **Exemplo:**
  - `=A1<=B1` retorna VERDADEIRO se o valor em A1 for menor ou igual ao valor em B1.

### **2. Operadores Lógicos Combinados**

#### **a. `E` (AND)**

- **Descrição:** Retorna VERDADEIRO se **todas** as condições especificadas forem verdadeiras.

- **Sintaxe:**

  ```excel
  =E(teste_lógico1, teste_lógico2, ...)
  ```

- **Exemplo:**

  - `=E(A1>10, B1<20)` retorna VERDADEIRO se A1 for maior que 10 e B1 for menor que 20.

#### **b. `OU` (OR)**

- **Descrição:** Retorna VERDADEIRO se **pelo menos uma** das condições especificadas for verdadeira.

- **Sintaxe:**

  ```excel
  =OU(teste_lógico1, teste_lógico2, ...)
  ```

- **Exemplo:**

  - `=OU(A1>10, B1<20)` retorna VERDADEIRO se A1 for maior que 10 **ou** B1 for menor que 20.

#### **c. `NÃO` (NOT)**

- **Descrição:** Inverte o resultado de um teste lógico. Se a condição for verdadeira, `NÃO` a torna falsa, e vice-versa.

- **Sintaxe:**

  ```excel
  =NÃO(teste_lógico)
  ```

- **Exemplo:**

  - `=NÃO(A1>10)` retorna VERDADEIRO se A1 **não** for maior que 10.

### **3. Usando Operadores Lógicos em Funções**

Operadores lógicos são frequentemente usados em funções como `SE` (IF) para criar condições mais sofisticadas.

#### **Exemplo 1: Usando `E` com `SE`**

- **Descrição:** Verifique se um aluno passou em duas matérias.

- **Fórmula:**

  ```excel
  =SE(E(A1>=50, B1>=50), "Passou", "Reprovou")
  ```

  - Retorna "Passou" se o aluno tirou 50 ou mais em ambas as matérias.

#### **Exemplo 2: Usando `OU` com `SE`**

- **Descrição:** Verifique se um produto está fora dos limites de estoque.

- **Fórmula:**

  ```excel
  =SE(OU(A1<10, A1>100), "Fora do Estoque Ideal", "Estoque Normal")
  ```

  - Retorna "Fora do Estoque Ideal" se o estoque estiver abaixo de 10 ou acima de 100.

#### **Exemplo 3: Usando `NÃO` com `SE`**

- **Descrição:** Verifique se um item não está disponível.

- **Fórmula:**

  ```excel
  =SE(NÃO(A1="Disponível"), "Indisponível", "Disponível")
  ```

  - Retorna "Indisponível" se A1 não for igual a "Disponível".

### **4. Combinações de Operadores Lógicos**

Você pode combinar múltiplos operadores lógicos para criar fórmulas complexas.

#### **Exemplo:**

- **Descrição:** Verifique se um número está dentro de um intervalo específico ou se é negativo.

- **Fórmula:**

  ```excel
  =SE(OU(E(A1>=10, A1<=20), A1<0), "Condição Satisfeita", "Condição Não Satisfeita")
  ```

  - Retorna "Condição Satisfeita" se A1 estiver entre 10 e 20 (inclusive) ou se for negativo.

### **5. Aplicação Prática**

#### **Exemplo: Tabela de Desempenho de Funcionários**

- **Objetivo:** Classificar o desempenho dos funcionários com base em suas avaliações.

- **Condições:**

  - Se a avaliação em Competência (`B1`) e Dedicação (`C1`) forem ambas acima de 8, classificar como "Excelente".
  - Se qualquer uma das avaliações for 6 ou menos, classificar como "A Melhorar".
  - Caso contrário, classificar como "Bom".

- **Fórmula:**

  ```excel
  =SE(E(B1>8, C1>8), "Excelente", SE(OU(B1<=6, C1<=6), "A Melhorar", "Bom"))
  ```

### **Conclusão**

Os operadores lógicos são ferramentas essenciais em Excel e Google Sheets para criar fórmulas complexas e dinâmicas, permitindo análises mais detalhadas e tomadas de decisões mais informadas. Ao combinar esses operadores com funções condicionais, você pode automatizar muitos processos e aumentar significativamente a eficiência e precisão do seu trabalho em planilhas.

## Função SE

A função `SE` (em inglês, `IF`) é uma das mais fundamentais e versáteis em Excel e Google Sheets. Ela permite que você faça cálculos e tome decisões com base em condições específicas. A função `SE` verifica se uma condição é verdadeira ou falsa e, dependendo do resultado, retorna um valor ou executa uma ação.

### **Sintaxe da Função `SE`**

```excel
=SE(teste_lógico, valor_se_verdadeiro, valor_se_falso)
```

- **`teste_lógico`:** A condição que você deseja avaliar. Pode ser uma comparação entre valores, como `A1 > 10`, `B1 = "Sim"`, etc.
- **`valor_se_verdadeiro`:** O valor ou ação que a função retorna/executa se a condição for verdadeira.
- **`valor_se_falso`:** O valor ou ação que a função retorna/executa se a condição for falsa.

### **Exemplos Simples de Uso da Função `SE`**

#### **1. Avaliar se um Número é Maior que Outro**

- **Exemplo:**

  ```excel
  =SE(A1 > 10, "Maior que 10", "Menor ou igual a 10")
  ```

  - **Explicação:** Se o valor em A1 for maior que 10, a função retorna "Maior que 10". Caso contrário, retorna "Menor ou igual a 10".

#### **2. Verificar a Aprovação de um Aluno**

- **Exemplo:**

  ```excel
  =SE(B1 >= 70, "Aprovado", "Reprovado")
  ```

  - **Explicação:** Se a nota em B1 for maior ou igual a 70, retorna "Aprovado". Caso contrário, retorna "Reprovado".

#### **3. Comparar Texto**

- **Exemplo:**

  ```excel
  =SE(C1 = "Sim", "Continuar", "Parar")
  ```

  - **Explicação:** Se o texto em C1 for "Sim", retorna "Continuar". Caso contrário, retorna "Parar".

### **Função `SE` Aninhada (Nested IF)**

Você pode usar várias funções `SE` dentro de uma fórmula para lidar com múltiplas condições. Isso é chamado de `SE` aninhado.

#### **Exemplo de Função `SE` Aninhada:**

- **Descrição:** Classificar um aluno com base em sua nota.

  - Se a nota for 90 ou mais, o aluno é "Excelente".
  - Se a nota for 75 ou mais, o aluno é "Bom".
  - Caso contrário, o aluno é "Regular".

- **Fórmula:**

  ```excel
  =SE(A1 >= 90, "Excelente", SE(A1 >= 75, "Bom", "Regular"))
  ```

### **Combinação de `SE` com Outras Funções**

#### **1. Usando `SE` com `E` (AND)**

- **Exemplo:**

  ```excel
  =SE(E(A1 >= 70, B1 >= 70), "Aprovado", "Reprovado")
  ```

  - **Explicação:** A função retorna "Aprovado" se ambas as condições forem verdadeiras (A1 e B1 maiores ou iguais a 70). Caso contrário, retorna "Reprovado".

#### **2. Usando `SE` com `OU` (OR)**

- **Exemplo:**

  ```excel
  =SE(OU(A1 >= 70, B1 >= 70), "Aprovado", "Reprovado")
  ```

  - **Explicação:** A função retorna "Aprovado" se pelo menos uma das condições for verdadeira (A1 ou B1 maior ou igual a 70). Caso contrário, retorna "Reprovado".

### **Exemplos Práticos de Uso da Função `SE`**

#### **1. Calculando Descontos**

- **Descrição:** Se o valor de uma compra (em B1) for maior que 1000, aplicar um desconto de 10%. Caso contrário, aplicar um desconto de 5%.

- **Fórmula:**

  ```excel
  =SE(B1 > 1000, B1 * 0.9, B1 * 0.95)
  ```

#### **2. Determinando uma Categoria com Base em Idade**

- **Descrição:** Categorizar uma pessoa como "Criança", "Adolescente" ou "Adulto" com base na idade.

- **Fórmula:**

  ```excel
  =SE(A1 < 13, "Criança", SE(A1 < 18, "Adolescente", "Adulto"))
  ```

#### **3. Verificando Disponibilidade de Estoque**

- **Descrição:** Se o estoque de um produto (em A1) for menor que 10, exibir "Repor Estoque". Caso contrário, exibir "Estoque Suficiente".

- **Fórmula:**

  ```excel
  =SE(A1 < 10, "Repor Estoque", "Estoque Suficiente")
  ```

### **Limitações e Considerações**

- **Complexidade:** Fórmulas `SE` aninhadas podem se tornar complexas e difíceis de ler quando muitas condições são envolvidas.
- **Erro ao Aninhar Demais:** Excesso de `SE` aninhados pode resultar em erros e dificultar a manutenção da planilha.
- **Alternativas:** Em cenários muito complexos, considere o uso de outras funções ou combinações, como `PROC.V` (VLOOKUP), `ÍNDICE/CORRESP` (INDEX/MATCH), ou até tabelas de referência para simplificar a lógica.

### **Conclusão**

A função `SE` é uma ferramenta poderosa e essencial para quem trabalha com planilhas. Ela permite tomar decisões e realizar cálculos baseados em condições específicas, tornando suas planilhas dinâmicas e interativas. Combinada com outras funções, a `SE` pode resolver uma ampla gama de problemas e simplificar muitos processos analíticos.

## Funções de contagem

Funções de contagem em Excel e Google Sheets são ferramentas poderosas para contar células, valores, caracteres, ou até mesmo resultados específicos que atendem a certos critérios. Essas funções são muito úteis para análise de dados, oferecendo uma maneira rápida e eficiente de quantificar informações em uma planilha.

### **1. Função `CONT.NÚM` (COUNT)**

#### **Descrição:**

A função `CONT.NÚM` (em inglês, `COUNT`) conta quantas células em um intervalo contêm números.

#### **Sintaxe:**

```excel
=CONT.NÚM(intervalo)
```

- **`intervalo`:** O intervalo de células que você deseja contar.

#### **Exemplo:**

```excel
=CONT.NÚM(A1:A10)
```

- **Explicação:** Conta quantas células em A1:A10 contêm números.

### **2. Função `CONT.VALORES` (COUNTA)**

#### **Descrição:**

A função `CONT.VALORES` (em inglês, `COUNTA`) conta quantas células em um intervalo não estão vazias. Isso inclui números, texto, datas, etc.

#### **Sintaxe:**

```excel
=CONT.VALORES(intervalo)
```

- **`intervalo`:** O intervalo de células que você deseja contar.

#### **Exemplo:**

```excel
=CONT.VALORES(B1:B10)
```

- **Explicação:** Conta quantas células em B1:B10 contêm qualquer tipo de valor (número, texto, etc.).

### **3. Função `CONTAR.VAZIO` (COUNTBLANK)**

#### **Descrição:**

A função `CONTAR.VAZIO` (em inglês, `COUNTBLANK`) conta quantas células em um intervalo estão vazias.

#### **Sintaxe:**

```excel
=CONTAR.VAZIO(intervalo)
```

- **`intervalo`:** O intervalo de células que você deseja contar.

#### **Exemplo:**

```excel
=CONTAR.VAZIO(C1:C10)
```

- **Explicação:** Conta quantas células em C1:C10 estão vazias.

### **4. Função `CONT.SE` (COUNTIF)**

#### **Descrição:**

A função `CONT.SE` (em inglês, `COUNTIF`) conta quantas células em um intervalo atendem a um critério específico.

#### **Sintaxe:**

```excel
=CONT.SE(intervalo, critério)
```

- **`intervalo`:** O intervalo de células que você deseja contar.
- **`critério`:** A condição que deve ser atendida para que uma célula seja contada.

#### **Exemplo:**

```excel
=CONT.SE(D1:D10, ">100")
```

- **Explicação:** Conta quantas células em D1:D10 têm valores maiores que 100.

#### **Exemplo 2:**

```excel
=CONT.SE(E1:E10, "Aprovado")
```

- **Explicação:** Conta quantas células em E1:E10 contêm o texto "Aprovado".

### **5. Função `CONT.SES` (COUNTIFS)**

#### **Descrição:**

A função `CONT.SES` (em inglês, `COUNTIFS`) conta quantas células atendem a vários critérios em diferentes intervalos.

#### **Sintaxe:**

```excel
=CONT.SES(intervalo1, critério1, [intervalo2, critério2], ...)
```

- **`intervalo1`, `critério1`:** O primeiro intervalo e critério que devem ser atendidos.
- **`intervalo2`, `critério2`:** (Opcional) Adicional intervalo e critério que devem ser atendidos.

#### **Exemplo:**

```excel
=CONT.SES(F1:F10, ">50", G1:G10, "<=100")
```

- **Explicação:** Conta quantas células em F1:F10 têm valores maiores que 50 e, simultaneamente, quantas células em G1:G10 têm valores menores ou iguais a 100.

### **6. Função `FREQÜÊNCIA` (FREQUENCY)**

#### **Descrição:**

A função `FREQÜÊNCIA` (em inglês, `FREQUENCY`) conta quantos valores em um intervalo caem dentro de intervalos específicos (bins). É muito útil para análise estatística e criação de histogramas.

#### **Sintaxe:**

```excel
=FREQÜÊNCIA(intervalo_dados, intervalo_bins)
```

- **`intervalo_dados`:** O intervalo de dados que você deseja analisar.
- **`intervalo_bins`:** Os intervalos que você deseja usar para agrupar os dados.

#### **Exemplo:**

```excel
=FREQÜÊNCIA(A1:A10, B1:B4)
```

- **Explicação:** Conta quantos valores em A1:A10 caem nos intervalos definidos em B1:B4. Esta função retorna uma matriz de valores que correspondem à contagem de cada intervalo.

### **7. Função `SOMARPRODUTO` (SUMPRODUCT)** (Como Contagem Avançada)

#### **Descrição:**

Embora não seja uma função de contagem direta, `SOMARPRODUTO` (em inglês, `SUMPRODUCT`) pode ser usada para contar quantas vezes certas condições são atendidas multiplicando matrizes lógicas.

#### **Sintaxe:**

```excel
=SOMARPRODUTO((intervalo1=critério1)*(intervalo2=critério2))
```

- **`intervalo1`, `critério1`:** Primeiro intervalo e critério.
- **`intervalo2`, `critério2`:** Segundo intervalo e critério.

#### **Exemplo:**

```excel
=SOMARPRODUTO((A1:A10="Sim")*(B1:B10>100))
```

- **Explicação:** Conta quantas vezes "Sim" aparece em A1:A10 e, simultaneamente, o valor correspondente em B1:B10 é maior que 100.

### **8. Função `NÚM.CARACT` (LEN)**

clo@1closet.ai

23h4n10l80!!

#### **Descrição:**

A função `NÚM.CARACT` (em inglês, `LEN`) conta o número de caracteres em uma célula, incluindo espaços.

#### **Sintaxe:**

```excel
=NÚM.CARACT(texto)
```

- **`texto`:** A célula ou texto que você deseja contar os caracteres.

#### **Exemplo:**

```excel
=NÚM.CARACT(A1)
```

- **Explicação:** Conta o número de caracteres na célula A1.

### **Conclusão**

As funções de contagem são essenciais para analisar e interpretar dados em Excel e Google Sheets. Elas ajudam a quantificar diferentes aspectos de seus dados, permitindo que você faça análises detalhadas e tome decisões informadas. Seja para contar células com números, células vazias, ou até mesmo para aplicar critérios específicos, essas funções oferecem a flexibilidade necessária para uma variedade de aplicações.

