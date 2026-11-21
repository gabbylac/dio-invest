# DIO INVEST – Ferramenta de Controle de Investimentos

## 1. Sobre a ferramenta

A DIO INVEST é uma ferramenta desenvolvida no Excel para auxiliar no planejamento e na organização de investimentos, permitindo simular o crescimento do patrimônio ao longo do tempo e sugerir uma distribuição mensal de investimentos em Fundos Imobiliários (FIIs), de acordo com o perfil do investidor.

A ferramenta utiliza fórmulas financeiras, funções de busca e intervalos nomeados para automatizar os cálculos.

## 2. Quais perguntas a ferramenta responde?

A planilha está organizada em duas partes principais:

### Configurações e planejamento financeiro – Planilha1

A ferramenta permite responder às seguintes perguntas:

- **Qual é o meu salário?** Informa a renda mensal do investidor.
- **Qual é o rendimento da carteira?** Define o percentual de rendimento mensal utilizado para estimar os dividendos.
- **Quanto posso investir do meu salário?** Calcula uma sugestão de investimento equivalente a 30% do salário.
- **Quanto investir por mês?** Permite informar o valor do aporte mensal.
- **Por quantos anos pretendo investir?** Define o período do investimento.
- **Qual é a taxa de rendimento mensal?** Informa a taxa utilizada na simulação financeira.
- **Qual será meu patrimônio acumulado?** Calcula o valor estimado acumulado ao final do período escolhido.
- **Quanto poderei receber de dividendos mensais?** Estima os dividendos com base no patrimônio acumulado e no rendimento da carteira.

### Simulação de cenários – Planilha1

A ferramenta também apresenta simulações de crescimento do patrimônio para diferentes períodos:

- 2 anos;
- 5 anos;
- 10 anos;
- 20 anos;
- 30 anos.

Os resultados aparecem nas colunas de patrimônio acumulado e dividendos estimados.

### Distribuição de investimentos em FIIs – Planilha1 e Planilha2

A ferramenta permite selecionar um perfil de investidor:

- Conservador;
- Moderado;
- Agressivo.

Com base nessa seleção, apresenta o percentual sugerido para cada categoria de FII e calcula quanto investir em cada uma delas, considerando o valor mensal informado.

As categorias utilizadas são:
- Papel;
- Tijolo;
- Híbridos;
- FOFs;
- Desenvolvimento;
- Hotelarias.

Os percentuais e valores sugeridos aparecem na tabela de distribuição de investimentos da Planilha1.

## 3. Como o VF e o PROCV entram nos cálculos?

### Função VF (Valor Futuro)

A função VF é utilizada para calcular quanto o dinheiro investido poderá acumular ao longo do tempo, considerando uma taxa de rendimento e aportes periódicos.

Na ferramenta, ela aparece no cálculo do patrimônio acumulado e nas simulações de cenários.

Exemplo da fórmula utilizada:

`=VF(taxa_mensal;Qtd_anos*12;Aporte*-1)`

A fórmula considera:
- Taxa mensal de rendimento;
- Quantidade de anos multiplicada por 12, para converter em meses;
- Valor do aporte mensal, utilizando o sinal negativo para representar a saída de dinheiro.

Dessa forma, é possível estimar o patrimônio acumulado em diferentes períodos.

### Função PROCV

A função PROCV é utilizada para buscar automaticamente os percentuais de distribuição dos investimentos na Planilha2.

A busca acontece por meio de uma chave formada pela combinação do perfil do investidor com o tipo de FII.

Exemplo de chave:

`Conservador.Tijolo`

A fórmula utilizada é:

`=PROCV($C$32&B36;Planilha2!$A:$D;4;)`

Ela procura a combinação do perfil com a categoria do fundo na primeira coluna da tabela da Planilha2 e retorna o percentual correspondente, localizado na quarta coluna.

Depois, esse percentual é multiplicado pelo valor mensal disponível para investimento, apresentando o valor sugerido para cada categoria.

## 4. Intervalos nomeados utilizados

Foram criados intervalos nomeados para facilitar a leitura das fórmulas e tornar a planilha mais organizada.

| Nome do intervalo | Célula de referência | Finalidade |
|---|---|---|
| Salario | Planilha1!D12 | Armazena o salário mensal. |
| Rendimento_carteira | Planilha1!D13 | Armazena o rendimento mensal da carteira. |
| Aporte | Planilha1!D17 | Armazena o valor do investimento mensal. |
| Qtd_anos | Planilha1!D18 | Define o período de investimento em anos. |
| taxa_mensal | Planilha1!D19 | Armazena a taxa de rendimento mensal. |
| Patrimonio | Planilha1!D20 | Representa o patrimônio acumulado. |

O uso dos intervalos nomeados facilita a identificação das informações utilizadas nas fórmulas e permite uma manutenção mais simples da ferramenta.

## 5. Percentuais de cada perfil de investidor

Os percentuais estão organizados na Planilha2, de acordo com os três perfis de investimento.

### Perfil Conservador

| Tipo de FII | Percentual |
|---|---:|
| Papel | 30% |
| Tijolo | 50% |
| Híbridos | 10% |
| FOFs | 10% |
| Desenvolvimento | 0% |
| Hotelarias | 0% |
| **Total** | **100%** |

### Perfil Moderado

| Tipo de FII | Percentual |
|---|---:|
| Papel | 32% |
| Tijolo | 35% |
| Híbridos | 8% |
| FOFs | 5% |
| Desenvolvimento | 10% |
| Hotelarias | 10% |
| **Total** | **100%** |

### Perfil Agressivo

| Tipo de FII | Percentual |
|---|---:|
| Papel | 50% |
| Tijolo | 10% |
| Híbridos | 5% |
| FOFs | 5% |
| Desenvolvimento | 20% |
| Hotelarias | 10% |
| **Total** | **100%** |

### De onde vieram os percentuais?

Os percentuais foram organizados na tabela de referência da Planilha2 e utilizados como parâmetros para a distribuição dos investimentos de cada perfil.

Eles representam os critérios de distribuição adotados na construção da ferramenta. O arquivo não apresenta uma fonte externa ou referência técnica que comprove a origem desses percentuais.

Por isso, devem ser entendidos como parâmetros de simulação da planilha, e não como recomendações financeiras personalizadas.

## 6. O que foi modificado em relação à ferramenta do Expert?

A ferramenta DIO INVEST foi desenvolvida a partir do modelo apresentado pelo Expert, mantendo sua estrutura principal de planejamento financeiro, simulação de patrimônio e distribuição de investimentos em Fundos Imobiliários.

Foram realizadas algumas personalizações e alterações:

- **Personalização das fórmulas:** substituição de referências diretas de células por intervalos nomeados, como na fórmula de sugestão de investimento, que passou de `=D12*30%` para `=Salario*30%`.
- **Alteração da chave de pesquisa:** substituição do hífen pelo ponto como separador entre o perfil do investidor e o tipo de FII nas chaves utilizadas pelo PROCV.
- **Perfil inicial:** alteração do perfil selecionado inicialmente de Moderado para Conservador.
- **Personalização visual e textual:** ajustes nos nomes das abas, títulos e descrições dos campos.
- **Personalização das referências:** alteração da nomenclatura dos intervalos nomeados, mantendo suas respectivas referências de células.
- **Distribuição dos investimentos:** manutenção das seis categorias de FIIs e dos percentuais originais de cada perfil.

Os percentuais de distribuição foram mantidos conforme o modelo original do Expert, sem alterações nos critérios de divisão entre os perfis Conservador, Moderado e Agressivo.

Dessa forma, a versão desenvolvida preserva a lógica principal da ferramenta original, mas apresenta personalizações nas fórmulas, referências e organização das informações.
## 7. Considerações finais

O desenvolvimento da DIO INVEST permitiu aplicar conhecimentos de Excel, fórmulas financeiras, funções de busca, referências e intervalos nomeados.

A ferramenta reúne cálculos e simulações em uma estrutura automatizada, facilitando a visualização de possíveis cenários de investimento e a organização dos aportes mensais.

O projeto demonstra como o Excel pode ser utilizado para desenvolver soluções práticas de controle e planejamento financeiro.

**Importante:** Os resultados são estimativas matemáticas e não representam garantia de rentabilidade ou recomendação de investimento.
