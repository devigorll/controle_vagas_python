 # ControleVagas-Python 🚗📊

 Sistema automatizado de auditoria e inteligência operacional para controle de ocupação e detecção de excedentes em vagas de estacionamento.

 ## 📌 Sobre o Projeto
 O **ControleVagas-Python** foi desenvolvido para resolver o desafio de fiscalização e auditoria em estacionamentos com cotas limite de vagas por sala ou conjunto comercial.

 A ferramenta processa planilhas de acessos (`.xlsx` ou `.csv`), cruza os horários de entrada e saída em tempo real e identifica automaticamente veículos que ultrapassaram a capacidade simultânea permitida, gerando relatórios consolidados e relatórios visuais interativos.

 ## ⚙️ Principais Funcionalidades
 * **Auditoria Simultânea de Vagas:** Cruzamento cronológico de entradas e saídas para verificar sobreposição de horários e estouro de capacidade.
 * **Suporte a Salas Múltiplas:** Cálculo inteligente para conjuntos unificados (ex: salas `61/62`), somando dinamicamente os limites de cada espaço.
 * **Formatação Monetária PT-BR:** Aplicação automática do padrão `R$ VALOR,00` em colunas financeiras (valores, taxas, cobranças).
 * **Relatório Detalhado de Infrações:** Exportação de dados consolidados em `.xlsx` e `.csv` detalhando a justificativa e o horário de cada veículo excedente.
 * **Dashboard Interativo em Streamlit:** Visualização gráfica dos principais KPIs de utilização, volume por dia da semana e ranking de salas com maior recorrência de excesso.

 ## 🛠️ Tecnologias Utilizadas
 * **Linguagem:** Python 3.10+
 * **Interface Web:** Streamlit
 * **Manipulação de Dados:** Pandas, NumPy
 * **Visualização de Dados:** Plotly Express
 * **Exportação de Arquivos:** OpenPyXL

 ## 📂 Estrutura do Repositório
 `text ll ├── data/                   # Diretório reservado para leitura/carga de arquivos ll ├── app.py                  # Aplicação principal Streamlit e lógica de auditoria ll ├── requirements.txt        # Dependências do projeto ll └── README.md               # Documentação do projeto ll `

 ## 🚀 Como Executar o Projeto

 1. **Clone o repositório:**
 `bash ll git clone https://github.com/seu-usuario/ControleVagas-Python.git ll cd ControleVagas-Python ll `

 2. **Crie e ative um ambiente virtual (opcional, mas recomendado):**
 `bash ll python -m venv venv ll # No Windows: ll venv\Scripts\activate ll # No Linux/Mac: ll source venv/bin/activate ll `

 3. **Instale as dependências:**
 `bash ll pip install -r requirements.txt ll `

 4. **Execute a aplicação:**
 `bash ll streamlit run app.py ll `

 5. **Acesse no navegador:**
 O Streamlit abrirá automaticamente no endereço `http://localhost:8501`.

 ## 📊 Demonstração dos Resultados
 O sistema gera dois relatórios principais na interface:
 * **Relatório de Excesso:** Tabela detalhada indicando a sala, quantidade de vagas autorizadas, horários de sobreposição dos veículos e o motivo explícito da infração.
 * **Dashboard & KPIs:** Painel executivo exibindo a taxa percentual de infração, tempo médio de permanência, volume diário de acessos e gráfico com o Top 7 de salas com mais estouros de capacidade.

 ## 📄 Licença
 Este projeto está sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.