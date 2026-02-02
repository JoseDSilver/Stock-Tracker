# Stock-Tracker
Descrição
Aplicação em Python que consulta dados de ações da B3 utilizando a API pública do Yahoo Finance por meio da biblioteca yfinance e atualiza automaticamente uma planilha Excel com informações financeiras relevantes.

O programa possui uma interface gráfica simples desenvolvida com CustomTkinter, permitindo que o usuário insira múltiplos tickers e atualize os dados diretamente na planilha.

Uma planilha de exemplo é disponibilizada neste repositório. O usuário deve ajustar manualmente o caminho da planilha no código para o local correto em sua máquina.

---

Funcionalidades
- Consulta de dados de ações da B3 via Yahoo Finance (yfinance)
- Atualização automática de planilha Excel
- Cálculo da diferença entre cotação atual e preço teto definido na planilha
- Destaque visual para valores positivos ou negativos
- Suporte à consulta de múltiplos tickers simultaneamente
- Interface gráfica simples e funcional

---

Fonte dos dados
As informações financeiras das ações são obtidas a partir do Yahoo Finance utilizando a biblioteca yfinance. A disponibilidade e precisão dos dados dependem diretamente desse serviço externo.

---

Requisitos
- Python 3.10 ou superior
- Acesso à internet
- Bibliotecas Python:
  - yfinance
  - openpyxl
  - customtkinter
  - pillow
  - beautifulsoup4

---

Configuração da planilha
O caminho da planilha deve ser ajustado manualmente no código, na variável:

planilha = "C:/caminho/para/sua/planilha.xlsx"

Certifique-se de que a planilha segue o mesmo formato da planilha de exemplo disponibilizada no repositório, incluindo nome da aba e disposição das colunas.

---

Uso
Execute o programa com:

python nome_do_arquivo.py

Digite os tickers das ações separados por espaço.
Exemplo:
PETR4 VALE3 ITUB4

Após a atualização, o programa perguntará se deseja salvar as alterações e abrirá automaticamente o local da planilha.

---

Observações importantes
- Alterações na estrutura do Yahoo Finance podem afetar a obtenção dos dados.
- A planilha não pode estar aberta no Excel durante o salvamento.
- O formato dos tickers deve seguir o padrão da B3, sem o sufixo .SA.

---

Licença
MIT. Consulte o arquivo LICENSE para mais detalhes.

---

Contato
Autor: Jose Davi Silveira Gomes
Email: josedavisilveiragomes@gmail.com
