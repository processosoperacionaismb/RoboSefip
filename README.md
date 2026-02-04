# 🤖 Robô de Automação SEFIP (v1.6)

Automação inteligente para processamento em lote no SEFIP. O sistema realiza a limpeza da base, importação de arquivos .RE, cadastro de funcionários e retificação de valores automaticamente.

📋 Pré-requisitos
Python 3.x instalado.

SEFIP instalado e com as telas padrão (sem redimensionamento).

Pasta imagens/ contendo os prints dos botões do sistema.

🛠️ Instalação (Rápida)
Abra o terminal ou CMD na pasta do projeto.

Instale todas as dependências de uma vez executando:

Bash

pip install -r requirements.txt

🚀 Manual de Uso

PREPARAÇÃO:

Certifique-se de que seus arquivos SEFIP estão na pasta: C:\Robo_SEFIP\[ANO]\[ANOMES]\SEFIP.RE.

1 -> Abra e Mantenha o SEFIP aberto na tela inicial.

2 -> Execute o robô: Botão de Play (verde) no Pycharm.

3 -> Clique em 📄 Criar Modelo CSV.

Abra o arquivo gerado e preencha as colunas ano, mes e valor.

No programa, clique em 📁 CSV e selecione o arquivo que você editou.

EXECUÇÃO:

Clique em ▶ Iniciar Processamento em Lote.

Aguarde o robô terminar.

Não utilize o computador enquanto o robô estiver movendo o mouse.

🔄 Recursos da Versão 1.6
Tratamento de Erros: Se um botão não for encontrado, o robô perguntará se você deseja Tentar Novamente, Pular ou Cancelar.

Logs em CSV: Relatórios automáticos gerados na pasta /logs para conferência de quais meses foram processados com sucesso.

Parada de Emergência: Caso precise parar o robô imediatamente, arraste o mouse para qualquer canto da tela.
