# R.A - Equivalencias — Equivalência & Alunos

Aplicativo desktop em Python com interface Tkinter, inspirado no portal de referência enviado. A aplicação consulta e grava os dados em uma pasta de trabalho Excel (`.xlsx`).

## Recursos

- Fluxo de navegação com cadastro de alunos antes de solicitar aproveitamento, além de um Dashboard com gráfico de solicitações por status e o maior tempo de espera em aberto.
- Configuração do ano letivo e do prefixo usado para gerar novos códigos de processo no formato `PREFIXO-ANO-001`.
- Cadastro, edição, busca e exclusão de processos.
- Cadastro separado de alunos e importação automática dos campos reconhecidos em relatórios PDF.
- Seleção de um aluno pré-cadastrado ao abrir um processo, com registro do vínculo pelo ID do processo.
- Busca por ID do processo, matrícula, nome, CPF e demais campos cadastrados.
- Requerimento preenchido a partir do processo selecionado, com geração de página HTML pronta para imprimir ou salvar como PDF pelo navegador.
- Cadastro, edição, remoção e exportação de equivalências de UCs.
- Planilha principal configurável com as abas `Processos`, `Equivalencias` e `Usuarios`; a base de alunos fica em outro arquivo Excel, na aba `Alunos`.
- Criação de uma planilha vazia no primeiro uso. Nenhum cadastro de exemplo é incluído.

## Fluxo de uso

1. **Acesse o sistema:** abra `run_sistema.bat`. Na primeira abertura, crie o usuário administrador; depois, entre com seu usuário e senha.
2. **Configure o ambiente:** em **Configurações (Planilha)**, defina o ano letivo, o prefixo do código do processo e os caminhos da planilha principal e da base separada de alunos. Novos IDs seguem `PREFIXO-ANO-001`.
3. **Cadastre o aluno:** em **Cadastro de Alunos**, preencha os dados manualmente ou use **Importar alunos (PDF)** para criar pré-cadastros a partir do relatório Senac.
4. **Abra a solicitação:** em **Solicitar Aproveitamento**, selecione o aluno pré-cadastrado, confira e complete os dados do processo, depois clique em **Salvar cadastro**. Também é possível buscar processos existentes.
5. **Registre as equivalências:** selecione o processo ativo, adicione as UCs e edite ou exporte o mapeamento quando necessário.
6. **Emita o documento:** abra **Requerimento Oficial**, confira a prévia e use **Abrir versão para impressão**. No navegador, imprima ou salve como PDF.
7. **Acompanhe:** o **Dashboard** abre junto com o sistema e mostra os processos por status e a maior espera entre as solicitações abertas.

O passo a passo completo está em [Guia de uso do sistema](guia_uso.html). Abra esse arquivo no navegador; ele funciona sem internet e pode ser impresso.

## Requisitos

- Python 3.10 ou superior com Tkinter habilitado.
- Biblioteca `openpyxl`.
- Biblioteca `pypdf` para ler relatórios PDF com texto selecionável.
- Biblioteca `cryptography` para abrir PDFs protegidos com AES.

## Instalação e execução no Windows

Abra `run_sistema.bat`. Na primeira execução, o atalho cria um ambiente Python local em `.venv` e instala as dependências listadas em `requirements.txt`. É necessário acesso à internet para baixá-las.

Se preferir configurar pelo PowerShell:

```powershell
py -m pip install -r requirements.txt
py main.py
```

Ao executar `main.py` diretamente pelo IDE, selecione `.venv\Scripts\python.exe` depois que o atalho concluir a configuração. Para preparar outro interpretador, instale todas as dependências com `py -m pip install -r requirements.txt`.

O workspace inclui uma tarefa do VS Code chamada **R.A - Equivalencias: iniciar interface**. Use `Ctrl+Shift+B` para iniciar a janela gráfica em um terminal próprio. Para esta aplicação Tkinter, essa tarefa é a forma recomendada de executar dentro do VS Code.

## Planilha

Na primeira abertura, o aplicativo usa `Base_Equivalencias_Senac.xlsx` para processos e usuários, e cria `Base_Alunos.xlsx` no mesmo diretório para os alunos. Em **Configurações (Planilha)**, os dois caminhos podem ser escolhidos separadamente; eles ficam guardados em `portal_config.json`.

A planilha principal armazena processos, equivalências e usuários. A base separada `Base_Alunos.xlsx` armazena uma ficha por aluno e os IDs dos processos relacionados. Ao atualizar um processo, o sistema atualiza o vínculo na base de alunos pelo ID do processo.

Na tela **Cadastro de Alunos**, use **Importar alunos (PDF)** e selecione o relatório. O sistema reconhece o layout tabular do Senac (**Relatório de Dados e Status do Aluno**) e procura ID, nome, RG, CPF, data de nascimento e data de matrícula em cada linha de aluno, mesmo quando há colunas vazias. Os registros encontrados são gravados na base separada de alunos como pré-cadastros. Na tela **Solicitar Aproveitamento**, selecione o aluno pré-cadastrado; ao salvar o processo, o sistema registra o vínculo pelo ID do processo. O conteúdo é lido localmente. PDFs protegidos por senha diferente de vazia ou sem texto selecionável não são importados.

Reimportações com o mesmo CPF atualizam a ficha e preservam o ID já cadastrado e os vínculos com processos. Se o ID e o CPF apontarem para alunos diferentes, se um ID aparecer com CPFs diferentes ou se houver cadastros duplicados para os identificadores envolvidos, o sistema cancela o lote inteiro e informa o conflito. Corrija os dados indicados antes de tentar novamente. A mensagem de sucesso conta cada aluno atualizado uma única vez.

## Arquivos

- `guia_uso.html`: guia passo a passo, disponível para abrir no navegador sem internet.
- `main.py`: interface gráfica e geração do requerimento para impressão.
- `pdf_import.py`: extração local dos campos dos relatórios PDF.
- `storage.py`: leitura e gravação da pasta de trabalho Excel.
- `requirements.txt`: dependências Python.
- `run_sistema.bat`: atalho para abrir o aplicativo no Windows.

## Acesso e senhas

Na primeira abertura, crie o usuário administrador na tela inicial. Depois, entre com esse usuário e use o botão **Usuários** no topo da janela para cadastrar outras contas e suas funções. Usuários com a função **Administrador** podem cadastrar novos usuários. Contas criadas pelo administrador precisam trocar a senha provisória no primeiro acesso; todos os usuários também podem usar **Alterar senha** no topo do sistema.

As contas ficam na aba **Usuarios** da mesma planilha selecionada. O sistema grava salt e hash PBKDF2-SHA256 (600.000 iterações), além de um indicador para a troca obrigatória; a senha original não é armazenada nem pode ser recuperada. Se a planilha for trocada em **Configurações**, as contas atuais são copiadas para ela junto com seus hashes.
