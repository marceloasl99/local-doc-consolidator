# Como usar: passo a passo (Windows e macOS)

O programa junta todos os seus documentos (PDF, Word, Excel, PowerPoint, imagens...) em **um único arquivo** para enviar ao Copilot. Tudo roda no seu computador: nada é enviado para a internet.

Você só precisa fazer a **Parte 1** uma vez.

---

## 🪟 Windows

### Parte 1: instalar o Python (uma vez só)
1. Acesse **https://www.python.org/downloads/** e clique no botão amarelo **Download Python**.
2. Abra o arquivo baixado.
3. ⚠️ **Na primeira tela, marque a caixinha "Add python.exe to PATH"** (fica embaixo).
4. Clique em **Install Now** e espere terminar. Depois clique em **Close**.

### Parte 2: baixar o programa
1. Na página do projeto no GitHub, clique no botão verde **Code** → **Download ZIP**.
2. Na pasta Downloads, clique com o botão direito no ZIP → **Extrair tudo...** → **Extrair**.

### Parte 3: abrir o programa
1. Abra a pasta extraída.
2. Dê **dois cliques** em **`iniciar_windows.bat`** (português) ou **`start.bat`** (inglês).
3. Se aparecer *"O Windows protegeu o computador"*, clique em **Mais informações** → **Executar assim mesmo**.
4. **Na primeira vez** ele instala o que precisa (alguns minutos, precisa de internet). Nas próximas, abre na hora.

> Se o arquivo aparecer só como `iniciar_windows` (sem `.bat`), é ele mesmo: o ícone tem uma engrenagem.

---

## 🍎 macOS

### Parte 1: instalar o Python (uma vez só)
1. Acesse **https://www.python.org/downloads/** e clique no botão amarelo **Download Python**.
2. Abra o arquivo `.pkg` baixado e clique em **Continuar** até **Instalar** (pede a senha do Mac).
3. Ao final, abre uma pasta do Python: pode fechar.

### Parte 2: baixar o programa
1. Na página do projeto no GitHub, clique em **Code** → **Download ZIP**.
2. Na pasta Downloads, dê dois cliques no ZIP para descompactar.

### Parte 3: abrir o programa
O jeito mais garantido é pelo Terminal (só copiar e colar):
1. Abra o **Terminal**: aperte `⌘ Command + Espaço`, digite **Terminal** e aperte Enter.
2. Digite `bash ` (com um **espaço** no final). **Não** aperte Enter ainda.
3. Arraste o arquivo **`iniciar_macos.command`** (português) ou **`start_macos.command`** (inglês) da pasta do programa **para dentro da janela do Terminal**.
4. Aperte **Enter**.
5. **Na primeira vez** ele instala o que precisa (alguns minutos, precisa de internet).

> Também dá para usar duplo clique no `iniciar_macos.command`. Se o Mac disser que *não pode abrir* ou *desenvolvedor não identificado*, clique com o **botão direito** → **Abrir** → **Abrir**. Se disser *"sem permissão"*, use o jeito do Terminal acima.

---

## 📋 Usando o programa (igual nos dois sistemas)

O programa abre um menu. Você escolhe **digitando o número e apertando Enter**.

1. Digite **`2`** (Pastas de entrada) → **`A`** → escolha a pasta com os documentos na janela que abrir.
   *Repita para outras pastas. Para voltar ao menu, digite `0`.*
2. Digite **`1`** (Gerar base) → confira o resumo → aperte **Enter** para começar.
3. Espere terminar (PDFs digitalizados e imagens demoram mais). No final, digite **`A`** para abrir a pasta do resultado.
4. Para fechar o programa: **`0`**.

### O que enviar ao Copilot
Na pasta do resultado (`copilot_consolidator/saida_copilot/ultimo_processamento`):

| Arquivo | O que fazer |
| --- | --- |
| **`corpus_completo.md`** | **Anexe no Copilot.** É a base com todos os documentos. |
| `corpus_parte_01.md`, `_02`... | Só existe se a base for grande demais: anexe as partes. |
| `PROMPT_PARA_COPILOT.txt` | Abra, copie o texto e cole no chat do Copilot junto com o anexo. |
| `relatorio_auditoria.md` | Mostra o que deu errado (arquivos que não foram lidos, etc.). |

Depois é só perguntar. O Copilot vai citar os documentos pelo número, por exemplo **[004] Contrato.pdf**.

---

## ❓ Problemas comuns

| Problema | Solução |
| --- | --- |
| "Python não foi encontrado" | Refaça a Parte 1. No Windows, **marque "Add python.exe to PATH"**. Depois feche e abra o programa de novo. |
| A janela abre e fecha na hora, ou nada acontece (Windows) | Dê dois cliques no `iniciar_windows.bat`. No PowerShell, digite `.iniciar_windows.bat` (com `.` na frente). Digitar só `start` não funciona: é um comando do Windows. |
| Erro ao instalar as bibliotecas | Verifique a internet e abra o programa de novo. Na rede da empresa, o proxy pode bloquear: tente em outra rede. |
| Um documento ficou "sem texto" | É um PDF digitalizado ou uma imagem sem texto legível. Veja o `relatorio_auditoria.md`. |
| Arquivos `.doc`, `.ppt` ou `.msg` foram ignorados | Abra no Office e use **Salvar como** `.docx`, `.pptx` ou PDF. |
| Quero esconder CPF, RG e nomes | Menu **`3`** → **`20`** (Mascarar dados pessoais). |
