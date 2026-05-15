# Meus Macros VBA

Esse repositório é onde eu organizo os meus scripts e macros em VBA. A ideia principal aqui é automatizar aquelas tarefas repetitivas do Excel no dia a dia e deixar o fluxo de trabalho mais rápido.

## O que você vai encontrar aqui

Por enquanto eu tenho três macros principais, que dividi em duas pastas para manter o código organizado:

### Formatação (/Formatacao)
* **Ajuste de Largura de Coluna:** São duas macros simples, mas que salvam muito tempo. Elas servem para aumentar ou diminuir a largura da coluna selecionada usando apenas o teclado, sem precisar tirar a mão para usar o mouse.

### Exportação (/Exportacao)
* **Excel para PDF ou Imagem:** Um script para agilizar na hora de exportar partes específicas da planilha, ou ela inteira, para PDF ou imagem. É bem prático para gerar relatórios fechados antes de enviar para alguém.

---

## Como instalar e configurar na sua máquina

O Git não lida muito bem com arquivos do Excel direto, então deixei os códigos exportados em arquivos `.bas`. 

Para que as macros funcionem bem em qualquer planilha que você abrir (principalmente as de atalho de coluna), o ideal é salvar o código no seu arquivo pessoal oculto do Excel (`PERSONAL.XLSB`). Aqui está o passo a passo de como fazer isso:

### Passo 1: Importar o código para o Excel

1. Baixe o arquivo `.bas` que você quer usar aqui do repositório.
2. Abra o Excel e aperte `Alt + F11` para abrir o editor do VBA.
3. No painel da esquerda (Project Explorer), procure por `VBAProject (PERSONAL.XLSB)`. 
   *Nota: Se você não tiver esse arquivo, grave uma macro vazia qualquer e escolha salvar na "Pasta de Trabalho Pessoal de Macros" para forçar o Excel a criar o arquivo para você.*
4. Clique no projeto do `PERSONAL.XLSB` para selecioná-lo.
5. Vá no menu **Arquivo > Importar Arquivo** (ou aperte `Ctrl + M`).
6. Selecione o arquivo `.bas` que você baixou e confirme. Pode fechar o editor do VBA.

### Passo 2: Configurar os atalhos de teclado

Agora precisamos conectar a macro a um atalho rápido no seu teclado.

1. Na tela normal do Excel, aperte `Alt + F8` para abrir a lista de macros.
2. No campo "Macros em:", mude para "PERSONAL.XLSB" para conseguir ver as macros que você acabou de importar.
3. Selecione a macro na lista (por exemplo, a de aumentar coluna) e clique no botão **Opções...**.
4. Escolha a letra para o atalho. Eu recomendo  (`Ctrl + D`) para aumentar, e (`Ctrl + A`) para diminuir.
5. Clique em OK e feche a janela.

**Muito importante:** Quando você fechar o Excel pela primeira vez após fazer isso, ele vai perguntar se você deseja salvar as alterações na Pasta de Trabalho Pessoal de Macros. Clique em **Salvar**, caso contrário, os atalhos que você configurou serão perdidos.

---

Feito com VBA e Markdown.
