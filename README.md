# ⚙️ Automatização de Histórico e Proteção de Fórmulas para Google Sheets

Projeto em Google Apps Script criado para proteger e restaurar fórmulas em planilhas compartilhadas, garantindo a integridade dos seus dados de forma automática. Chega de dores de cabeça com fórmulas apagadas por engano! getSheetByName

## 🤔 Qual problema este projeto resolve?

Em planilhas com muitos colaboradores, é comum que alguém acidentalmente apague ou substitua uma fórmula importante por um valor fixo (`=SOMA(A1:A10)` vira `1500`). Isso quebra a automação da planilha e pode gerar dados incorretos. Este script monitora suas células e restaura a fórmula original sempre que uma alteração indevida é detectada.

## ✨ Funcionalidades Principais

* **🛡️ Proteção de Fórmulas:** Monitora células específicas e impede que as fórmulas sejam substituídas permanentemente.
* **🔄 Restauração Automática:** Caso uma fórmula seja alterada, o script a restaura para a versão original.
* **🎛️ Menu Personalizado:** Adiciona um menu simples à sua planilha para facilitar a execução manual.
* **⚙️ Fácil Configuração:** Você só precisa indicar quais abas e células deseja proteger.

## 🚀 Instalação e Configuração (Passo a Passo)

Siga estas instruções para adicionar o projeto à sua Planilha Google:

1.  **Abra sua Planilha:** Acesse a planilha que você deseja proteger.
2.  **Abra o Editor de Script:** No menu superior, clique em `Extensões` > `Apps Script`.
3.  **Limpe o Código Padrão:** Apague todo o conteúdo do arquivo `Código.gs` que vem por padrão.
4.  **Crie os Arquivos do Projeto:**
    * **Arquivo `corrigeFormulas.gs`:**
        * No editor do Apps Script, renomeie o arquivo `Código.gs` para `corrigeFormulas.gs`.
        * Copie todo o conteúdo do arquivo [`corrigeFormulas.gs`](https://github.com/jackyebattista/public_projects/blob/main/corrigeFormulas.gs) do meu repositório e cole no seu arquivo.
    * **Arquivo `automation.gs`:**
        * Clique no ícone de `+` e selecione `Script` para criar um novo arquivo. Nomeie-o como `automation.gs`.
        * Copie o conteúdo do arquivo [`automation.gs`](https://github.com/jackyebattista/public_projects/blob/main/automation.gs) e cole nele.
    * **Arquivo `index.html`:**
        * Clique no ícone de `+` e selecione `HTML` para criar um novo arquivo. Nomeie-o como `index.html`.
        * Copie o conteúdo do arquivo [`index.html`](https://github.com/jackyebattista/public_projects/blob/main/index.html) e cole nele.

5.  **Configure suas Planilhas:**
    * Dentro do arquivo `corrigeFormulas.gs`, localize a seção de configurações (geralmente no início do arquivo).
    * Altere as variáveis para indicar os nomes das suas abas e os intervalos de células que o script deve proteger (ex: `A1:F50`).

6.  **Salve e Autorize:**
    * Clique no ícone de disquete 💾 para salvar todos os arquivos.
    * Recarregue a sua Planilha Google. Um novo menu deve aparecer. Ao clicar em uma das opções pela primeira vez, o Google solicitará permissão para que o script seja executado. **É necessário autorizar para que tudo funcione.**

## 🕹️ Como Usar

Após a instalação, um novo menu (ex: "🛠️ Ferramentas de Automação") aparecerá na barra de menus da sua planilha.

* **Para verificar e corrigir as fórmulas manualmente**, clique no menu e selecione a opção correspondente.
* O script também pode ser configurado para rodar automaticamente através de **acionadores (triggers)**, como "Ao abrir" a planilha ou "Baseado em tempo" (a cada hora, por exemplo).

## 🤝 Como Contribuir

Sua ajuda é muito bem-vinda para tornar este projeto ainda melhor!

* **Reportar Erros:** Encontrou um problema? Abra uma [**Issue**](https://github.com/jackyebattista/public_projects/issues) detalhando o erro.
* **Sugerir Melhorias:** Tem uma ideia para uma nova funcionalidade? Abra uma [**Issue**](https://github.com/jackyebattista/public_projects/issues) para discutirmos.
* **Enviar Código:** Quer contribuir diretamente?
    1.  Faça um **Fork** deste repositório.
    2.  Crie uma nova *branch* para sua modificação (`git checkout -b feature/minha-melhoria`).
    3.  Faça o *commit* das suas alterações (`git commit -m 'Adiciona minha melhoria'`).
    4.  Envie para a sua *branch* (`git push origin feature/minha-melhoria`).
    5.  Abra um **Pull Request**.

---
Feito com ❤️ por Jaqueline Battista (https://github.com/jackyebattista).
Conecte-se comigo: https://br.linkedin.com/in/jaquelinebattista
