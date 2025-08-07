# Sistema de Geração e Automação de Memorandos

<div align="center">
  <img src="[COLE A URL DE UMA IMAGEM/LOGO PARA O PROJETO AQUI]" alt="Logo do Projeto de Memorandos" width="600">
</div>

<div align="center">
  <img src="https://img.shields.io/badge/Google_Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white" alt="Google Sheets">
  <img src="https://img.shields.io/badge/Apps_Script-4285F4?style=for-the-badge&logo=google-apps-script&logoColor=white" alt="Google Apps Script">
  <img src="https://img.shields.io/badge/html5-%23E34F26.svg?style=for-the-badge&logo=html5&logoColor=white" alt="HTML5">
  <img src="https://img.shields.io/badge/bootstrap-%238511FA.svg?style=for-the-badge&logo=bootstrap&logoColor=white" alt="Bootstrap">
</div>

## 📜 Sobre o Projeto

Sistema desenvolvido para automatizar a criação, o preenchimento e o arquivamento de memorandos para unidades administrativas da Polícia Militar, utilizando o poder do Google Sheets e Google Apps Script.

Este projeto nasceu da necessidade de otimizar e padronizar o processo de criação de memorandos. O fluxo de trabalho manual, suscetível a erros e demorado, foi transformado em um sistema web simples e centralizado, que guia o usuário, preenche os dados automaticamente e arquiva o documento final em PDF no Google Drive.

O sistema garante a padronização dos documentos, a integridade dos dados (puxando informações de uma base central) e cria um registro histórico confiável e de fácil acesso.

## 🏛️ Acesso ao Sistema e Documentação

Para uma experiência completa, incluindo o acesso ao sistema e manuais de usuário, visite o site oficial do projeto construído com Google Sites.

**➡️ Acesse o Portal do Sistema Aqui: https://sites.google.com/view/geradormemorando/ **

---

## ✨ Funcionalidades Principais

* **Interface Web Amigável:** Um formulário simples e intuitivo para o usuário preencher os dados do memorando.
* **Seleção Dinâmica de Pessoal:** A lista de policiais é carregada automaticamente de uma base de dados, evitando erros de digitação.
* **Geração Automática de PDF:** Com um clique, o sistema preenche o modelo de memorando e gera um arquivo PDF profissional.
* **Arquivamento Automático:** O PDF gerado é salvo automaticamente em uma pasta designada no Google Drive, com um nome padronizado contendo o nome do interessado e a data/hora da criação.
* **Data Automatizada:** A data no documento é sempre a data atual, graças à fórmula `=TODAY()` na planilha modelo.

## 🛠️ Arquitetura e Tecnologias Utilizadas

O sistema é construído inteiramente dentro do ecossistema Google, garantindo integração, segurança e baixo custo.

| Componente                | Tecnologia                  | Propósito                                                                      |
| :------------------------ | :-------------------------- | :----------------------------------------------------------------------------- |
| **Base de Dados e Template** | Google Sheets               | Armazena a lista de pessoal e serve como modelo para o memorando.              |
| **Motor e Lógica** | Google Apps Script          | Executa todas as tarefas de back-end: serve a interface, busca dados, preenche a planilha e gera o PDF. |
| **Interface do Usuário (UI)** | HTML, CSS com Bootstrap     | Cria o formulário web que o usuário final interage.                              |
| **Armazenamento de Arquivos** | Google Drive                | Guarda os PDFs finais de forma organizada.                                     |
| **Portal de Documentação** | Google Sites                | Centraliza o acesso ao sistema e a documentação do usuário.                      |

### Como a Planilha é Estruturada

A Planilha Google é o coração do sistema, dividida em duas abas principais:

1.  **Aba `MEMORANDO` (Página 1):**
    * Serve como o template visual do documento a ser gerado.
    * **Célula `D3`**: Contém a fórmula `=TODAY()` e utiliza a formatação personalizada `"Campinas, "d" de "mmmm" de "yyyy"` para exibir a data atual de forma elegante.
    * **Célula `E8`**: Ponto central de entrada dos dados do policial. Esta célula:
        * Recebe o nome do policial selecionado através da interface do Web App.
        * Possui uma **Validação de Dados** (Menu suspenso de um intervalo) configurada para permitir a seleção manual diretamente na planilha. O intervalo para esta validação é `'BASE DE DADOS'!$B$3:$B$150`, que corresponde à lista de nomes concatenados, garantindo consistência entre o uso manual e o automatizado.

2.  **Aba `BASE DE DADOS` (Página 2):**
    * **Fonte de Dados:** Puxa a lista de efetivo de uma outra planilha central através da fórmula:
        ```excel
        =IMPORTRANGE("URL_DA_PLANILHA_FONTE"; "EFETIVO!B10:F50")
        ```
    * **Concatenação de Nomes (Coluna B):** Em uma coluna auxiliar (a coluna `B`, que é usada pela Validação de Dados), utiliza a fórmula `TEXTJOIN` para formatar os nomes que serão exibidos na lista suspensa.
        ```excel
        =TEXTJOIN("-"; TRUE; TEXTJOIN(" "; TRUE; G2; H2); TEXTJOIN(" "; TRUE; I2; J2))
        ```

## 🚀 Como Replicar o Projeto (Setup)

Caso queira criar sua própria versão deste sistema, siga os passos abaixo:

1.  **Crie a Planilha Google:** Crie uma nova Planilha com as abas `MEMORANDO` e `BASE DE DADOS` e configure as fórmulas (`=TODAY()`, `=IMPORTRANGE`, etc.) conforme descrito acima.
2.  **Crie o Script:** Na sua planilha, vá em `Extensões > Apps Script`.
3.  **Adicione o Código:**
    * Crie um arquivo de script chamado `WebApp.gs` e copie o conteúdo do arquivo [`WebApp.gs`](./WebApp.gs) deste repositório.
    * Crie um arquivo HTML chamado `Index.html` e copie o conteúdo do arquivo [`Index.html`](./Index.html) deste repositório.
4.  **Atualize os IDs:** No arquivo `WebApp.gs`, atualize o ID da pasta do Google Drive na função `salvarPdfNoDrive` para a sua pasta de destino.
5.  **Implante o Web App:**
    * Clique em `Implantar > Nova implantação`.
    * Selecione o tipo "App da Web".
    * Configure o acesso para "Qualquer pessoa com uma Conta do Google".
    * Clique em `Implantar` e autorize as permissões necessárias.
6.  **Atualize a URL:** Copie a URL do Web App implantado e cole-a na função `abrirWebApp` dentro do seu `WebApp.gs`.
7.  **Reimplante:** Faça uma nova implantação (como "Nova versão") para que a alteração do passo 6 tenha efeito.

## 👨‍💻 Código-Fonte do Projeto

O código-fonte do sistema está contido neste repositório. Os arquivos principais são:

* **[`WebApp.gs`](./WebApp.gs):** O script do Google Apps Script que contém toda a lógica de back-end. Ele serve a interface, busca dados da planilha, preenche o template, gera o PDF e o salva no Google Drive.

* **[`Index.html`](./Index.html):** O arquivo HTML que define a estrutura da interface web, incluindo o formulário de entrada de dados e o código JavaScript do lado do cliente para interagir com o back-end.

## ⚖️ Licença

Este projeto é distribuído sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.