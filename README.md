# 🧾 AutoCert

Uma aplicação moderna e intuitiva para geração automática de certificados personalizados com envio por email. Desenvolvido em Python com interface gráfica aprimorada utilizando `ttkbootstrap`.

## 🎯 Funcionalidades

- Interface gráfica moderna e responsiva (modo escuro e claro);
- Geração de certificados em lote a partir de planilhas Excel;
- Personalização de posição, fonte e tamanho do texto;
- Pré-visualização em tempo real dos certificados;
- Envio automático por email com conteúdo personalizado;
- Armazenamento seguro de credenciais no arquivo `config.ini`.

## 🧰 Tecnologias Utilizadas

- `Python 3.10+`
- `ttkbootstrap` (interface gráfica moderna)
- `Pillow` (manipulação de imagens)
- `pandas` (leitura de planilhas)
- `yagmail` (envio de emails)
- `tkinter` (interface gráfica base)
- `configparser` (gerenciamento de configurações)

## Imagens da Interface

![Configurações](https://github.com/gabrielbtt/AutoCert/blob/main/docs/Config.png)
![Design](https://github.com/gabrielbtt/AutoCert/blob/main/docs/Design.png)

## 📦 Instalação

Clone o repositório:

git clone https://github.com/gabrielbtt/AutoCert.git
cd AutoCert

Instale as dependências:

bash
Copiar
Editar
pip install -r requirements.txt
Obs: Certifique-se de estar utilizando um ambiente com acesso a fontes do Windows para funcionamento completo.

## 🚀 Como Usar
Dentro do aplicativo:

Selecione o modelo do certificado (formato .png, .jpg ou .jpeg);

Escolha a planilha com os dados (.xlsx);

Ajuste as posições e estilos do texto;

Clique em Enviar Certificados para começar.

Use {name} no corpo do email para personalizar com o nome do destinatário.

## 📁 Estrutura da Planilha
A planilha .xlsx deve conter obrigatoriamente as seguintes colunas:

Nome	Email	Numero do Certificado
João da Silva	joao@email.com	001
Maria Souza	maria@email.com	002

## Vídeo Tutorial

[Assista ao tutorial completo](https://www.youtube.com/watch?v=ImA_r1pWFK0)

## 🔒 Segurança
As credenciais são armazenadas localmente no arquivo config.ini;

Recomenda-se usar uma conta de email com senha de aplicativo (Gmail).

## ❗ Problemas Conhecidos
O aplicativo depende de fontes do Windows. Em outros sistemas operacionais, pode ser necessário adaptar o caminho das fontes.

## 👨‍💻 Autor
Desenvolvido por [Gabriel Batista]([https://www.linkedin.com/in/gabrielbtt/])

## 📃 Licença
Este projeto está licenciado sob a MIT License.

