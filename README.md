<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">

<body>

  <h1>Envio Automático de E-mails via Google Sheets</h1>
  <p>Este projeto permite o envio automático de e-mails a partir de dados armazenados em uma planilha do Google Sheets. Ele utiliza o Google Apps Script para processar os dados e enviar e-mails com base em links <code>mailto:</code> presentes na planilha.</p>

  <h2>Funcionalidades</h2>
  <ul>
    <li>Extrai dados de e-mail, assunto e corpo de mensagem de links <code>mailto:</code>.</li>
    <li>Decodifica URLs de forma segura para evitar erros.</li>
    <li>Envia e-mails com formatação HTML.</li>
    <li>Atualiza o status de envio na planilha ("Enviado" ou "Erro").</li>
  </ul>

  <h2>Pré-requisitos</h2>
  <ol>
    <li><strong>Google Sheets</strong>: Uma planilha com uma coluna contendo links <code>mailto:</code>.</li>
    <li><strong>Google Apps Script</strong>: O código é executado no ambiente do Google Apps Script.</li>
    <li><strong>Permissões</strong>: O script precisa de permissão para acessar a planilha e enviar e-mails.</li>
  </ol>

  <h2>Como Usar</h2>

  <h3>1. Configuração da Planilha</h3>
  <ul>
    <li>Crie uma planilha no Google Sheets com o nome <code>Página1</code>.</li>
    <li>Na coluna <strong>F</strong> (ou coluna 6), insira os links <code>mailto:</code> com os dados do e-mail.</li>
    <li>Na coluna <strong>J</strong> (ou coluna 10), o script atualizará o status de envio.</li>
  </ul>

  <p>Exemplo de link <code>mailto:</code>:</p>
  <pre><code>mailto:exemplo@dominio.com?subject=Assunto do E-mail&body=Corpo do E-mail</code></pre>

  <h3>2. Configuração do Script</h3>
  <ol>
    <li>Abra a planilha no Google Sheets.</li>
    <li>Vá para <code>Extensões</code> > <code>Apps Script</code>.</li>
    <li>Substitua o conteúdo do arquivo <code>Code.gs</code> pelo código fornecido neste repositório.</li>
    <li>Salve o projeto com um nome, por exemplo, <code>Envio de E-mails</code>.</li>
  </ol>

  <h3>3. Executando o Script</h3>
  <ol>
    <li>No editor do Apps Script, selecione a função <code>myFunction</code> no menu de execução.</li>
    <li>Clique em <code>Executar</code>.</li>
    <li>Na primeira execução, o Google solicitará permissões para acessar a planilha e enviar e-mails. Conceda as permissões necessárias.</li>
    <li>O script processará a planilha e enviará os e-mails.</li>
  </ol>

  <h3>4. Monitoramento</h3>
  <ul>
    <li>O status de envio será atualizado na coluna <strong>J</strong> da planilha.</li>
    <li>Logs detalhados podem ser visualizados em <code>Execuções</code> no menu do Apps Script.</li>
  </ul>

  <h2>Estrutura do Código</h2>
  <p>O código foi desenvolvido seguindo o princípio de <strong>Responsabilidade Única (SRP)</strong>, com funções separadas para cada tarefa:</p>
  <ul>
    <li><strong><code>safeDecodeURI</code></strong>: Decodifica URLs de forma segura.</li>
    <li><strong><code>getEmailData</code></strong>: Extrai e formata os dados do e-mail.</li>
    <li><strong><code>sendEmail</code></strong>: Envia o e-mail e retorna o status.</li>
    <li><strong><code>processSheet</code></strong>: Processa a planilha e chama as funções acima.</li>
    <li><strong><code>myFunction</code></strong>: Coordena a execução do script.</li>
  </ul>

  <h2>Exemplo de Planilha</h2>
  <table>
    <thead>
      <tr>
        <th>Nome</th>
        <th>E-mail (Coluna F)</th>
        <th>Status (Coluna J)</th>
      </tr>
    </thead>
    <tbody>
      <tr>
        <td>João Silva</td>
        <td><code>mailto:joao@dominio.com?subject=Assunto&body=Corpo%20do%20e-mail</code></td>
        <td>Enviado</td>
      </tr>
      <tr>
        <td>Maria Souza</td>
        <td><code>mailto:maria@dominio.com?subject=Outro%20Assunto&body=Outro%20corpo%20de%20e-mail</code></td>
        <td>Erro</td>
      </tr>
    </tbody>
  </table>

  <h2>Contribuição</h2>
  <p>Contribuições são bem-vindas! Sinta-se à vontade para abrir <a href="https://github.com/seu-usuario/seu-repositorio/issues">issues</a> ou <a href="https://github.com/seu-usuario/seu-repositorio/pulls">pull requests</a>.</p>

  <h2>Licença</h2>
  <p>Este projeto está licenciado sob a <a href="LICENSE">MIT License</a>.</p>

</body>
</html>
