# Histórico de Versões - Controles de Combustiveis

## [1.0.3] - 2026-03-10

### Adicionado
- **Licença do Projeto**: Adicionado arquivo de licença Proprietária com restrição de uso comercial, garantindo os direitos do autor sobre a exploração lucrativa do software.

### Alterado
- **Qualidade do Código**: O código-fonte dos arquivos `configurador.py` e `combustivel.pyw` foi revisado para padronizar e enriquecer os comentários e a documentação (docstrings), melhorando a clareza e a manutenibilidade sem alterar a lógica funcional.

## [1.0.2] - 2026-03-09

### Corrigido
- Erro `WinError 123` que impedia o salvamento de arquivos devido a aspas extras nos caminhos de configuração.
- Falha na compilação do `.exe` que impedia a leitura de e-mails existentes (dependência `win32timezone` não era incluída).
- Erro de permissão (`Acesso Negado`) durante a compilação ao forçar o encerramento de processos antigos.
- Avisos de sintaxe (`SyntaxWarning`) nos scripts Python.

### Alterado
- O nome do arquivo de ícone foi generalizado para `icone.ico` para proteger a identidade visual da empresa.
- O script de compilação (`compilar.bat`) foi aprimorado para usar o ícone apenas se ele existir, não gerando erro caso esteja ausente.

## [1.0.1] - 2026-03-09

### Alterado
- Unificação da versão do Configurador e do Robô.
- Ajustes na interface e lógica de histórico.

## \[1.0.0] - 2026-03-06

### Adicionado

-Entregue a 1° Versão
