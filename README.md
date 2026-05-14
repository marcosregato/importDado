# importDado - Sistema de Importação de Dados Excel para SQL

Sistema Java para converter planilhas Excel em scripts SQL de criação de tabelas, com detecção automática de tipos de dados.

## 🚀 Funcionalidades

- **Conversão Excel para SQL**: Gera automaticamente scripts CREATE TABLE a partir de planilhas Excel
- **Detecção Automática de Tipos**: Identifica tipos de dados (INT, VARCHAR, DATE, DECIMAL, BOOLEAN)
- **Suporte Multi-aba**: Processa múltiplas abas de uma planilha
- **Configuração Flexível**: Arquivo de configuração para personalizar caminhos e parâmetros
- **Cross-Platform**: Funciona em Windows, Linux e macOS

## 📋 Pré-requisitos

- Java 21 ou superior
- Maven 3.6 ou superior

## 🔧 Instalação

1. Clone o repositório:
```bash
git clone https://github.com/marcosregato/importDado.git
cd importDado
```

2. Compile o projeto com Maven:
```bash
mvn clean install
```

## ⚙️ Configuração

Edite o arquivo `src/config.properties` para personalizar:

```properties
# Caminhos (use relativos para cross-platform)
base.path=src/
excel.file=planilha_teste_2.xls
output.directory=arqSQL/
default.sheet.name=TURMA

# Limites de processamento
max.columns.to.read=50
max.rows.to.read=1000
```

## 🎯 Como Usar

### Compilação
```bash
mvn clean compile
```

### Execução
```bash
mvn exec:java -Dexec.mainClass="importaDado.CriarTableSQL"
```

Ou após o build:
```bash
java -cp target/importDado-1.0-SNAPSHOT.jar:target/lib/* importaDado.CriarTableSQL
```

### Saída
O sistema irá:
1. Ler as abas da planilha Excel
2. Extrair cabeçalhos das colunas
3. Detectar tipos de dados automaticamente
4. Gerar arquivo SQL em `src/arqSQL/tables.sql`

## 📊 Exemplo de Saída SQL

```sql
CREATE TABLE TURMA (
    id INT PRIMARY KEY NOT NULL AUTO_INCREMENT,
    ALUNO VARCHAR(50),
    PERIODO VARCHAR(50),
    MATRICULA DECIMAL(10,2)
);

CREATE TABLE PROFESSOR (
    id INT PRIMARY KEY NOT NULL AUTO_INCREMENT,
    NOME VARCHAR(50),
    MATERIA VARCHAR(50)
);
```

## 🔍 Detecção de Tipos

O sistema detecta automaticamente:
- **INT**: Números inteiros
- **DECIMAL(10,2)**: Números com casas decimais
- **DATE**: Datas (formatos com / ou -)
- **BOOLEAN**: Valores true/false, sim/não, yes/no
- **VARCHAR(n)**: Textos (tamanho ajustado automaticamente)

## 📁 Estrutura do Projeto

```
importDado/
├── src/
│   ├── config.properties          # Arquivo de configuração
│   ├── importaDado/
│   │   └── CriarTableSQL.java     # Classe principal de geração SQL
│   ├── manipularArquivo/
│   │   └── ManipularArquivo.java  # Classe legada (compatibilidade)
│   ├── util/
│   │   ├── ConfigManager.java     # Gerenciador de configuração
│   │   ├── ExcelReader.java       # Leitor unificado Excel
│   │   └── Utils.java            # Utilitários
│   └── rascunho/
│       └── SqlExcel.java          # Código em desenvolvimento
├── pom.xml                        # Configuração Maven
└── README.md                      # Este arquivo
```

## 🛠️ Tecnologias Utilizadas

- **Java 21**: Linguagem principal
- **Apache POI 5.2.5**: Processamento de arquivos Excel
- **JXL 2.6.12**: Suporte legado para arquivos .xls
- **Fillo 1.21**: Manipulação de dados Excel
- **Log4j 2.20.0**: Sistema de logging
- **Maven**: Gerenciamento de dependências e build

## 🐛 Solução de Problemas

### Erros Comuns

- **Arquivo não encontrado**: Verifique o caminho em `config.properties`
- **Classe não encontrada**: Certifique-se de executar `mvn clean install` primeiro
- **Aba não encontrada**: Verifique `default.sheet.name` na configuração
- **Erro de dependência**: Execute `mvn clean install` para baixar as dependências

### Logs

O sistema exibe logs detalhados do processo:
- Arquivos sendo processados
- Abas encontradas
- Colunas detectadas
- Erros específicos

## 📝 Desenvolvimento

### Adicionar novas funcionalidades

1. Crie novas classes em `src/util/` para funcionalidades reutilizáveis
2. Adicione novas dependências no `pom.xml`
3. Atualize `src/config.properties` se necessário

### Testes

Para testar com sua própria planilha:
1. Coloque o arquivo .xls em `src/`
2. Atualize `excel.file` em `config.properties`
3. Execute o projeto

## 📄 Licença

Este projeto é de código aberto e pode ser modificado conforme necessidade.

## 👤 Autor

Marcos Regato

## 🤝 Contribuindo

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.
