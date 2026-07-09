# Analista Field

App Streamlit para analisar planilhas XLSX de eventos de campo.

## Funcionalidades

- Upload de planilha `.xlsx` ou `.xlsm`.
- Selecao da aba a ser analisada.
- Deteccao automatica das colunas de instalador, cidade, bairro, polo, regional, tipo de evento, problema e data.
- Ajuste manual do mapeamento quando a deteccao nao for suficiente.
- Filtros por regional, polo, tipo de evento, problema, instalador, cidade, bairro e periodo.
- Explorador visual interativo com pizza de participacao e Pareto/ranking.
- Rankings por instalador, cidade, bairro, polo, dia, mes, tipo de evento e regional.
- Download da analise em Excel.
- Relatorio PDF da analise realizada, com a marca Valenet no cabecalho.

## Como rodar

```bash
cd analista_field
streamlit run app.py
```

## Publicacao

Arquivo principal:

```text
app.py
```

Comando de start:

```bash
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

Para Streamlit Community Cloud, conecte este repositorio e selecione `app.py` como arquivo principal.
