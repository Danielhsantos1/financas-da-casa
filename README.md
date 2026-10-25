# Finanças da Casa

Controle financeiro da casa (Daniel e Suélen): contas fixas por mês, marcar pago/pendente,
gastos variáveis e quanto sobra no mês. Feito para usar no celular.

## Como funciona
- `index.html`: o app (uma página só). Lê e grava numa planilha Google pela ponte abaixo.
- `apps-script/Code.gs`: ponte que fica dentro da planilha Google (Apps Script). Guarda os dados nas abas
  `contas`, `lanc`, `gastos`, `entradas` e monta a aba `Visao` (grade mensal, boa para Excel).
- A chave de acesso NÃO fica no código. O app pede a chave na primeira vez, ou abre pelo link com `#k=CHAVE`.

## Configuração
1. Trocar `API_URL` no `index.html` pela URL do app da web do Apps Script.
2. Trocar `TOKEN` no `Code.gs` pela chave escolhida (só dentro da planilha, nunca no GitHub).
