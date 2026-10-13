# Simulador IBS e CBS — 2027

Aplicação Streamlit que lê os documentos da pasta `ARQUIVOS`, projeta os
débitos de IBS/CBS dos serviços médicos, recalcula as retenções federais e
estima os créditos das contas a pagar.

## Executar

No Windows, abra `executar.bat` ou use:

```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Premissas padrão

- IBS: 0,10%;
- CBS: 8,50%;
- redução dos serviços de saúde: 60%;
- IR: 1,50% e CSLL: 1,00%, somente quando já destacados no arquivo original;
- PIS/Cofins: retirados da projeção de retenções de 2027;
- fornecedores do Simples: primeira faixa do Anexo III por padrão, com opção
  de trocar o anexo no painel;
- lançamentos sem conciliação ou sem informação Sim/Não: sem crédito, por
  prudência (há uma opção no painel para simular esses casos como regulares).

Os parâmetros são editáveis no painel lateral. O Excel gerado pelo aplicativo
leva a memória de cálculo, os detalhes dos documentos e a tabela dos Anexos I
a V do Simples Nacional para 2027/2028.

## Observação sobre os arquivos XLS

Os relatórios de contas a pagar possuem um índice interno de planilha
inconsistente. O leitor corrige o índice apenas em memória. Os arquivos
originais não são alterados.

## Fontes legais principais

- [LC 214/2025 compilada](https://www.planalto.gov.br/ccivil_03/leis/lcp/lcp214compilado.htm),
  especialmente os arts. 128 e 130 (redução para saúde);
- [LC 123/2006 atualizada](https://legis.senado.gov.br/norma/40180341), com as
  alterações da LC 227/2026 (crédito do Simples e tabelas de 2027/2028);
- [orientações da Receita Federal](https://www.gov.br/receitafederal/pt-br/acesso-a-informacao/acoes-e-programas/programas-e-atividades/reforma-tributaria-do-consumo/entenda)
  para a transição de 2027.

O resultado é uma simulação gerencial e não substitui a validação fiscal de
cada documento e do enquadramento efetivo do fornecedor.
