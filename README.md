# Controle da Qualidade — Análise de Risco (Reunião de Antecipação)

Aplicação React (Vite) para consolidar a **Sequência de Produção** com a
**Disponibilidade de Instrumentos** em uma tabela resumo, com exportação para
impressão/PDF em formato **A4 paisagem**.

## Como funciona

1. Importe a planilha **Sequência de Produção** (`.xlsx`, `.xls`, `.csv`).
2. Importe a planilha **Disponibilidade de Instrumentos**.
3. A tabela resumo é gerada automaticamente, correlacionando a coluna **HALB**
   (Sequência de Produção) com a coluna **HALBQ** (Disponibilidade de Instrumentos).
4. Linhas com algum instrumento **"Indisponível"** ou **em branco** são destacadas
   em **vermelho claro**.
5. Botão **Gerar impressão / PDF** (canto superior direito) abre o diálogo de
   impressão já configurado para A4 paisagem — basta escolher "Salvar como PDF".

### Colunas exibidas

**Sequência de Produção:** Pedido/item · Ordem · Halb · Início (Enfornam.) ·
Cliente externo · Teste EMI · Descrição produto

**Disponibilidade de Instrumentos:** Tubo Padrão UT · Tubo Padrão EMI · Drift ·
Sapata UT · Qualificação ABENDI (para NDT)

> O reconhecimento das colunas é tolerante a variações de acento, maiúsculas e
> espaços, e detecta automaticamente a linha de cabeçalho mesmo que existam
> linhas de título acima dela.

## Rodar localmente

```bash
npm install
npm run dev
```

## Publicar na Vercel

**Opção A — pelo painel:**
1. Suba esta pasta para um repositório (GitHub/GitLab/Bitbucket).
2. Em vercel.com → **Add New → Project** → importe o repositório.
3. Framework detectado: **Vite**. Build: `npm run build` · Output: `dist`.
4. **Deploy**.

**Opção B — pela CLI:**
```bash
npm i -g vercel
vercel
```
