// Baixa o histórico de preços diários da EPAGRI/Cepa, extrai o último preço
// ("Val. Mais Comum") de Soja, Milho e Trigo na praça EXTREMO OESTE e grava precos.json.

import * as XLSX from 'xlsx';
import { writeFileSync } from 'node:fs';

const URL_XLSX = 'https://docweb.epagri.sc.gov.br/website_cepa/precos/Historico_precos_diario.xlsx';
const PRACA_ALVO = 'extremo oeste';
const PRODUTOS = { soja: 'soja', milho: 'milho', trigo: 'trigo' };

const MESES = {
  janeiro:1, fevereiro:2, marco:3, abril:4, maio:5, junho:6,
  julho:7, agosto:8, setembro:9, outubro:10, novembro:11, dezembro:12
};

const norm = s => String(s ?? '')
  .normalize('NFD').replace(/\p{Diacritic}/gu, '')
  .trim().toLowerCase();

const fmtValor = v => Number(v).toFixed(2).replace('.', ',');
const fmtData  = (a,m,d) => `${String(d).padStart(2,'0')}/${String(m).padStart(2,'0')}/${a}`;

async function main(){
  console.log('Baixando xlsx da EPAGRI...');
  const resp = await fetch(URL_XLSX);
  if(!resp.ok) throw new Error('Falha ao baixar o xlsx: HTTP ' + resp.status);
  const buf = Buffer.from(await resp.arrayBuffer());

  const wb = XLSX.read(buf, { type:'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const linhas = XLSX.utils.sheet_to_json(ws, { header:1, blankrows:false });

  const hIdx = linhas.findIndex(r => r.some(c=>norm(c)==='produto') && r.some(c=>norm(c)==='nom_praca'));
  if(hIdx < 0) throw new Error('Cabecalho (Produto / nom_praca) nao encontrado.');
  const head = linhas[hIdx].map(norm);

  const col = alvo => head.findIndex(h => h === alvo);
  const cAno    = 0;
  const cMes    = col('mes preco');
  const cDia    = col('dia');
  const cProd   = col('produto');
  const cPraca  = col('nom_praca');
  const cComum  = col('val. mais comum');
  if([cMes,cDia,cProd,cPraca,cComum].some(i=>i<0))
    throw new Error('Alguma coluna esperada nao foi encontrada no cabecalho.');

  const melhor = {};
  for(let i=hIdx+1; i<linhas.length; i++){
    const r = linhas[i];
    if(norm(r[cPraca]) !== PRACA_ALVO) continue;
    const prod = norm(r[cProd]);
    let chave = null;
    for(const [k,pref] of Object.entries(PRODUTOS)) if(prod.startsWith(pref)) chave = k;
    if(!chave) continue;
    const a = Number(r[cAno]);
    const m = MESES[norm(r[cMes])];
    const d = Number(r[cDia]);
    const comum = r[cComum];
    if(!a || !m || !d || comum==null || comum==='') continue;
    const rank = a*10000 + m*100 + d;
    if(!melhor[chave] || rank > melhor[chave].rank){
      melhor[chave] = { rank, valor: comum, a, m, d };
    }
  }

  if(!melhor.soja || !melhor.milho)
    throw new Error('Nao encontrei preco de soja e/ou milho para Extremo Oeste.');

  const bloco = m => ({ valor: fmtValor(m.valor), data: fmtData(m.a, m.m, m.d) });
  const out = {
    soja:  bloco(melhor.soja),
    milho: bloco(melhor.milho),
    fonte: 'EPAGRI/Cepa - praca Extremo Oeste - Val. Mais Comum',
    atualizado_em: new Date().toISOString().slice(0,10)
  };
  if(melhor.trigo) out.trigo = bloco(melhor.trigo);

  writeFileSync('precos.json', JSON.stringify(out, null, 2) + '\n', 'utf8');
  console.log('precos.json gravado:', JSON.stringify(out));
}

main().catch(e => { console.error('ERRO:', e.message); process.exit(1); });
