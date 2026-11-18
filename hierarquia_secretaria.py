# -*- coding: utf-8 -*-
"""Hierarquia SECRETARIA -> EXECUTOR para a base de Ações em Monitoramento
("Ações do Cronograma.xlsx"), espelhando a hierarquia SECRETARIA_LIMPA ->
EXECUTOR já existente na base do PAC (df_original) — pedido do usuário em
18/11/2026.

A base do Monitoramento não tem coluna de secretaria própria; só "Executor"
(CONDER, SEDUR, CTB/SEDUR, CONV/PREF/SUDESB...) e "Eixo" (taxonomia diferente
da do PAC). Esse módulo deriva a Secretaria de cada linha a partir de:
1. Um de-para EDITÁVEL ("DE-PARA SECRETARIA EXECUTOR.xlsx", na pasta do
   programa — ver caminho_de_para()), que o usuário revisa/ajusta.
2. Quando o de-para não decide, suposições tiradas da própria base do PAC
   (mapa_do_pac) e do fato de alguns executores serem, eles mesmos, uma das
   secretarias (SEC, SECULT, SEDUR, SEINFRA, SESAB...).
Ver secretaria_de() para a ordem de precedência exata.

Ligado ao módulo principal pelo mesmo padrão dos outros satélites
(resumo_acoes.py, gerar_pptx_pac.py): configurar(sys.modules[__name__]).
"""
import math
import os
import re
import unicodedata

import pandas as pd

pac = None


def configurar(modulo):
    global pac
    pac = modulo


NAO_IDENTIFICADA = "NÃO IDENTIFICADA"
ARQUIVO_DE_PARA = "DE-PARA SECRETARIA EXECUTOR.xlsx"
COLUNAS_DE_PARA = ["EXECUTOR", "EIXO", "SECRETARIA", "OBSERVAÇÃO"]


def normalizar_executor(texto):
    """Maiúsculas, sem acento, espaços colapsados — usada tanto pra Executor
    quanto pra Eixo (mesma normalização; comparação insensível a caixa/acento/
    espaçamento, que é toda a variação que aparece nas duas bases)."""
    if texto is None or (isinstance(texto, float) and math.isnan(texto)):
        return ""
    s = unicodedata.normalize("NFKD", str(texto))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().upper()


def _segmentos(executor):
    """O executor inteiro, seguido de cada pedaço de um composto tipo
    "CONV/PREF/SUDESB" ou "CTB/SEDUR" — pra tentar cada um como fallback
    quando o composto inteiro não bate com nada."""
    texto = str(executor or "")
    vistos, lista = set(), []
    for cand in [texto] + (texto.split("/") if "/" in texto else []):
        n = normalizar_executor(cand)
        if n and n not in vistos:
            vistos.add(n)
            lista.append(n)
    return lista


# ---- suposições tiradas da base do PAC -------------------------------------
_pac_cache = None  # (id(pac.df), mapa_unico, secretarias_conhecidas)


def mapa_do_pac():
    """executor normalizado -> secretaria, só quando esse executor mapeia pra
    UMA ÚNICA secretaria em toda a base do PAC (df_original), independente do
    eixo — ex.: CONDER aparece com SEDUR/SESAB/SECULT/SEADES conforme o eixo,
    então fica fora daqui (ambíguo); SEC, SEINFRA, CERB etc. têm uma só."""
    global _pac_cache
    if pac is None or not hasattr(pac, "df"):
        return {}
    df = pac.df
    if _pac_cache and _pac_cache[0] == id(df):
        return _pac_cache[1]
    pares = df[["EXECUTOR", "SECRETARIA_LIMPA"]].dropna().drop_duplicates()
    por_executor = {}
    for _, r in pares.iterrows():
        sec = str(r["SECRETARIA_LIMPA"]).strip()
        if not sec:
            continue
        # indexa o composto inteiro E cada pedaço ("SECULT/IPAC" -> também
        # "SECULT" e "IPAC" sozinhos), pra um executor composto de um lado
        # achar o outro mesmo quando só um dos dois lados é composto.
        for seg in _segmentos(r["EXECUTOR"]):
            por_executor.setdefault(seg, set()).add(sec)
    mapa_unico = {k: next(iter(v)) for k, v in por_executor.items() if len(v) == 1}
    secretarias_conhecidas = {
        normalizar_executor(s): str(s).strip()
        for s in df["SECRETARIA_LIMPA"].dropna().astype(str).unique() if str(s).strip()
    }
    _pac_cache = (id(df), mapa_unico, secretarias_conhecidas)
    return mapa_unico


def _secretarias_conhecidas():
    mapa_do_pac()  # garante o cache preenchido
    return _pac_cache[2] if _pac_cache else {}


# ---- de-para editável --------------------------------------------------------
def caminho_de_para():
    return pac.caminho_recurso(ARQUIVO_DE_PARA) if pac else ARQUIVO_DE_PARA


_de_para_cache = None  # (caminho, mtime, exato, exec_so)


def _carregar_de_para():
    caminho = caminho_de_para()
    if not os.path.exists(caminho):
        return {}, {}
    global _de_para_cache
    mtime = os.path.getmtime(caminho)
    if _de_para_cache and _de_para_cache[:2] == (caminho, mtime):
        return _de_para_cache[2], _de_para_cache[3]
    df = pd.read_excel(caminho, dtype=object)
    exato, exec_so = {}, {}
    for _, r in df.iterrows():
        ex_n = normalizar_executor(r.get("EXECUTOR"))
        sec = str(r.get("SECRETARIA") or "").strip()
        if not ex_n or not sec or sec.upper() == NAO_IDENTIFICADA:
            continue
        eixo_bruto = r.get("EIXO")
        eixo_n = normalizar_executor(eixo_bruto) if pd.notna(eixo_bruto) and str(eixo_bruto).strip() else ""
        if eixo_n:
            exato[(ex_n, eixo_n)] = sec
        else:
            exec_so[ex_n] = sec
    _de_para_cache = (caminho, mtime, exato, exec_so)
    return exato, exec_so


def _sugestao_pac(executor):
    """Sugestão derivada só da base do PAC (sem olhar o de-para) — usada pra
    pré-preencher linhas novas do de-para, nunca pra decidir em tempo real
    (secretaria_de olha o de-para primeiro)."""
    mapa_unico = mapa_do_pac()
    conhecidas = _secretarias_conhecidas()
    for seg in _segmentos(executor):
        if seg in mapa_unico:
            return mapa_unico[seg]
    for seg in _segmentos(executor):
        if seg in conhecidas:
            return conhecidas[seg]
    return ""


def secretaria_de(executor, eixo):
    """Secretaria derivada de um (Executor, Eixo) da base de Monitoramento.
    Ordem de precedência: de-para exato (executor+eixo) -> de-para só-executor
    -> secretaria única do executor no PAC -> o próprio executor já é uma
    secretaria conhecida -> "NÃO IDENTIFICADA". Em cada etapa, se o executor
    for composto ("CONV/PREF/SUDESB"), tenta o composto inteiro e depois cada
    pedaço separado."""
    exato, exec_so = _carregar_de_para()
    mapa_unico = mapa_do_pac()
    conhecidas = _secretarias_conhecidas()
    segmentos = _segmentos(executor)
    eixo_n = normalizar_executor(eixo)
    for seg in segmentos:
        if (seg, eixo_n) in exato:
            return exato[(seg, eixo_n)]
    for seg in segmentos:
        if seg in exec_so:
            return exec_so[seg]
    for seg in segmentos:
        if seg in mapa_unico:
            return mapa_unico[seg]
    for seg in segmentos:
        if seg in conhecidas:
            return conhecidas[seg]
    return NAO_IDENTIFICADA


def secretaria_de_serie(serie_executor, serie_eixo):
    """Versão vetorizada de secretaria_de(), para uma coluna inteira."""
    pares = list(zip(serie_executor.fillna(""), serie_eixo.fillna("")))
    return pd.Series([secretaria_de(ex, ei) for ex, ei in pares], index=serie_executor.index)


def garantir_de_para_e_avisar(df, col_executor, col_eixo):
    """Garante que o de-para exista e tenha uma linha pra cada par
    (Executor, Eixo) que aparece em df — cria o arquivo na primeira vez
    (com sugestões tiradas do PAC), ou só ACRESCENTA pares novos a um arquivo
    já existente, sem tocar em linhas que o usuário já revisou. Imprime um
    resumo (total de pares, quantos vieram com sugestão x em branco, quantos
    são novos nesta execução) — pedido do usuário: mostrar essa contagem e
    parar pra ele revisar antes de qualquer outra mudança."""
    pares_base = (
        df[[col_executor, col_eixo]]
        .dropna(subset=[col_executor])
        .drop_duplicates()
    )
    caminho = caminho_de_para()
    if os.path.exists(caminho):
        existente = pd.read_excel(caminho, dtype=object)
    else:
        existente = pd.DataFrame(columns=COLUNAS_DE_PARA)
    chaves_existentes = {
        (normalizar_executor(r.get("EXECUTOR")), normalizar_executor(r.get("EIXO")))
        for _, r in existente.iterrows()
    }
    linhas_novas = []
    for _, r in pares_base.iterrows():
        executor, eixo = r[col_executor], r.get(col_eixo)
        chave = (normalizar_executor(executor), normalizar_executor(eixo))
        if chave in chaves_existentes:
            continue
        chaves_existentes.add(chave)
        linhas_novas.append({
            "EXECUTOR": str(executor).strip(),
            "EIXO": str(eixo).strip() if pd.notna(eixo) else "",
            "SECRETARIA": _sugestao_pac(executor),
            "OBSERVAÇÃO": "",
        })
    final = (pd.concat([existente, pd.DataFrame(linhas_novas)], ignore_index=True)
             if linhas_novas else existente)
    if linhas_novas:
        final[COLUNAS_DE_PARA].to_excel(caminho, index=False)
        global _de_para_cache
        _de_para_cache = None

    total = len(final)
    sugeridas = int(final["SECRETARIA"].fillna("").astype(str).str.strip().ne("").sum())
    em_branco = total - sugeridas
    print(
        f"[DE-PARA SECRETARIA/EXECUTOR] {caminho}\n"
        f"  pares (Executor, Eixo) no total: {total}\n"
        f"  novos nesta execução: {len(linhas_novas)}\n"
        f"  com sugestão de secretaria: {sugeridas}\n"
        f"  em branco (precisam de revisão manual): {em_branco}"
    )
    return {"total": total, "novos": len(linhas_novas), "sugeridas": sugeridas, "em_branco": em_branco}
