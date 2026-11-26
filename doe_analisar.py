"""Monitoramento do Diário Oficial do Estado (DOE) para o CGAPE — pedido do
usuário em 26/11/2026, a partir do recurso DOE do projeto RELATÓRIO SSP
(DOE\relatorios\ferramenta\analisar.py), com duas mudanças:

  - sem o alerta de exoneração de uma pessoa específica;
  - licitações e contratos de TODAS as secretarias (antes só da Segurança
    Pública), e cada ato/licitação marcado com a SECRETARIA da seção do
    diário em que saiu — é o que alimenta a caixa de seleção múltipla de
    secretarias do painel do DOE.

Uso como módulo: doe_analisar.executar(pasta_md, arquivo_json[, manual_json])
lê os doe_*.md da pasta e grava o JSON lido pelo painel (doe_monitor.py).
As funções de leitura do texto são as mesmas do analisar.py original.
"""
import json
import re
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

MARCADOR = re.compile(r"^<#E\.G\.B#[\d#]+/?>\s*$")
MARK = re.compile(r"^#+ .*<mark>(.*?)</mark>")
PLAIN_SEC = re.compile(
    r"^(DECRETOS (NUMERADOS|SIMPLES|FINANCEIROS)|DESPACHOS|GABINETE DO GOVERNADOR|CASA CIVIL|CASA MILITAR.*|"
    r"PROCURADORIA GERAL DO ESTADO|CONTROLADORIA GERAL DO ESTADO|DEFENSORIA P[ÚU]BLICA.*|"
    r"SECRETARIA (D[AEO]S?|DE) [A-ZÀ-Ü ,]+|AVISOS? DE LICITA[ÇC][ÃA]O|RESULTADOS E HOMOLOGA[ÇC][ÕO]ES|"
    r"PODER (LEGISLATIVO|JUDICI[ÁA]RIO)|MINIST[ÉE]RIO P[ÚU]BLICO.*|TRIBUNAL DE CONTAS.*|MUNIC[ÍI]PIOS|"
    r"PUBLICA[ÇC][ÕO]ES DE TERCEIROS.*)\s*$"
)
SEI = re.compile(r"\d{3}\.\d{3,5}\.?\d{4}\.\d{7}-\d{2}")
NOME_UP = r"[A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý'´`.\-]*(?:\s+[A-ZÀ-ÖØ-Ý][A-ZÀ-ÖØ-Ý'´`.\-]*)+"
SIMB = re.compile(r"s[íi]mbolo\s+([A-Z]{2,4}\s?-?\s?[0-9][0-9A-Z]?|[A-Z]{2,4})")
ORGAO_TXT = re.compile(
    r"(Secretaria de Ci[êe]ncia, Tecnologia e Inova[çc][ãa]o|Secretaria do Trabalho, Emprego, Renda e Esporte|"
    r"Secretaria d[aeo]s? [^,.;]+|Secretaria de [^,.;]+|Casa Civil|Casa Militar d[oa] [^,.;]+|Gabinete do Governador|"
    r"Procuradoria Geral do Estado|Controladoria Geral do Estado|Defensoria P[úu]blica[^,.;]*|"
    r"Pol[íi]cia (?:Militar|Civil)[^,.;]*|Corpo de Bombeiros Militar[^,.;]*|Departamento de Pol[íi]cia T[ée]cnica[^,.;]*|"
    r"Universidade [^,.;]+|Assessoria Geral de Comunica[çc][ãa]o Social[^,.;]*|Vice-Governadoria[^,.;]*|"
    r"Funda[çc][ãa]o Estatal Sa[úu]de da Fam[íi]lia)"
)
EMENTA = re.compile(
    r"^(Declara|Altera|Homologa|Disp[õo]e|Aprova|Autoriza|Abre|Institui|Regulamenta|Renova|Renovação|Credencia|"
    r"Recredencia|Transfere|D[áa] |Fixa|Estabelece|Cria|Reconhece|Concede|Prorroga|Delega|Convoca|Revoga|Torna|"
    r"Introduz|Inclui|Promove|Qualifica|Atualiza|Designa|Decreta|Determina|Exclui|Extingue|Modifica|Reorganiza|"
    r"Aprova|Acrescenta|Ratifica|Suspende|Reabre|Define|Prorrogação|Autorização|Credenciamento|Reconhecimento)"
)
INICIO_LIC = re.compile(
    r"(?i)^(aviso|resultado|extrato|resumo|termo|\d+º termo|primeiro|segundo|terceiro|dispensa|inexigib|ata |homolog|"
    r"adjudic|ratific|apostila|chamamento|credenciamento|contrato|conv[êe]nio|preg[ãa]o|concorr[êe]ncia|errata|"
    r"retifica|autoriza[çc][ãa]o de fornecimento|afm|revoga[çc][ãa]o|anula[çc][ãa]o|suspens[ãa]o|adiamento)"
)
SSP_TXT = re.compile(
    r"(SSP|PMBA|PM/BA|PCBA|PC/BA|DPT|CBMBA|CBM/BA|Seguran[çc]a P[úu]blica|SEGURAN[ÇC]A P[ÚU]BLICA|Pol[íi]cia Militar|"
    r"POL[ÍI]CIA MILITAR|Pol[íi]cia Civil|POL[ÍI]CIA CIVIL|Bombeiros|BOMBEIROS|Pol[íi]cia T[ée]cnica|POL[ÍI]CIA T[ée]CNICA)"
)


def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


def limpa(s):
    s = re.sub(r"<br\s*/?>", " ", s)
    s = re.sub(r"\*\*|<[^>]+>|~~|^#+\s*", "", s)
    s = s.replace("_", " ") if s.strip().startswith("_") else s
    return re.sub(r"[ \t]+", " ", s).strip()


def junta(linhas):
    """Junta linhas quebradas (texto corrido), desfazendo hifenização."""
    out = ""
    for l in linhas:
        l = l.strip()
        if not l:
            continue
        if out.endswith("-") and l[:1].islower():
            out = out[:-1] + l
        else:
            out = (out + " " + l) if out else l
    return limpa(out)


# ---------------------------------------------------------------- leitura


class Diario:
    def __init__(self, arq):
        self.arq = arq
        self.linhas = arq.read_text(encoding="utf-8").splitlines()
        self.corrido = sum(1 for l in self.linhas if MARCADOR.match(l.strip())) > 50
        self.blocos = self._blocos()  # (linha_ini, linha_fim, secao, [linhas])

    def _secao(self, l, atual):
        m = MARK.match(l)
        if m:
            nome = limpa(m.group(1))
            return nome if nome == nome.upper() else atual
        if PLAIN_SEC.match(l.strip()) and l.strip() == l.strip().upper():
            return l.strip()
        return atual

    def _blocos(self):
        res, cur, ini, sec = [], [], 0, "(início)"
        for i, l in enumerate(self.linhas):
            novo = self._secao(l, sec)
            if self.corrido:
                quebra = MARCADOR.match(l.strip()) or novo != sec or (PLAIN_SEC.match(l.strip()) and l.strip() == l.strip().upper())
                if quebra:
                    if cur:
                        res.append((ini, i, sec_bloco, cur))
                    cur = []
                    sec = novo
                    continue
                if not l.strip() or l.startswith("## Página") or re.match(r"^[<#E.G0-9/>]{1,3}$", l.strip()):
                    continue
                if not cur:
                    ini, sec_bloco = i, sec
                cur.append(l)
            else:
                if novo != sec:
                    if cur:
                        res.append((ini, i, sec_bloco, cur))
                    cur, sec = [], novo
                    continue
                if l.strip():
                    if not cur:
                        ini, sec_bloco = i, sec
                    cur.append(l)
                elif cur:
                    res.append((ini, i, sec_bloco, cur))
                    cur = []
        if cur:
            res.append((ini, len(self.linhas), sec_bloco, cur))
        return res


# ---------------------------------------------------------------- pessoal


def orgao_portaria(txt):
    m = re.search(r"\bdo\s?\(a\)\s+([A-ZÀ-Ü].+?)\s+-\s+([A-ZÀ-Ü][A-ZÀ-Ü/ ]{1,15}?)\s*,", txt) or \
        re.search(r"\bd[oa]\s+([A-ZÀ-Ü][^,]+?)\s+-\s+([A-ZÀ-Ü][A-ZÀ-Ü/ ]{1,15}?)\s*,", txt)
    if m:
        return f"{m.group(1).strip()} - {m.group(2).strip()}"
    m = re.search(r"(?:O|A) (?:Reitor|Diretor Geral|Presidente|Secret[áa]ri[oa])[^,]*? d[aeo] ([^,]+?),", txt)
    return m.group(1).strip() if m else ""


AUTORIDADE = re.compile(
    r"(?i:diretor[a]?[- ]geral|presidente|reitor[a]?|secret[áa]ri[oa])\s+(?i:d[aeo])\s+"
    r"([A-ZÀ-Ü][A-ZÀ-Ü ]{5,}?)\s*(?:[-–]\s*|\()([A-Z]{2,10})\)?")


def orgao_autoridade(txt):
    m = AUTORIDADE.search(txt)
    return f"{m.group(1).strip()} - {m.group(2)}" if m else ""


def orgao_texto(seg):
    achados = ORGAO_TXT.findall(seg)
    return achados[-1].strip() if achados else ""


def nome_em(seg):
    for m in re.finditer(NOME_UP, seg):
        nome = m.group(0).strip(" ,.-")
        if not re.fullmatch(r"[X.\s-]+", nome):  # ignora CPF mascarado (XXX.XXX)
            return nome
    return ""


def atos_frase(txt):
    """Atos de nomeação/exoneração escritos em frase (decretos e portarias individuais)."""
    atos = []
    verbos = list(re.finditer(r"(?i)\b(considerar exonerad[oa]s?|exonerar|nomear)\b", txt))
    for k, v in enumerate(verbos):
        fim = verbos[k + 1].start() if k + 1 < len(verbos) else len(txt)
        seg = txt[v.start():fim]
        verbo = v.group(1).lower()
        tipo = "Nomeação" if verbo == "nomear" else "Exoneração"
        mc = re.search(r"(?i)(?:para (?:exercer )?o|do|da) (?:cargo(?:/fun[çc][ãa]o)?|fun[çc][ãa]o estrat[ée]gica de gest[ãa]o) "
                       r"(?:em comiss[ãa]o )?(?:tempor[áa]rio )?(?:de )?([^,;]+)", seg)
        if not mc:
            continue  # ex.: nomear membros de conselho (sem cargo)
        antes = seg[: mc.start()]
        nome = nome_em(re.sub(r"(?i)^(considerar exonerad[oa]s?|exonerar|nomear)", "", antes))
        if not nome:  # nomes em caixa mista (ex.: atos da FESF)
            mm = re.search(r"(?:o|a)\s+(?:profissional|empregad[oa](?: p[úu]blic[oa])?|servidor(?:a)?)\s+([A-ZÀ-Ý][\wÀ-ÿ']+(?:\s+(?:d[aeo]s?|e|[A-ZÀ-Ý][\wÀ-ÿ']+))+)", antes)
            nome = mm.group(1).strip() if mm else ""
        if not nome or len(nome) < 6:
            continue
        cargo = re.split(r"\s+s[íi]mbolo\b", mc.group(1).strip().rstrip("."))[0]
        ms = SIMB.search(seg)
        simbolo = re.sub(r"\s", "", ms.group(1)) if ms else "—"
        resto = seg[ms.end():] if ms else seg[mc.end():]
        resto = re.split(r"(?i)(?:\.\s|\.$|, a partir de|, com efeito|, em cumprimento|, nos termos|PALÁCIO)", resto)[0]
        resto = resto.strip(" ,;")
        orgao = orgao_texto(seg) or ""
        if re.search(r"Funda[çc][ãa]o Estatal Sa[úu]de da Fam[íi]lia", txt):
            orgao = "Fundação Estatal Saúde da Família - FESF-SUS"
        lot = resto
        if orgao and orgao in lot:
            lot = lot[: lot.rfind(orgao)].strip(" ,;")
        lot = re.sub(r"^(?:d[oa]\(a\)|d[oa]s?)\s+", "", lot).strip(" ,;")
        lot = re.sub(r",?\s+d[oa]s?$", "", lot).strip(" ,;")
        if lot.lower().startswith("matrícula") or lot.lower().startswith("cadastro"):
            lm = re.search(r"lotad[oa] (?:n[oa] |em )?([^,.;]+)", seg)
            lot = lm.group(1).strip() if lm else ""
        atos.append(dict(nome=nome, tipo=tipo, cargo=cargo, simbolo=simbolo, lotacao=lot[:160], orgao=orgao))
    return atos


def celulas_md(linha):
    l = linha.strip()
    l = l[1:] if l.startswith("|") else l
    l = l[:-1] if l.endswith("|") else l
    return [limpa(p) for p in l.split("|")]


def tabela_md(ls):
    """Linhas markdown de uma tabela -> (cabeçalho, linhas)."""
    rows = [celulas_md(l) for l in ls if not re.match(r"^\|[-|\s:]+\|?$", l.strip())]
    return (rows[0], rows[1:]) if rows else ([], [])


def tabela_corrida(ls, n):
    """Tabela do texto corrido: células começam com espaço; o resto é continuação."""
    cel = []
    for l in ls:
        if l.startswith(" ") and l.strip():
            cel.append(l.strip())
        elif cel and l.strip():
            sep = "" if cel[-1].endswith("-") and l[:1].islower() else " "
            cel[-1] = (cel[-1][:-1] if sep == "" else cel[-1]) + sep + l.strip()
    return [cel[i:i + n] for i in range(0, len(cel), n)]


def idx(cab, *nomes):
    for i, c in enumerate(cab):
        cn = sem_acento(c).lower()
        if any(cn.startswith(sem_acento(n).lower()) for n in nomes):
            return i
    return None


def linhas_tabela(cab, rows):
    """Funde linhas de continuação e devolve dicts com nome/cargo/simbolo/lotacao."""
    i_nome, i_cargo = idx(cab, "Nome"), idx(cab, "Cargo")
    i_simb, i_mat = idx(cab, "Símbolo"), idx(cab, "Matrícula")
    i_lot = idx(cab, "Lotação", "Unidade")
    out = []
    for r in rows:
        r = r + [""] * (len(cab) - len(r))
        if i_nome is None or i_nome >= len(r):
            continue
        novo = (i_mat is not None and r[i_mat].strip()) or (i_mat is None and i_simb is not None and r[i_simb].strip()) \
            or (i_mat is None and i_simb is None and i_cargo is not None and r[i_cargo].strip())
        if not out or novo:
            out.append([c for c in r])
        else:
            prev = out[-1]
            for j, c in enumerate(r):
                if c.strip() and j < len(prev):
                    prev[j] = (prev[j] + " " + c).strip()
    res = []
    for r in out:
        g = lambda i: re.sub(r"([a-zà-ü])- ([a-zà-ü])", r"\1\2", re.sub(r"\s+", " ", r[i])).strip() \
            if i is not None and i < len(r) else ""
        nome, cargo, simb, lot = g(i_nome), g(i_cargo), g(i_simb), g(i_lot)
        if not nome or nome.lower() == "nome":
            continue
        # célula mesclada na origem ("NOME Cargo" numa célula só): desloca as colunas
        mc = re.match(r"^([A-ZÀ-Ü' .-]+?)\s+((?:Coordenador|Diretor|Secret[áa]rio|Assessor|Subgerente|Gerente|Chefe|"
                      r"Supervisor|Assistente|Vice-Diretor)\b.*)$", nome)
        if mc and i_simb is not None and re.fullmatch(r"[A-Z]{2,4}-?\w*", cargo or ""):
            nome, lot, simb, cargo = mc.group(1), simb, cargo, mc.group(2)
        lot = re.sub(r"\s*\d{1,2}\s*de\s*\w+(\s*de)?\s*(\d{4})?\s*$", "", lot)  # data de início colada
        lot = re.sub(r"([A-ZÀ-Ü]{3,})- ([A-ZÀ-Ü]{2,})", r"\1\2", lot)
        cargo = re.sub(r"([A-ZÀ-Üa-zà-ü]{3,})- ([a-zà-ü]{2,})", r"\1\2", cargo)
        res.append(dict(nome=nome, cargo=cargo or "—", simbolo=simb.replace(" ", "") or "—", lotacao=lot))
    return res


def ultimo_resolve(diario, linha, janela=60):
    """Verbo do 'resolve ...' mais próximo antes da linha e o órgão emissor
    (este só quando o parágrafo do 'resolve' está colado à tabela)."""
    for i in range(linha - 1, max(0, linha - janela) - 1, -1):
        t = limpa(diario.linhas[i])
        m = list(re.finditer(r"resolve\s+(nomear|exonerar|\w+)", t))
        if m:
            org = orgao_portaria(t) if linha - i <= 6 else ""
            return m[-1].group(1), org
    return None, ""


def assinatura_depois(diario, linha, janela=14):
    """Órgão escrito na linha abaixo do nome de quem assina, logo após a tabela."""
    ls = [limpa(l) for l in diario.linhas[linha: linha + janela] if l.strip()]
    for j, t in enumerate(ls[:-1]):
        if t.startswith("|") or t.startswith("Portaria") or t.startswith("O(A)"):
            break
        if re.fullmatch(r"[A-ZÀ-Ü][A-ZÀ-Ü .'-]{5,}", t):
            org = ls[j + 1]
            if re.fullmatch(r"[A-ZÀ-Ü0-9][A-ZÀ-Ü0-9 .,'/–-]{4,}", org) and not org.startswith("CÓPIA"):
                return org
    return ""


def tipo_tabela(cabn):
    if "inscricao" in cabn and any(c.startswith("classifica") for c in cabn) \
            and any(c.startswith("pontuacao final") or c.startswith("documento") for c in cabn):
        return "concurso"
    if any(c.startswith("nome completo") for c in cabn) and "acao" in cabn:
        return "judicial"
    tem_simb = any(c.startswith("simbolo") for c in cabn)
    if tem_simb and cabn[:1] == ["nome"]:
        return "nom_comissao"
    if tem_simb and cabn[:1] == ["matricula"]:
        return "exo_comissao"
    if cabn[:2] == ["matricula", "nome"] and "processo" in cabn and any(c.startswith("unidade") for c in cabn) \
            and not any(c.startswith("unidade origem") for c in cabn):
        return "efetivo"
    return None


def tipo_cabecalho(t):
    m = re.search(r"resolve\s*:?\s+(\w+)", t)
    if not m:
        return None
    v = m.group(1).lower()
    if v == "nomear":
        if "candidato" in t:
            return "judicial" if "provis" in t else "concurso"
        return "nom_comissao"
    if v == "exonerar":
        return "efetivo" if re.search(r"art\.\s*46", t) else "exo_comissao"
    return "outro:" + v


COMPATIVEL = {
    "concurso": {"concurso"}, "judicial": {"judicial"}, "nom_comissao": {"nom_comissao"},
    "exo_comissao": {"exo_comissao"}, "efetivo": {"efetivo", "outro"},
}


def tabelas_pessoal_md(d):
    """Pareia tabelas de pessoal com o cabeçalho 'resolve ...' compatível mais próximo."""
    cabs, tabs = [], []
    for ini, fim, sec, ls in d.blocos:
        if sec.startswith("PUBLICA") or sec.startswith("MUNIC"):
            continue
        if ls[0].lstrip().startswith("|"):
            cab, rows = tabela_md(ls)
            k = tipo_tabela([sem_acento(c).lower() for c in cab])
            if k:
                tabs.append(dict(ini=ini, fim=fim, cab=cab, rows=rows, tipo=k))
            continue
        for i, l in enumerate(ls):
            t = limpa(l)
            k = tipo_cabecalho(t)
            if k and k.startswith("outro") and not re.search(r"(abaixo|servidor\(es\)|:\s*$)", t):
                continue  # ato individual, sem tabela
            if k and not (k in ("nom_comissao", "exo_comissao") and atos_frase(t)):
                org = orgao_portaria(t)
                if not org:  # cabeçalho partido pela quebra de página: o "O(A) ..." ficou acima
                    for l2 in reversed(d.linhas[max(0, ini + i - 8): ini + i]):
                        if limpa(l2).startswith("O(A)"):
                            org = orgao_portaria(limpa(l2))
                            break
                cabs.append(dict(linha=ini + i, txt=t, tipo=k, org=org, usado=False))
    atos, pend = [], []
    for tb in tabs:
        aceitos = COMPATIVEL[tb["tipo"]]
        cand = [c for c in cabs if not c["usado"] and c["tipo"].split(":")[0] in aceitos
                and -40 <= tb["ini"] - c["linha"] <= 70]
        cand.sort(key=lambda c: (c["linha"] > tb["ini"], abs(tb["ini"] - c["linha"])))
        cab_ok = cand[0] if cand else None
        if cab_ok:
            cab_ok["usado"] = True
            if cab_ok["tipo"].startswith("outro"):
                continue  # tabela de abono, averbação etc.
        elif tb["tipo"] == "efetivo":
            continue  # sem 'resolve exonerar' por perto: não dá para saber o que é
        org = (cab_ok["org"] if cab_ok else "") or assinatura_depois(d, tb["fim"])
        cabn = [sem_acento(c).lower() for c in tb["cab"]]
        if tb["tipo"] in ("concurso", "judicial"):
            acima = [limpa(l) for l in d.linhas[max(0, tb["ini"] - 12):tb["ini"]]]
            cargo_l = next((l.split(":", 1)[1].strip() for l in reversed(acima) if l.startswith("Cargo:")), "—")
            fator_l = next((l.split(":", 1)[1].strip() for l in reversed(acima) if l.startswith("Fator:")), "")
            if "Fator:" in cargo_l:  # "Cargo: X Fator: Y" na mesma linha
                cargo_l, fator_l = (p.strip() for p in cargo_l.split("Fator:", 1))
            i_nome = next((i for i, c in enumerate(cabn) if c.startswith("nome") or c.startswith("candidat")), None)
            if i_nome is None:
                continue
            i_cargo = idx(tb["cab"], "Cargo")
            i_fator = idx(tb["cab"], "Fator")
            for r in tb["rows"]:
                if len(r) <= i_nome or not r[i_nome].strip():
                    continue
                cargo = r[i_cargo] if i_cargo is not None and i_cargo < len(r) else cargo_l
                fator = r[i_fator] if i_fator is not None and i_fator < len(r) else fator_l
                simb = "— (efetivo, concurso)" if tb["tipo"] == "concurso" else "— (provisório, decisão judicial)"
                atos.append(dict(nome=r[i_nome].strip(), tipo="Nomeação", cargo=cargo, simbolo=simb,
                                 lotacao=fator, orgao=org, linha=tb["ini"] + 1))
            continue
        tipo = "Nomeação" if tb["tipo"] == "nom_comissao" else "Exoneração"
        for r in linhas_tabela(tb["cab"], tb["rows"]):
            atos.append(dict(r, tipo=tipo, orgao=org, linha=tb["ini"] + 1))
    for c in cabs:
        if not c["usado"] and not c["tipo"].startswith("outro"):
            depois = " / ".join(limpa(l) for l in d.linhas[c["linha"] + 1: c["linha"] + 14] if l.strip())
            pend.append(f"L{c['linha'] + 1} (tabela não encontrada) {c['txt'][:500]} >>> {depois[:600]}")
    return atos, pend


def pessoal(d):
    atos, pend = [], []
    cabecalhos, tabelas_lidas = [], set()
    if not d.corrido:
        atos, pend = tabelas_pessoal_md(d)
    anteriores = []
    for ini, fim, sec, ls in d.blocos:
        txt = junta(ls) if d.corrido else limpa(" ".join(l.strip() for l in ls))
        anteriores = (anteriores + [txt])[-4:]
        if (sec.startswith("PUBLICA") or sec.startswith("MUNIC")) and "Estatal Saúde da Família" not in txt:
            continue
        if not d.corrido and (ls[0].lstrip().startswith("|") or re.search(r"resolve\s*:?\s+(nomear|exonerar)", txt)
                              and not atos_frase(txt)):
            continue  # já tratado em tabelas_pessoal_md
        # ---- tabelas markdown (antigo; desativado)
        if False:
            cab, rows = tabela_md(ls)
            cabn = [sem_acento(c).lower() for c in cab]
            if "nome" not in cabn:
                continue
            verbo, org = ultimo_resolve(d, ini)
            if "inscricao" in cabn and any(c.startswith("classifica") for c in cabn):
                # nomeação de aprovados em concurso (ex.: UNEB): cargo e fator vêm nas linhas acima
                if verbo != "nomear":
                    continue
                acima = [limpa(l) for l in d.linhas[max(0, ini - 12):ini]]
                cargo = next((l.split(":", 1)[1].strip() for l in reversed(acima) if l.startswith("Cargo:")), "—")
                fator = next((l.split(":", 1)[1].strip() for l in reversed(acima) if l.startswith("Fator:")), "")
                org_conc = next((orgao_portaria(limpa(l)) for l in reversed(d.linhas[max(0, ini - 40):ini])
                                 if re.search(r"resolve\s+nomear", limpa(l))), "")
                org = org_conc or assinatura_depois(d, fim) or org
                for r in rows:
                    nome = r[cabn.index("nome")] if len(r) > cabn.index("nome") else ""
                    if nome:
                        atos.append(dict(nome=nome, tipo="Nomeação", cargo=cargo, simbolo="— (efetivo, concurso)",
                                         lotacao=fator, orgao=org, linha=ini + 1))
                tabelas_lidas.add(ini)
                continue
            tem_simb = any(c.startswith("simbolo") for c in cabn)
            if verbo not in ("nomear", "exonerar"):
                if not tem_simb:
                    continue
                verbo = "exonerar" if cabn and cabn[0].startswith("matricula") else "nomear"
            if not tem_simb and not ("processo" in cabn and "unidade" in cabn):
                continue
            org = org or assinatura_depois(d, fim)
            for r in linhas_tabela(cab, rows):
                atos.append(dict(r, tipo="Nomeação" if verbo == "nomear" else "Exoneração", orgao=org, linha=ini + 1))
                tabelas_lidas.add(ini)
            continue
        # ---- tabelas do texto corrido (dentro do bloco da publicação)
        if d.corrido and re.search(r"resolve\s+(nomear|exonerar)", txt):
            labels = []
            corpo_ini = None
            for j, l in enumerate(ls):
                if l.strip() in ("Matrícula", "Nome", "Cargo", "Símbolo", "Lotação", "Data início", "Data Início",
                                 "Unidade", "Processo", "Unidade Organizacional", "Unidade Organiza-", "cional",
                                 "Sub Grupo", "Inscrição", "Documento (RG)", "Pontuação Final", "Classificação",
                                 "Processo SEI", "Nome Completo", "Ação", "Fator") and not l.startswith(" "):
                    if l.strip() == "cional":
                        continue
                    labels.append(l.strip().rstrip("-"))
                    corpo_ini = j + 1
                elif labels:
                    break
            if labels and "Nome" in labels:
                verbo = re.search(r"resolve\s+(nomear|exonerar)", txt).group(1)
                corpo = ls[corpo_ini:]
                # linhas finais sem espaço inicial, em caixa alta, são assinatura (e às vezes o órgão)
                n_assin = 0
                while n_assin < 2 and len(corpo) > n_assin and not corpo[-1 - n_assin].startswith(" ") \
                        and re.fullmatch(r"[A-ZÀ-Ü .,'-]{6,}", corpo[-1 - n_assin].strip()):
                    n_assin += 1
                org = orgao_portaria(txt) or (corpo[-1].strip() if n_assin else "")
                corpo_cel = corpo[: len(corpo) - n_assin]
                linhas = tabela_corrida(corpo_cel, len(labels))
                lidos = linhas_tabela(labels, linhas)
                concurso = "Inscrição" in labels
                judicial = "Ação" in labels
                if concurso or judicial:
                    mc = re.search(r"Cargo:\s*(.+?)\s*(?:Fator:|Inscrição|$)", txt)
                    mf = re.search(r"Fator:\s*(.+?)\s*(?:Inscrição|Cargo:|$)", txt)
                    i_fator = labels.index("Fator") if "Fator" in labels else None
                    for r, cel in zip(lidos, linhas):
                        if concurso:
                            r["cargo"] = mc.group(1) if mc else "—"
                            r["lotacao"] = mf.group(1) if mf else ""
                            r["simbolo"] = "— (efetivo, concurso)"
                        else:
                            r["lotacao"] = cel[i_fator] if i_fator is not None and i_fator < len(cel) else ""
                            r["simbolo"] = "— (provisório, decisão judicial)"
                for r in lidos:
                    atos.append(dict(r, tipo="Nomeação" if verbo == "nomear" else "Exoneração", orgao=org, linha=ini + 1))
                if not lidos:
                    pend.append(f"L{ini + 1} [{sec}] (tabela não lida) " + " / ".join(l.strip() for l in ls)[:900])
                continue
            if d.corrido and not atos_frase(txt):
                pend.append(f"L{ini + 1} [{sec}] (sem tabela) " + " / ".join(l.strip() for l in ls)[:900])
                continue
        if not d.corrido and re.search(r"resolve\s+(nomear|exonerar)", txt) and not atos_frase(txt):
            cabecalhos.append((ini, sec, txt))
            continue
        # ---- frases
        if re.search(r"(?i)\b(nomear|exonerar|exonerad[oa]s?)\b", txt):
            achou = atos_frase(txt)
            org_port = orgao_portaria(txt)
            for a in achou:
                if not a["orgao"]:
                    a["orgao"] = org_port
                if not a["orgao"]:
                    a["orgao"] = next((o for o in map(orgao_autoridade, reversed(anteriores)) if o), "")
                if not a["orgao"] and (sec.startswith("DECRETOS") or sec == "(início)" or "GOVERNADOR" in txt[:200]):
                    a["orgao"] = "Governador — decretos simples (secretaria não indicada no ato)"
                a["linha"] = ini + 1
                atos.append(a)
            if not achou:
                pend.append(f"L{ini + 1} [{sec}] {txt[:700]}")
    # 'resolve nomear/exonerar' sem tabela lida logo depois
    for ini, sec, txt in cabecalhos:
        if any(ini < t <= ini + 12 for t in tabelas_lidas):
            continue
        depois = " / ".join(limpa(l) for l in d.linhas[ini + 1: ini + 14] if l.strip())
        pend.append(f"L{ini + 1} [{sec}] (tabela não encontrada) {txt[:500]} >>> {depois[:600]}")
    return atos, pend


# ---------------------------------------------------------------- decretos / Casa Civil


def decretos(d):
    res = []
    bl = d.blocos
    textos = [junta(b[3]) if d.corrido else limpa(" ".join(x.strip() for x in b[3])) for b in bl]
    cab_dec = re.compile(r"DECRETO\s+(FINANCEIRO\s+)?N[º°o]\s*([\d.]+)\s+DE\s+(\d{1,2}\s+DE\s+[A-ZÇ]+\s+DE\s+\d{4})")
    for k, (ini, fim, sec, ls) in enumerate(bl):
        for i, linha in enumerate(ls):
            m = cab_dec.match(limpa(linha))
            if not m:
                continue
            resto_bloco = junta(ls[i:]) if d.corrido else limpa(" ".join(x.strip() for x in ls[i:]))
            m = cab_dec.match(resto_bloco)
            if not m:
                continue
            resto = resto_bloco[m.end():].strip(" .")
            ementa = ""
            cand = [resto] + textos[k + 1: k + 7]
            for j, c in enumerate(cand):
                if EMENTA.match(c):
                    ementa = re.split(r"\s+O GOVERNADOR", c)[0][:400]
                    seguinte = cand[j + 1] if j + 1 < len(cand) else ""
                    if not ementa.rstrip().endswith(".") and seguinte[:1].islower():
                        ementa += " " + re.split(r"\s+O GOVERNADOR", seguinte)[0][:200]
                    break
            if not ementa:
                inicio = re.split(r"\s+O GOVERNADOR", resto)[0].strip()
                if 15 < len(inicio) < 400 and inicio[:1].isupper() and not re.match(r"(Art\.|D E C R E T A|ANEXO|IMÓVEL)", inicio):
                    ementa = inicio
            if not ementa:
                art1 = next((re.search(r"Art\.\s*1[º°o]?\s*[-–.]?\s*(.{20,260})", c) for c in cand
                             if re.search(r"Art\.\s*1[º°o]?\s*[-–.]", c)), None)
                if art1:
                    ementa = "(sem ementa no texto) Art. 1º - " + art1.group(1)
            res.append(dict(tipo="Decreto financeiro" if m.group(1) else "Decreto", numero=m.group(2),
                        data=m.group(3).title().replace(" De ", " de "),
                        ementa=ementa or "(ementa não localizada no texto extraído)", linha=ini + 1))
    vistos, uniq = set(), []
    for r in res:
        if (r["tipo"], r["numero"]) not in vistos:
            vistos.add((r["tipo"], r["numero"]))
            uniq.append(r)
    return uniq


def resumo_ato(textos, k):
    """Texto do ato a partir do 'resolve' (juntando até 3 blocos seguintes)."""
    t = textos[k]
    j = k + 1
    while not re.search(r"(?i)resolv|R E S O L V E", t) and j < len(textos) and j <= k + 3:
        t += " " + textos[j]
        j += 1
    m = re.search(r"(?i)\bresolve\s*:?|R E S O L V E", t)
    if m:
        depois = t[m.end():].strip()
        if len(depois) < 60 and j < len(textos):
            depois += " " + textos[j]
        return depois[:420]
    return t[:420]


ATO_CC = re.compile(r"(?i)^(?:O|A)\s+SECRET[ÁA]RI[OA](?:\(A\))?\s+(?:DE ESTADO\s+(?:EM EXERC[ÍI]CIO\s+)?do\(a\)\s+|DA\s+)CASA CIVIL")


def casa_civil(d):
    port, outros = [], []
    textos = [junta(b[3]) if d.corrido else limpa(" ".join(l.strip() for l in b[3])) for b in d.blocos]
    for k, (ini, fim, sec, ls) in enumerate(d.blocos):
        txt = textos[k]
        na_cc = sec == "CASA CIVIL" or sec.endswith(" CASA CIVIL")
        titulo = re.match(r"(?i)^portaria\b.{0,80}?\b20\d{2}\b\.?", txt)
        if titulo and na_cc:
            corpo = txt[titulo.end():].strip()
            if len(corpo) < 40 and k + 1 < len(textos) and not re.match(r"(?i)^portaria", textos[k + 1]):
                corpo = textos[k + 1]
            if not ATO_CC.match(corpo) and not re.search(r"(?i)casa civil", titulo.group(0) + corpo[:400]):
                continue  # portaria de outro órgão que caiu na seção por intercalação
            port.append(dict(numero=titulo.group(0).strip(" ."), texto=resumo_ato([corpo] + textos[k + 2:k + 5], 0),
                             linha=ini + 1))
        elif ATO_CC.match(txt):
            if k and any(p["linha"] == d.blocos[k - 1][0] + 1 for p in port):
                continue  # já registrado junto com o título
            num = None
            for j in range(k - 1, max(-1, k - 7), -1):
                num = re.match(r"(?i)^portaria\b.{0,80}?\b20\d{2}\b", textos[j])
                if num:
                    break
            port.append(dict(numero=num.group(0) if num else "(sem número no trecho)", texto=resumo_ato(textos, k), linha=ini + 1))
        elif na_cc and txt and txt != "CASA CIVIL":
            outros.append(dict(texto=txt[:400], linha=ini + 1))
    return port, outros


# ---------------------------------------------------------------- SSP


def ssp(d):
    res = []
    bl = d.blocos
    for k, (ini, fim, sec, ls) in enumerate(bl):
        txt = junta(ls) if d.corrido else limpa(" ".join(l.strip() for l in ls))
        if txt.startswith("|"):
            continue
        na_secao = "SEGURANÇA PÚBLICA" in sec
        em_licit = sec.startswith("AVISO") or sec.startswith("RESULTADOS")
        if not (na_secao or em_licit):
            continue
        if not INICIO_LIC.match(txt):
            continue
        full = txt
        j = k + 1
        # Junta os parágrafos seguintes até ~260 caracteres; se ainda for só
        # cabeçalho (ex.: "APOSTILA Nº 074/2026/CAF" + considerandos, como
        # vem do site), segue até achar o conteúdo do extrato.
        while j < len(bl) and ((len(full) < 260 and j <= k + 3) or (sem_conteudo(full) and j <= k + 6)):
            nt = junta(bl[j][3]) if d.corrido else limpa(" ".join(x.strip() for x in bl[j][3]))
            if INICIO_LIC.match(nt) and len(full) > 80:
                break
            full += " " + nt
            j += 1
        if not SSP_TXT.search(full) or re.match(r"(?i)ata d[ae] (reuni|assembl)", full):
            continue
        if em_licit and not na_secao and not re.search(r"(SSP|PMBA|PM/BA|PCBA|DPT|CBMBA|CBM/BA|Segurança Pública)", full):
            continue
        res.append(extrai_ssp(full, ini + 1))
    vistos, uniq = set(), []
    for r in res:
        if r["sei"] == "—" and r["valor"] == "—" and not r["tem_objeto"] and r["empresa"] == "—":
            continue  # cabeçalho solto, sem conteúdo do extrato
        chave = (r["sei"], r["valor"]) if r["sei"] != "—" else (r["tipo"][:60], r["valor"])
        chave_obj = ("obj", r["objeto"][:90], r["empresa"])
        if chave not in vistos and chave_obj not in vistos:
            vistos.update([chave, chave_obj])
            uniq.append(r)
    return uniq


def sem_conteudo(t):
    """Trecho sem processo SEI, contratada, objeto nem valor."""
    return not (SEI.search(t.replace(" ", "")) or re.search(r"(?i)contratad|objeto|R\$", t))


NUM_RS = r"R\$\s?(?:R\$\s?)?([\d.]+,\d{2})"


def _num(s):
    return float(s.replace(".", "").replace(",", "."))


def _rs(v):
    s = f"{abs(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return ("-" if v < 0 else "") + "R$ " + s


def valor_aditivo(t):
    """Valor do próprio aditivo, não o do contrato depois dele (o último R$
    dos resumos de acréscimo/supressão costuma ser o novo valor total):
    1) valor escrito ("quantia total de", "acréscimo de", "valor supressão",
       "fica acrescido ... o montante de"), somando se houver mais de um;
    2) acréscimo/supressão só como "era de X ... passa a ser Y": Y - X;
    3) sem isso (prorrogação: mensal ... global do período): None, e fica o
       último valor do texto, como antes."""
    explicitos = list(re.finditer(
        r"(?i)(quantia (?:total )?(?:de|em)|acr[ée]scimo (?:total )?(?:de|no valor de|em)|"
        r"(?:acrescid|suprimid)[oa][^$]{0,60}?montante de|"
        r"valor (?:da |do |de )?(supress[ãa]o|acr[ée]scimo)(?: de)?|"
        r"supress[ãa]o (?:total )?(?:de|no valor de|em)|"
        r"valor (?:total |global )?do (?:presente )?(?:\d+º )?(?:termo )?aditivo(?: de)?)"
        r"\s*:?\s*" + NUM_RS, t))
    if explicitos:
        # várias parcelas (ex.: revisão INPC + CCT): soma, sem repetir valor
        vistos, total = set(), 0.0
        for m in explicitos:
            if m.group(3) in vistos:
                continue
            vistos.add(m.group(3))
            v = _num(m.group(3))
            total += -v if re.search(r"(?i)supress|suprimid", m.group(1)) else v
        return _rs(total)
    if re.search(r"(?i)acr[ée]scimo|acrescid|supress", t):
        m = re.search(NUM_RS + r"[^$]{0,250}?(?:para|passa\w* a ser|passou a ser)"
                      r"[^$\d]{0,40}" + NUM_RS, t)
        if m:
            return _rs(_num(m.group(2)) - _num(m.group(1)))
    return None


def extrai_ssp(t, linha):
    tipo = re.split(r"(?<=[a-zà-ü0-9\)])\.\s|\s[-–]\s(?=Abertura|O |A |Processo|Contratante)|\s\|\s|(?i:processo)|"
                    r"(?i:contratante)|Abertura:|\sTermo Aditivo n", t)[0][:160].strip(" .-|;")
    sei = ", ".join(dict.fromkeys(SEI.findall(t.replace(" ", "")))) or "—"
    mo = re.search(r"(?i)objeto\s*:?\s*(.+?)(?:\.\s(?:N[º°]|Fam[íi]lia|Reg[êe]ncia|Valor|Vig[êe]ncia|Prazo|Unidade|Data|Dota|Salvador|Fundamento|O Edital)|;\s|$)", t)
    objeto = (mo.group(1) if mo else "")[:260].strip()
    mc = re.search(r"(?i)(?:contratad[oa]s?|empresas? (?:vencedora|adjudicat[áa]ria)s?[^:]{0,40}|fornecedor(?:a)?|credenciad[oa]|adjudicat[áa]ria|empresa)\s*:\s*(.{3,220}?)(?=,?\s*(?:CNPJ|inscrita|Objeto|OBJETO|Valor|VALOR|Crit[ée]rio|Prazo|PRAZO|Vig[êe]ncia|$)|\.\s)", t)
    empresa = mc.group(1).strip(" ,;-")[:200] if mc else "—"
    mv = re.findall(r"R\$\s?[\d.]+,\d{2}", t)
    valor = mv[-1] if mv else "—"
    if re.search(r"(?i)aditivo", t[:200]):
        valor = valor_aditivo(t) or valor
    DATA = r"(?:0?[1-9]|[12]\d|3[01])[./](?:0?[1-9]|1[0-2])[./]20\d{2}"
    md = re.findall(rf"\b{DATA}\b", t)
    mda = re.search(rf"(?i)(?:data da assinatura|assinatura)\s*:?\s*({DATA})", t)
    data = mda.group(1) if mda else (md[-1] if md else "—")
    return dict(tipo=tipo, sei=sei, objeto=objeto or t[:200], tem_objeto=bool(objeto), empresa=empresa,
                valor=valor, data=data, linha=linha)


# ---------------------------------------------------------------- relatório


CANONICO = [  # (palavras que precisam aparecer, rótulo) — siglas como aparecem no próprio diário
    (["ADM", "PENIT"], "Secretaria de Administração Penitenciária e Ressocialização (SEAP)"),
    (["SECRETARIA", "SAUDE"], "Secretaria da Saúde (SESAB)"),
    (["SEC", "EDUCACAO"], "Secretaria da Educação (SEC)"),
    (["SEGURANCA PUBLICA"], "Secretaria da Segurança Pública (SSP)"),
    (["SECRETARIA", "ADMINISTRACAO"], "Secretaria da Administração (SAEB)"),
    (["SECRETARIA", "CULTURA"], "Secretaria de Cultura (SECULT)"),
    (["DESENVOLVIMENTO RURAL"], "Secretaria de Desenvolvimento Rural (SDR)"),
    (["SECRETARIA", "DESENVOLVIMENTO URBANO"], "Secretaria de Desenvolvimento Urbano (SEDUR)"),
    (["SECRETARIA", "DESENVOLVIMENTO ECONOMICO"], "Secretaria de Desenvolvimento Econômico (SDE)"),
    (["SECRETARIA", "FAZENDA"], "Secretaria da Fazenda (SEFAZ)"),
    (["JUSTICA", "DIREITOS HUMANOS"], "Secretaria de Justiça e Direitos Humanos (SJDH)"),
    (["ASSIST", "DESENV", "SOCIAL"], "Secretaria de Assistência e Desenvolvimento Social (SEADES)"),
    (["PROMOCAO DA IGUALDADE"], "Secretaria de Promoção da Igualdade Racial e dos Povos e Comunidades Tradicionais (SEPROMI)"),
    (["SECRETARIA", "TURISMO"], "Secretaria de Turismo (SETUR)"),
    (["CIENCIA", "TECNOLOGIA"], "Secretaria de Ciência, Tecnologia e Inovação (SECTI)"),
    (["TRAB", "EMPREGO"], "Secretaria do Trabalho, Emprego, Renda e Esporte (SETRE)"),
    (["INFRAESTRUTURA HIDRICA"], "Secretaria de Infraestrutura Hídrica e Saneamento (SIHS)"),
    (["RELACOES INSTITUCIONAIS"], "Secretaria de Relações Institucionais"),
    (["COMUNICACAO SOCIAL"], "Secretaria de Comunicação Social"),
    (["SECRETARIA", "PLANEJAMENTO"], "Secretaria do Planejamento"),
    (["SECRETARIA", "MEIO AMBIENTE"], "Secretaria do Meio Ambiente (SEMA)"),
    (["VICE-GOVERNADOR"], "Secretaria do Gabinete do Vice-Governador"),
    (["GABINETE DO GOVERNADOR"], "Gabinete do Governador"),
    (["UNIVERSIDADE DO ESTADO DA BAHIA"], "Universidade do Estado da Bahia (UNEB)"),
    (["FEIRA DE SANTANA"], "Universidade Estadual de Feira de Santana (UEFS)"),
    (["SUDOESTE DA BAHIA"], "Universidade Estadual do Sudoeste da Bahia (UESB)"),
    (["SANTA CRUZ"], "Universidade Estadual de Santa Cruz (UESC)"),
    (["ESTADUAL DE TRANSITO"], "Departamento Estadual de Trânsito (DETRAN)"),
    (["MEIO AMB", "RECURSOS HIDRICOS"], "Instituto do Meio Ambiente e Recursos Hídricos (INEMA)"),
    (["PROCURADORIA GERAL"], "Procuradoria Geral do Estado (PGE)"),
    (["SECRETARIA", "INFRAESTRUTURA"], "Secretaria de Infraestrutura (SEINFRA)"),
    (["SECRETARIA", "AGRICULTURA"], "Secretaria da Agricultura, Pecuária, Irrigação, Pesca e Aquicultura (SEAGRI)"),
    (["POLITICAS", "MULHERES"], "Secretaria de Políticas para as Mulheres (SPM)"),
    (["CASA CIVIL"], "Casa Civil"),
    (["CASA MILITAR"], "Casa Militar do Governador"),
    (["CONTROLADORIA GERAL"], "Controladoria Geral do Estado (CGE)"),
]


def canonico(org):
    k = sem_acento(org).upper()
    for palavras, rotulo in CANONICO:
        if all(p in k for p in palavras):
            return rotulo
    return re.sub(r"\s+\d+$", "", org).strip()


def chave_orgao(org):
    k = sem_acento(canonico(org)).upper()
    k = re.sub(r"\s*-\s*[A-Z/]{2,12}\s*$", "", k)
    k = re.sub(r"\s+(DO ESTADO DA BAHIA|EST DA BAHIA|DA BAHIA)$", "", k)
    return re.sub(r"\s+", " ", k).strip()


def esc(s):
    return str(s).replace("|", "/").replace("\n", " ").strip()




# ---------------------------------------------------------------- secretaria de cada publicação

# Seções do diário que são uma secretaria/órgão do Executivo (o resto —
# decretos, despachos, avisos de licitação, municípios, terceiros, outros
# poderes — não define secretaria)
SECAO_ORGAO = re.compile(r"^(SECRETARIA|CASA CIVIL|CASA MILITAR|PROCURADORIA GERAL|CONTROLADORIA GERAL|"
                         r"GABINETE DO GOVERNADOR|DEFENSORIA)")
SEM_SECRETARIA = "Secretaria não identificada"
_MINUSC = {"da", "de", "do", "das", "dos", "e", "a", "o", "para", "as", "os"}


def _titulo(s):
    p = s.lower().split()
    return " ".join(w if (i and w in _MINUSC) else w[:1].upper() + w[1:] for i, w in enumerate(p))


def secretaria_da_secao(sec):
    """Rótulo da secretaria de uma seção do diário, ou None."""
    s = sem_acento(sec or "").upper().strip()
    if not SECAO_ORGAO.match(s):
        return None
    c = canonico(sec)
    return c if c != re.sub(r"\s+\d+$", "", sec).strip() else _titulo(sec)


def secretaria_do_orgao(org):
    """Secretaria a partir do nome do órgão escrito no ato (quando a seção
    não diz — ex.: decretos simples do Governador)."""
    if not org:
        return None
    c = canonico(org)
    if re.match(r"(?i)(secretaria|casa civil|casa militar|procuradoria geral|controladoria geral|gabinete do governador)", c):
        return c if c != re.sub(r"\s+\d+$", "", org).strip() else _titulo(c)
    return None


def secoes_por_linha(d):
    return {ini + 1: sec for ini, fim, sec, ls in d.blocos}


def licitacoes(d):
    """Licitações, contratos, aditivos etc. de todas as secretarias (a
    mesma leitura do ssp() original, sem exigir que seja da SSP)."""
    res = []
    bl = d.blocos
    em_licit = False
    for k, (ini, fim, sec, ls) in enumerate(bl):
        txt = junta(ls) if d.corrido else limpa(" ".join(l.strip() for l in ls))
        su = sem_acento(sec).upper()
        if su.startswith("AVISO") or su.startswith("RESULTADOS"):
            em_licit = True
        elif not SECAO_ORGAO.match(su):
            em_licit = False
        if txt.startswith("|"):
            continue
        secretaria = secretaria_da_secao(sec)
        if not (secretaria or em_licit):
            continue
        if not INICIO_LIC.match(txt):
            continue
        full = txt
        j = k + 1
        while j < len(bl) and ((len(full) < 260 and j <= k + 3) or (sem_conteudo(full) and j <= k + 6)):
            nt = junta(bl[j][3]) if d.corrido else limpa(" ".join(x.strip() for x in bl[j][3]))
            if INICIO_LIC.match(nt) and len(full) > 80:
                break
            full += " " + nt
            j += 1
        if re.match(r"(?i)ata d[ae] (reuni|assembl)", full):
            continue
        r = extrai_ssp(full, ini + 1)
        r["secretaria"] = secretaria or secretaria_do_orgao(orgao_texto(full)) or SEM_SECRETARIA
        res.append(r)
    vistos, uniq = set(), []
    for r in res:
        if r["sei"] == "—" and r["valor"] == "—" and not r["tem_objeto"] and r["empresa"] == "—":
            continue  # cabeçalho solto, sem conteúdo do extrato
        chave = (r["sei"], r["valor"]) if r["sei"] != "—" else (r["tipo"][:60], r["valor"])
        chave_obj = ("obj", r["objeto"][:90], r["empresa"])
        if chave not in vistos and chave_obj not in vistos:
            vistos.update([chave, chave_obj])
            uniq.append(r)
    return uniq


# ---------------------------------------------------------------- execução


LIMITACOES = [
    "Os dados vêm do texto do Diário Oficial baixado do site oficial (dool.egba.ba.gov.br); "
    "convém conferir no diário qualquer ato importante.",
    "Em grupos de decretos parecidos publicados juntos (ex.: várias desapropriações), a ementa pode ter sido "
    "associada ao decreto vizinho.",
    "A secretaria de cada ato ou licitação é a da seção do diário em que ele saiu (autarquias, fundações, "
    "empresas e universidades entram na secretaria a que estão vinculadas no diário). Quando o diário não "
    "indica, aparece como “" + SEM_SECRETARIA + "”.",
    "Licitações e contratos: avisos, resultados, extratos, aditivos, dispensas, inexigibilidades, convênios e "
    "congêneres. Portarias internas (férias, diárias, designações etc.) não são listadas.",
    "Decretos e portarias da Casa Civil não são filtrados por secretaria.",
]


def executar(entrada, arquivo_json, arquivo_manual=None):
    """Lê os doe_*.md de `entrada` e grava `arquivo_json` (painel do DOE).
    Devolve o dicionário gravado."""
    entrada, arquivo_json = Path(entrada), Path(arquivo_json)
    manual = {}
    if arquivo_manual and Path(arquivo_manual).exists():
        manual = json.loads(Path(arquivo_manual).read_text(encoding="utf-8"))
    diarios, secretarias = [], set()
    for arq in sorted(entrada.glob("doe_*.md")):
        d = Diario(arq)
        data_iso = re.search(r"\d{4}-\d{2}-\d{2}", arq.stem).group(0)
        sup = "suplemento" in arq.stem
        atos, _pend = pessoal(d)
        decs = decretos(d)
        port_cc, _outros = casa_civil(d)
        lic = licitacoes(d)
        man = manual.get(arq.stem, {})
        remover = {sem_acento(n).upper() for n in man.get("remover_atos", [])}
        atos = [a for a in atos if sem_acento(a["nome"]).upper() not in remover]
        atos += [dict(a, linha=a.get("linha", 0)) for a in man.get("atos", [])]
        remover_lic = set(man.get("remover_ssp", []))
        lic = [r for r in lic if r["linha"] not in remover_lic]
        lic += [dict(r, linha=r.get("linha", 0), secretaria="Secretaria da Segurança Pública (SSP)")
                for r in man.get("ssp", [])]
        port_cc += man.get("casa_civil_portarias", [])
        secao_de = secoes_por_linha(d)

        atos_json, vistos = [], set()
        for a in atos:
            org = a["orgao"] if a["orgao"] and len(a["orgao"]) < 90 else "(órgão não identificado)"
            assin = (sem_acento(a["nome"]).upper(), a["tipo"], a["cargo"].lower())
            if assin in vistos:
                continue
            vistos.add(assin)
            sec = (secretaria_da_secao(secao_de.get(a.get("linha"), "")) or secretaria_do_orgao(a["orgao"])
                   or SEM_SECRETARIA)
            secretarias.add(sec)
            atos_json.append(dict(secretaria=sec, orgao=canonico(org), nome=a["nome"], tipo=a["tipo"],
                                  cargo=a["cargo"], simbolo=a["simbolo"], lotacao=a["lotacao"] or "—"))
        atos_json.sort(key=lambda x: (sem_acento(x["secretaria"]).upper(), sem_acento(x["nome"]).upper()))
        lic_json = []
        for r in lic:
            secretarias.add(r["secretaria"])
            lic_json.append({k: r[k] for k in ("secretaria", "tipo", "sei", "objeto", "empresa", "valor", "data")})
        diarios.append(dict(
            data=data_iso, suplemento=sup, arquivo=arq.name,
            decretos=[dict(numero=r["numero"], data=r["data"], tipo=r["tipo"], ementa=r["ementa"]) for r in decs],
            casa_civil=[dict(numero=p["numero"], texto=p["texto"][:450]) for p in port_cc],
            atos=atos_json, colegiados=man.get("colegiados", []), licitacoes=lic_json,
            notas=man.get("notas", [])))

    datas = sorted({x["data"] for x in diarios})
    sem_arquivo = []
    if datas:
        dia, d1 = date.fromisoformat(datas[0]), date.fromisoformat(datas[-1])
        while dia <= d1:
            if dia.weekday() in (1, 2, 3, 4, 5) and dia.isoformat() not in datas:  # terça a sábado
                sem_arquivo.append(dia.strftime("%d/%m"))
            dia += timedelta(days=1)
    painel = dict(
        gerado_em=datetime.now().strftime("%d/%m/%Y %H:%M"),
        periodo=[datas[0], datas[-1]] if datas else [],
        sem_arquivo=sem_arquivo,
        secretarias=sorted(secretarias, key=lambda s: (s == SEM_SECRETARIA, sem_acento(s).upper())),
        limitacoes=LIMITACOES,
        diarios=diarios)
    arquivo_json.parent.mkdir(parents=True, exist_ok=True)
    tmp = arquivo_json.with_suffix(".tmp")
    tmp.write_text(json.dumps(painel, ensure_ascii=False, separators=(",", ":")), encoding="utf-8")
    tmp.replace(arquivo_json)
    return painel


if __name__ == "__main__":
    import sys
    p = executar(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
    print(len(p["diarios"]), "diários,", len(p["secretarias"]), "secretarias")
