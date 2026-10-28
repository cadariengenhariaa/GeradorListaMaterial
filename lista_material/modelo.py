"""Organização dos quantitativos: itens, níveis, ambientes (colunas) e lançamento nas células.

Regras (as mesmas da lista A056 preenchida manualmente):
- cada nível de referência vira um grupo de colunas e cada local atendido vira uma coluna;
- local atendido com vários ambientes (ex.: SAT-MOTORISTA) tem os ambientes em colunas vizinhas e o
  quantitativo vai numa célula mesclada; se a linha já tiver valor próprio de um desses ambientes,
  o valor vai para um ambiente livre do grupo (ou é somado a um deles), com comentário na célula;
- local atendido com o nome de um nível (ex.: "1º SUBSOLO") ou com "GERAL" vira a coluna GERAL do nível;
- poucas peças com nível de referência diferente do ambiente atendido (ex.: condensadora no 2º subsolo
  atendendo a academia do 1º subsolo) são lançadas no ambiente atendido, com comentário;
- itens com a mesma descrição e unidade dentro da mesma subseção são somados.
"""
import itertools
import re
from collections import OrderedDict, defaultdict
from dataclasses import dataclass, field

from . import config
from .leitura import normalizar

COMUM, PRIVATIVA = "COMUM", "PRIVATIVA"
NOME_ABA = {COMUM: config.ABA_COMUM, PRIVATIVA: config.ABA_PRIVATIVA}


def chave(txt):
    t = str(txt or "").replace("º", "°").replace("ª", "°")
    return normalizar(t).rstrip(":").strip()


def limpar_local(txt):
    return re.sub(r"\s+", " ", str(txt or "")).strip().rstrip(":").strip()


def titulo_ambiente(nome):
    t = limpar_local(nome)
    for de, para in config.ABREVIACOES_LOCAL:
        if t.upper().startswith(de):
            t = para + t[len(de):]
    return t


def titulo_nivel(nome):
    return limpar_local(nome).replace("º", "°")


def separar_ambientes(local):
    """'VESTIÁRIO FEM.-VESTIÁRIO MASC.-BANHO KIDS' -> 3 ambientes. ' - ' com espaço não separa."""
    nome = limpar_local(local)
    partes = re.split(r"(?<=\S)-(?=\S)", nome)
    if len(partes) == 1:
        return [nome]
    juntas = []
    for p in partes:
        if juntas and normalizar(juntas[-1]) in config.PREFIXOS_HIFENIZADOS:
            juntas[-1] = juntas[-1] + "-" + p
        else:
            juntas.append(p)
    juntas = [j.strip() for j in juntas if j.strip()]
    if len(juntas) < 2 or any(len(j) < 2 for j in juntas):
        return [nome]
    return juntas


def _ordem_natural(txt):
    return [int(t) if t.isdigit() else t for t in re.split(r"(\d+)", chave(txt))]


def elevacao(nivel):
    k = chave(nivel)
    m = re.search(r"(\d+)\s*°?\s*SUBSOLO", k)
    if m:
        return -int(m.group(1))
    if "SUBSOLO" in k:
        return -1
    if "TERREO" in k:
        return 0
    m = re.search(r"(\d+)", k)
    if m:
        return int(m.group(1))
    if "COBERTURA" in k:
        return 1000
    return None


def aba_padrao(nivel):
    k = chave(nivel)
    return COMUM if any(normalizar(t) in k for t in config.TERMOS_AREA_COMUM) else PRIVATIVA


# ---------------------------------------------------------------------------
# Estruturas
# ---------------------------------------------------------------------------
@dataclass
class ConfigNivel:
    nivel: str
    cabecalho: str
    aba: str
    repeticoes: int
    instancias: int


@dataclass
class Coluna:
    chave: tuple            # (nível no modelo, chave do ambiente)
    nivel: str
    ambiente: str           # nome do ambiente (para o título padrão)
    titulo: str
    origens: list = field(default_factory=list)
    instancias: int = 0


@dataclass
class Item:
    secao: str
    subsecao: object
    descricao: str
    unidade: str
    ordem: int
    # aba -> {(nível, (ambientes...)): qtd}
    qtd: dict = field(default_factory=lambda: {COMUM: defaultdict(float), PRIVATIVA: defaultdict(float)})
    # aba -> {(nível, ambiente): [(qtd, nível original, local original)]}  (peças realocadas)
    realocadas: dict = field(default_factory=lambda: {COMUM: defaultdict(list), PRIVATIVA: defaultdict(list)})
    # aba -> {(nível, (ambientes...)): local atendido original}
    rotulos: dict = field(default_factory=lambda: {COMUM: {}, PRIVATIVA: {}})


@dataclass
class Lancamento:
    """Uma célula (ou células mescladas) de uma linha de item."""
    inicio: int
    fim: int
    qtd: float
    comentario: str = ""


# ---------------------------------------------------------------------------
# Análise
# ---------------------------------------------------------------------------
class Analise:
    def __init__(self, instancias, realocar=True):
        self.instancias = list(instancias)
        self.realocar = realocar
        self._contar()
        self.niveis = self._niveis_padrao()
        self.titulos_colunas = {}           # chave da coluna -> título editado
        self.ordem_ambientes = {}           # nível -> [chaves de ambiente] (ordem editada)

    # ----- estatísticas -------------------------------------------------
    def _contar(self):
        self.cont_nivel = OrderedDict()
        self.cont_nivel_local = defaultdict(int)
        self.cont_local = defaultdict(int)
        self.primeira_linha_nivel = {}
        for inst in self.instancias:
            n = inst.nivel
            self.cont_nivel[n] = self.cont_nivel.get(n, 0) + 1
            self.primeira_linha_nivel.setdefault(n, inst.linha)
            self.cont_nivel_local[(n, chave(inst.local))] += 1
            self.cont_local[chave(inst.local)] += 1
        self.chaves_niveis = {chave(n): n for n in self.cont_nivel}

    def _niveis_padrao(self):
        def ordem(n):
            e = elevacao(n)
            return (e if e is not None else 10 ** 6, 1 if "TIPO" in chave(n) else 0, self.primeira_linha_nivel[n])
        out = OrderedDict()
        for n in sorted(self.cont_nivel, key=ordem):
            out[n] = ConfigNivel(nivel=n, cabecalho=titulo_nivel(n), aba=aba_padrao(n), repeticoes=1,
                                 instancias=self.cont_nivel[n])
        return out

    def mover_nivel(self, nivel, passo):
        nomes = list(self.niveis)
        i = nomes.index(nivel)
        j = max(0, min(len(nomes) - 1, i + passo))
        nomes.insert(j, nomes.pop(i))
        self.niveis = OrderedDict((n, self.niveis[n]) for n in nomes)

    # ----- destino de cada instância -------------------------------------
    def _destinos(self):
        pares = {}
        for inst in self.instancias:
            par = (inst.nivel, inst.local)
            if par not in pares:
                pares[par] = self._destino(*par)
        return pares

    def _destino(self, nivel, local):
        """Retorna (nível de destino, [chaves dos ambientes], [nomes dos ambientes], realocada?)."""
        k_local = chave(local)
        geral_de = self.chaves_niveis.get(k_local)
        eh_geral = geral_de is not None or re.search(r"\bGERAL\b", k_local) is not None
        destino = nivel
        if self.realocar:
            if geral_de is not None and geral_de != nivel and \
                    self.niveis[geral_de].aba == self.niveis[nivel].aba:
                destino = geral_de
            elif geral_de is None:
                cont = {n: self.cont_nivel_local.get((n, k_local), 0) for n in self.cont_nivel}
                casa = max(cont, key=lambda n: (cont[n], -list(self.cont_nivel).index(n)))
                c = cont[nivel]
                if casa != nivel and self.niveis[casa].aba == self.niveis[nivel].aba and \
                        c <= config.REALOCAR_MAX_INSTANCIAS and \
                        c / max(1, self.cont_local[k_local]) <= config.REALOCAR_MAX_FRACAO:
                    destino = casa
        if eh_geral and (geral_de is None or geral_de == destino):
            return destino, [config.COLUNA_GERAL], [config.COLUNA_GERAL], destino != nivel
        nomes = separar_ambientes(local)
        return destino, [chave(x) for x in nomes], [titulo_ambiente(x) for x in nomes], destino != nivel

    # ----- colunas ------------------------------------------------------
    def colunas(self):
        """Colunas de cada aba, na ordem final. Retorna {aba: [Coluna]}."""
        por_nivel = OrderedDict((n, OrderedDict()) for n in self.niveis)
        grupos = defaultdict(list)      # nível -> [conjuntos de ambientes que dividem a célula]
        vistos = set()
        destinos = self._destinos()
        for inst in self.instancias:
            destino, chaves, nomes, _ = destinos[(inst.nivel, inst.local)]
            amb = por_nivel[destino]
            for k, nome in zip(chaves, nomes):
                if k not in amb:
                    amb[k] = Coluna(chave=(destino, k), nivel=destino, ambiente=nome,
                                    titulo=config.COLUNA_GERAL if k == config.COLUNA_GERAL else nome)
                col = amb[k]
                col.instancias += 1
                origem = limpar_local(inst.local) + ("" if inst.nivel == destino else f" (nível {inst.nivel})")
                if origem not in col.origens:
                    col.origens.append(origem)
            if len(chaves) > 1 and (destino, tuple(chaves)) not in vistos:
                vistos.add((destino, tuple(chaves)))
                grupos[destino].append(tuple(chaves))
        saida = {COMUM: [], PRIVATIVA: []}
        for nivel, amb in por_nivel.items():
            if not amb:
                continue
            ordem = self._ordenar_ambientes(nivel, amb, grupos[nivel])
            for k in ordem:
                col = amb[k]
                col.titulo = self.titulos_colunas.get(col.chave, col.titulo)
                saida[self.niveis[nivel].aba].append(col)
        return saida

    def _ordenar_ambientes(self, nivel, amb, grupos):
        chaves = [k for k in amb if k != config.COLUNA_GERAL]
        if nivel in self.ordem_ambientes:
            editada = [k for k in self.ordem_ambientes[nivel] if k in chaves]
            ordem = editada + [k for k in chaves if k not in editada]
        else:
            base = sorted(chaves, key=lambda k: _ordem_natural(amb[k].titulo))
            ordem = _ordem_contigua(base, grupos)
        if config.COLUNA_GERAL in amb:
            ordem.append(config.COLUNA_GERAL)
        return ordem

    def mover_ambiente(self, coluna_chave, passo):
        nivel, k = coluna_chave
        atual = [c.chave[1] for c in self.colunas()[self.niveis[nivel].aba] if c.nivel == nivel]
        atual = [x for x in atual if x != config.COLUNA_GERAL]
        if k not in atual:
            return
        i = atual.index(k)
        j = max(0, min(len(atual) - 1, i + passo))
        atual.insert(j, atual.pop(i))
        self.ordem_ambientes[nivel] = atual

    # ----- itens --------------------------------------------------------
    def itens(self):
        """Itens agrupados em seções/subseções, com os quantitativos por aba."""
        secoes_cfg = [(chave(g), t, [(chave(s), st) for s, st in subs]) for g, t, subs in config.SECOES_PADRAO]
        itens = OrderedDict()
        ordem_secao, ordem_sub = {}, {}
        destinos = self._destinos()
        for pos, inst in enumerate(self.instancias):
            g, s = chave(inst.grupo), chave(inst.subgrupo)
            secao = subsecao = None
            for gi, (gk, titulo, subs) in enumerate(secoes_cfg):
                if gk == g:
                    secao = titulo
                    ordem_secao.setdefault(secao, gi)
                    for si, (sk, st) in enumerate(subs):
                        if sk == s:
                            subsecao = st
                            ordem_sub.setdefault((secao, subsecao), si)
                            break
                    else:
                        subsecao = inst.subgrupo.strip() or None
                        ordem_sub.setdefault((secao, subsecao), 100 + pos)
                    break
            if secao is None:
                secao = inst.grupo.strip() or "OUTROS ITENS"
                subsecao = inst.subgrupo.strip() or None
                ordem_secao.setdefault(secao, 100 + pos)
                ordem_sub.setdefault((secao, subsecao), pos)
            k_item = (secao, subsecao, inst.descricao.strip(), inst.unidade.strip())
            if k_item not in itens:
                itens[k_item] = Item(secao, subsecao, inst.descricao.strip(), inst.unidade.strip(), pos)
            item = itens[k_item]
            destino, chaves, _, realocada = destinos[(inst.nivel, inst.local)]
            aba = self.niveis[destino].aba
            chave_q = (destino, tuple(chaves))
            item.qtd[aba][chave_q] += inst.qtd
            item.rotulos[aba].setdefault(chave_q, limpar_local(inst.local))
            if realocada:
                item.realocadas[aba][(destino, chaves[0])].append((inst.qtd, inst.nivel, limpar_local(inst.local)))
        # estrutura ordenada
        estrutura = OrderedDict()
        for secao in sorted(ordem_secao, key=lambda s: ordem_secao[s]):
            subs = [sb for (sc, sb) in ordem_sub if sc == secao]
            subs.sort(key=lambda sb: ordem_sub[(secao, sb)])
            estrutura[secao] = OrderedDict()
            for sb in subs:
                grupo = [it for it in itens.values() if it.secao == secao and it.subsecao == sb]
                estrutura[secao][sb] = _ordenar_itens(grupo)
        return estrutura


def _ordem_contigua(base, grupos):
    """Ordena os ambientes deixando vizinhos os que dividem célula (local com vários ambientes)."""
    if not grupos:
        return list(base)
    pai = {k: k for k in base}

    def achar(x):
        while pai[x] != x:
            pai[x] = pai[pai[x]]
            x = pai[x]
        return x
    for g in grupos:
        for a, b in zip(g, g[1:]):
            if a in pai and b in pai:
                pai[achar(a)] = achar(b)
    comps = defaultdict(list)
    for k in base:
        comps[achar(k)].append(k)

    def contiguos(ordem, gs):
        pos = {k: i for i, k in enumerate(ordem)}
        n = 0
        for g in gs:
            idx = sorted(pos[k] for k in g if k in pos)
            if idx and idx[-1] - idx[0] == len(idx) - 1:
                n += 1
        return n
    melhor = {}
    for raiz, membros in comps.items():
        gs = [g for g in grupos if g[0] in membros]
        if len(membros) <= 1 or not gs:
            melhor[raiz] = membros
            continue
        if len(membros) <= 7:
            escolha, nota = membros, -1
            for perm in itertools.permutations(membros):
                c = contiguos(perm, gs)
                if c > nota:
                    escolha, nota = list(perm), c
                    if c == len(gs):
                        break
            melhor[raiz] = escolha
        else:
            ordem = []
            for g in sorted(gs, key=len, reverse=True):
                for k in g:
                    if k not in ordem:
                        ordem.append(k)
            melhor[raiz] = ordem + [k for k in membros if k not in ordem]
    saida, feitos = [], set()
    for k in base:
        r = achar(k)
        if r in feitos:
            continue
        feitos.add(r)
        saida.extend(melhor[r])
    return saida


def _fracoes(txt):
    vals = []
    for a, b in re.findall(r"DN\s*(\d+)\s*/\s*(\d+)", txt):
        vals.append(int(a) / int(b))
    return vals


def _ordenar_itens(grupo):
    """Mantém a ordem do Revit; tubos e conexões de cobre ficam em ordem de diâmetro."""
    grupo = sorted(grupo, key=lambda it: it.ordem)
    com_dn = [it for it in grupo if _fracoes(it.descricao)]
    if len(com_dn) < max(2, len(grupo) / 2):
        return grupo

    def tipo(it):
        return re.split(r"[‚,]", it.descricao, maxsplit=1)[0].strip().upper()
    ordem_tipo = {}
    for it in grupo:
        t = tipo(it)
        ordem_tipo.setdefault(t, (1 if "FLEX" in normalizar(t) else 0, len(ordem_tipo)))
    return sorted(grupo, key=lambda it: (ordem_tipo[tipo(it)], _fracoes(it.descricao) or [99], it.ordem))


# ---------------------------------------------------------------------------
# Lançamento de uma linha
# ---------------------------------------------------------------------------
def _fmt(q, unidade):
    if unidade.strip().lower() in config.UNIDADES_DECIMAIS:
        return f"{q:.2f}".replace(".", ",") + " " + unidade
    return f"{int(round(q))} {unidade}"


def lancar_linha(item, aba, colunas):
    """Distribui os quantitativos do item nas colunas da aba. Retorna [Lancamento] em ordem de coluna."""
    idx = {c.chave: i for i, c in enumerate(colunas)}
    titulos = {i: c.titulo for i, c in enumerate(colunas)}
    unicos, multiplos = defaultdict(float), OrderedDict()
    for (nivel, chaves), q in item.qtd[aba].items():
        if abs(q) < 1e-9:
            continue
        cols = [idx[(nivel, k)] for k in chaves if (nivel, k) in idx]
        if len(cols) == 1:
            unicos[cols[0]] += q
        elif cols:
            multiplos[(nivel, chaves)] = (sorted(cols), cols, q)
    usados = set(unicos)
    faixas = {m: set(range(v[0][0], v[0][-1] + 1)) for m, v in multiplos.items()}
    mescladas = []                      # Lancamento com fim > inicio
    avulsos = {}                        # coluna -> Lancamento de um ambiente escolhido
    comentarios = defaultdict(list)
    pendentes = []
    for m, (ordenadas, cols, q) in multiplos.items():
        contigua = ordenadas[-1] - ordenadas[0] == len(ordenadas) - 1
        outras = set()
        for m2, f in faixas.items():
            if m2 != m:
                outras |= f
        if contigua and not (faixas[m] & usados) and not (faixas[m] & outras):
            mescladas.append(Lancamento(ordenadas[0], ordenadas[-1], q))
            usados |= faixas[m]
        else:
            pendentes.append(m)
    for m in pendentes:
        ordenadas, cols, q = multiplos[m]
        rotulo = item.rotulos[aba].get(m, "-".join(m[1]))
        livres = [i for i in cols if i not in usados]
        compartilhadas = set()
        for m2 in pendentes:
            if m2 != m:
                compartilhadas |= set(multiplos[m2][1])
        exclusivas = [i for i in livres if i not in compartilhadas]
        if livres:
            alvo = (exclusivas or livres)[0]
            avulsos[alvo] = Lancamento(alvo, alvo, q)
            usados.add(alvo)
            comentarios[alvo].append(
                f'Inclui {_fmt(q, item.unidade)} do local atendido "{rotulo}". Não foi possível mesclar as '
                f'colunas desses ambientes nesta linha; valor lançado em {titulos[alvo]}.')
        else:
            alvo = cols[0]
            dentro = next((l for l in mescladas if l.inicio <= alvo <= l.fim), None)
            if alvo in unicos:
                unicos[alvo] += q
            elif alvo in avulsos:
                avulsos[alvo].qtd += q
            elif dentro is not None:
                dentro.qtd += q
                alvo = dentro.inicio
            else:
                unicos[alvo] += q
            comentarios[alvo].append(
                f'Inclui {_fmt(q, item.unidade)} do local atendido "{rotulo}". Não foi possível mesclar as '
                f'colunas desses ambientes nesta linha; valor somado ao quantitativo de {titulos[alvo]}.')
    lanc = mescladas + list(avulsos.values()) + [Lancamento(i, i, q) for i, q in unicos.items()]
    for (nivel, k), pecas in item.realocadas[aba].items():
        i = idx.get((nivel, k))
        if i is None:
            continue
        por_origem = OrderedDict()
        for q, n_orig, l_orig in pecas:
            por_origem[(n_orig, l_orig)] = por_origem.get((n_orig, l_orig), 0.0) + q
        for (n_orig, l_orig), q in por_origem.items():
            comentarios[i].append(f'Inclui {_fmt(q, item.unidade)} com nível de referência {n_orig} e local '
                                  f'atendido "{l_orig}" no modelo.')
    for i, textos in comentarios.items():
        dono = next((l for l in lanc if l.inicio <= i <= l.fim), None)
        if dono is not None:
            dono.comentario = (dono.comentario + " " if dono.comentario else "") + " ".join(textos)
    return sorted(lanc, key=lambda l: l.inicio)
