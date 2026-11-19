"""
buscar_api.py — roda no GitHub Actions.

Busca as atividades da empresa na API do mapeia e monta, numa pasta de trabalho, os mesmos
arquivos que o robo antigo baixava (Relatorio, Aprovacoes e Usuarios em .xlsx), para o
gerar_dashboard.py funcionar sem mudancas.

Uso: python buscar_api.py <pasta_de_trabalho>
A chave vem da variavel de ambiente MAPEIA_API_KEY (segredo do repositorio).
"""
import json, os, sys, shutil, gzip, unicodedata
import urllib.request, urllib.parse, urllib.error
from datetime import datetime
from pathlib import Path
import openpyxl

API   = "https://ruwfvtgacnimpebelxfd.supabase.co/functions/v1/api-relatorio"
AQUI  = Path(__file__).parent
RAIZ  = AQUI.parent
CFG   = json.loads((AQUI / "painel.json").read_text(encoding="utf-8"))
EMPRESA = CFG["empresa"]
HOJE  = datetime.now().strftime("%d%m%Y")

COLS_REL = ['#', 'Empresa', 'Setor', 'Funcionário', 'Cargo', 'Nome', 'Detalhamento',
            'Periodicidade', 'Como realiza', 'Tempo', 'Horário fixo', 'Pontos críticos',
            'Sugestões de melhoria', 'Status', 'Decisão gestor', 'Decisão consultor',
            'Edições do consultor', 'Parecer do consultor', 'Criada em', 'Atualizada em']


def _get(params):
    chave = os.environ.get("MAPEIA_API_KEY", "").strip()
    if not chave:
        sys.exit("ERRO: segredo MAPEIA_API_KEY nao configurado no repositorio.")
    url = API + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"Authorization": f"Bearer {chave}",
                                               "Accept-Encoding": "gzip"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r:
            corpo = r.read()
            if r.headers.get("Content-Encoding") == "gzip":
                corpo = gzip.decompress(corpo)
    except urllib.error.HTTPError as e:
        corpo = e.read()
        if e.headers.get("Content-Encoding") == "gzip" or corpo[:2] == b"\x1f\x8b":
            corpo = gzip.decompress(corpo)
        msg = corpo[:300].decode("utf-8", "replace")
        sys.exit(f"ERRO {e.code} da API ({params.get('recurso', 'atividades')}): {msg}")
    return json.loads(corpo.decode("utf-8"))


def _lista(resp):
    """A resposta pode ser uma lista ou um objeto com a lista dentro."""
    if isinstance(resp, list):
        return resp
    for v in resp.values():
        if isinstance(v, list):
            return v
    return []


def _id_empresa():
    if CFG.get("empresa_id"):
        return CFG["empresa_id"]
    filtros = _get({"recurso": "filtros"})
    grupos = filtros.get("empresas") if isinstance(filtros, dict) else None
    if grupos is None:
        grupos = _lista(filtros)
    for e in grupos or []:
        if isinstance(e, dict) and any(str(v).strip().lower() == EMPRESA.lower() for v in e.values()):
            return e.get("id") or e.get("ID")
    sys.exit(f"ERRO: empresa {EMPRESA!r} nao encontrada nos filtros da chave. "
             f"Resposta: {json.dumps(filtros, ensure_ascii=False)[:500]}")


def _chave(s):
    """'Funcionário', 'funcionario', 'decisao_gestor' -> 'funcionario', 'decisaogestor'."""
    s = unicodedata.normalize("NFD", str(s)).encode("ascii", "ignore").decode().lower()
    return "".join(ch for ch in s if ch.isalnum())


# nomes alternativos que a API pode usar para cada coluna do Excel
_SINONIMOS = {
    "Funcionário": ["funcionario", "funcionarionome", "colaborador", "usuario", "autor"],
    "Nome": ["nome", "titulo", "atividade", "nomeatividade"],
    "Criada em": ["criadaem", "criadoem", "createdat", "datacriacao"],
    "Atualizada em": ["atualizadaem", "atualizadoem", "updatedat"],
    "Status": ["status", "statusgeral", "situacao"],
}


def _padronizar(a):
    """Devolve a atividade com as chaves exatamente como as colunas do Excel antigo."""
    idx = {_chave(k): v for k, v in a.items()}
    out = {}
    for c in COLS_REL[1:]:
        for k in [_chave(c)] + _SINONIMOS.get(c, []):
            if k in idx and idx[k] is not None:
                v = idx[k]
                if isinstance(v, dict):  # ex.: {"nome": "..."}
                    v = v.get("nome") or v.get("name") or ""
                out[c] = v
                break
        else:
            out[c] = ""
    return out


def _horas(t):
    try:
        h, m = str(t).split(":")[:2]
        return int(h) + int(m) / 60
    except Exception:
        return 0


def _salvar(caminho, abas):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for nome, linhas in abas:
        ws = wb.create_sheet(nome)
        for l in linhas:
            ws.append(l)
    # o gerador le a aba ativa: como no export do mapeia, ela e a de "Atividades"
    nomes = [n for n, _ in abas]
    wb.active = nomes.index("Atividades") if "Atividades" in nomes else 0
    wb.save(caminho)


def main():
    trab = Path(sys.argv[1]).resolve()
    dest = trab / "ResumoGerencialMapeamentoMillena"
    dest.mkdir(parents=True, exist_ok=True)

    # gerador, logo e prazo
    shutil.copy(AQUI / "gerar_dashboard.py", dest / "gerar_dashboard.py")
    shutil.copy(AQUI / "logo_millena.png", trab / "logo_millena.png")
    (trab / "config.json").write_text(json.dumps({"prazo": CFG.get("prazo", "")}), encoding="utf-8")

    # atividades da API
    brutas = _lista(_get({"recurso": "atividades", "formato": "json", "empresa": _id_empresa()}))
    if brutas:
        print("Campos da API:", sorted(brutas[0].keys()))
    atv = [_padronizar(a) for a in brutas]
    # garante que so entra a empresa do painel, mesmo se a chave enxergar outras
    outras = {a.get("Empresa") for a in atv} - {EMPRESA, ""}
    if outras:
        print(f"Ignorando atividades de outras empresas: {sorted(outras)}")
    atv = [a for a in atv if a.get("Empresa") in (EMPRESA, "")]
    sem_func = sum(not str(a.get("Funcionário", "")).strip() for a in atv)
    print(f"API: {len(atv)} atividades de {EMPRESA} ({len({a.get('Funcionário') for a in atv})} funcionarios; "
          f"{sem_func} sem funcionario)")
    linhas = [COLS_REL]
    for i, a in enumerate(atv, 1):
        linhas.append([i] + [a.get(c, "") if a.get(c) is not None else "" for c in COLS_REL[1:]])

    st = [str(a.get("Status", "")) for a in atv]
    resumo = [
        ["Relatório de Atividades", None],
        [f"Gerado em {datetime.now():%d/%m/%Y, %H:%M:%S} (API)", None],
        [None, None], ["Filtros aplicados", None], ["Empresa", EMPRESA], [None, None],
        ["Resumo executivo", None],
        ["Total de atividades", len(atv)],
        ["Funcionários distintos", len({a.get("Funcionário") for a in atv})],
        ["Aprovadas", sum(s == "Aprovada" for s in st)],
        ["Ressalvas", sum(s == "Ressalva" for s in st)],
        ["Aguardando gestor", sum(s == "Aguardando gestor" for s in st)],
        ["Aguardando consultor", sum(s == "Aguardando consultor" for s in st)],
        ["Tempo total mapeado (horas)", round(sum(_horas(a.get("Tempo")) for a in atv), 1)],
        ["Atividades editadas pelo consultor",
         sum(str(a.get("Edições do consultor", "0")).strip() not in ("", "0") for a in atv)],
        ["Atividades com parecer do consultor",
         sum(bool(str(a.get("Parecer do consultor", "") or "").strip()) for a in atv)],
    ]
    versoes = [["Atividade", "Empresa", "Setor", "Funcionário", "Tipo", "Autor", "Data/hora", "Conteúdo"],
               ["(Nenhuma edição ou parecer)"] + [None] * 7]
    _salvar(dest / f"Relatorio_Millena_{HOJE}.xlsx",
            [("Resumo", resumo), ("Atividades", linhas), ("Versões & Parecer", versoes)])

    # Aprovacoes: mesma fonte, nas colunas do export antigo
    cab = ['#', 'Funcionário', 'Cargo', 'Nome', 'Status geral', 'Decisão gestor',
           'Decisão consultor', 'Criada em', 'Atualizada em']
    apr = [cab] + [[i, a.get("Funcionário", ""), a.get("Cargo", ""), a.get("Nome", ""),
                    a.get("Status", ""), a.get("Decisão gestor", ""), a.get("Decisão consultor", ""),
                    a.get("Criada em", ""), a.get("Atualizada em", "")] for i, a in enumerate(atv, 1)]
    _salvar(dest / f"Aprovacoes_Millena_{HOJE}.xlsx", [("Atividades", apr)])

    # Usuarios: lista mantida em dados/usuarios.json (a API ainda nao tem cadastro)
    usr = json.loads((RAIZ / "dados" / "usuarios.json").read_text(encoding="utf-8"))
    cab_u = ['Nome', 'Cargo', 'Empresa', 'Setor', 'Perfil', 'Status']
    _salvar(dest / f"Usuarios_Millena_{HOJE}.xlsx",
            [("Usuarios", [cab_u] + [[u.get(c.lower(), "") for c in cab_u] for u in usr["usuarios"]])])
    print(f"Usuarios: {len(usr['usuarios'])} (lista de {usr.get('atualizado_em', '?')})")


if __name__ == "__main__":
    main()
