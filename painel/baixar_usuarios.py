"""
baixar_usuarios.py — roda no GitHub Actions uma vez por dia.

A API do mapeia ainda nao traz o cadastro de colaboradores. Ate trazer, este script entra no
mapeia com MAPEIA_EMAIL / MAPEIA_SENHA (segredos do repositorio), exporta a lista de Usuarios
da empresa e grava so nome, cargo, setor, perfil e status em dados/usuarios.json
(sem e-mail nem matricula, porque o repositorio e publico).
"""
import json, os, sys, time, tempfile
from datetime import datetime
from pathlib import Path
import openpyxl
from playwright.sync_api import sync_playwright

AQUI = Path(__file__).parent
SAIDA = AQUI.parent / "dados" / "usuarios.json"
EMPRESA = json.loads((AQUI / "painel.json").read_text(encoding="utf-8"))["empresa"]
EMAIL, SENHA = os.environ.get("MAPEIA_EMAIL", ""), os.environ.get("MAPEIA_SENHA", "")


def selecionar_empresa(page):
    page.evaluate("""() => {
      const lab = [...document.querySelectorAll('label,span,div,p')]
        .find(e => e.childElementCount === 0 && e.textContent.trim() === 'Empresa');
      for (let c = lab && lab.parentElement, i = 0; c && i < 3; c = c.parentElement, i++) {
        const ctl = c.querySelector('button[role=combobox], button');
        if (ctl) { ctl.setAttribute('data-filtro', 'empresa'); return; }
      }
    }""")
    combo = page.locator('[data-filtro=empresa]').first
    if combo.inner_text().strip() != EMPRESA:
        combo.click()
        page.locator('[role="option"]').filter(has_text=EMPRESA).first.click()
        time.sleep(1)
    if combo.inner_text().strip() != EMPRESA:
        raise RuntimeError("filtro de empresa nao aplicado")


def main():
    if not (EMAIL and SENHA):
        print("Sem MAPEIA_EMAIL/MAPEIA_SENHA: lista de usuarios mantida como esta.")
        return
    with sync_playwright() as p:
        b = p.chromium.launch()
        ctx = b.new_context(accept_downloads=True)
        page = ctx.new_page()
        entrou = False
        for tent in range(3):
            page.goto("https://mapeia.sinerggia.com.br/login", wait_until="networkidle")
            campo_email = page.locator('input[type="email"], input[name="email"], input[placeholder*="mail" i]').first
            campo_email.wait_for(state="visible", timeout=20_000)
            time.sleep(2)  # deixa a pagina terminar de montar antes de digitar
            campo_email.fill(EMAIL.strip())
            time.sleep(0.5)
            page.locator('input[type="password"]').first.fill(SENHA.strip())
            time.sleep(0.5)
            page.locator('button[type="submit"], button:has-text("Entrar"), button:has-text("Acessar")').first.click()
            try:
                page.wait_for_url(lambda u: "login" not in u.lower(), timeout=30_000)
                entrou = True
                break
            except Exception:
                # mostra so o aviso da tela (sem e-mail nem senha) para saber o motivo
                avisos = page.locator('[role="alert"], [data-sonner-toast], .text-destructive, .text-red-500')
                textos = [t.strip() for t in avisos.all_inner_texts() if t.strip()]
                print(f"login tentativa {tent + 1}: nao entrou. Aviso na tela: {textos or 'nenhum'}")
        if not entrou:
            sys.exit("ERRO: login no mapeia falhou (confira MAPEIA_EMAIL e MAPEIA_SENHA nos segredos)")
        for tent in range(3):
            try:
                page.goto("https://mapeia.sinerggia.com.br/admin/usuarios")
                page.get_by_text("Empresa", exact=True).first.wait_for(timeout=15_000)
                selecionar_empresa(page)
                with page.expect_download(timeout=30_000) as dl:
                    page.locator('button:has-text("Exportar Excel")').first.click()
                arq = Path(tempfile.gettempdir()) / "usuarios.xlsx"
                dl.value.save_as(str(arq))
                break
            except Exception as e:
                print(f"tentativa {tent + 1}: {e}")
                if tent == 2:
                    sys.exit("ERRO: nao foi possivel exportar os usuarios")
        b.close()

    ws = openpyxl.load_workbook(arq).active
    cab = [str(c.value or "") for c in ws[1]]
    usuarios = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(cab, ["" if v is None else str(v).strip() for v in row]))
        if r.get("Nome") and r.get("Empresa", EMPRESA) == EMPRESA:
            usuarios.append({k.lower(): r.get(k, "") for k in ("Nome", "Cargo", "Empresa", "Setor", "Perfil", "Status")})
    usuarios.sort(key=lambda u: u["nome"])
    antigo = json.loads(SAIDA.read_text(encoding="utf-8"))["usuarios"] if SAIDA.exists() else None
    if usuarios == antigo:
        print(f"Usuarios sem mudanca ({len(usuarios)}).")
        return
    SAIDA.write_text(json.dumps({"atualizado_em": datetime.now().strftime("%d/%m/%Y %H:%M"),
                                 "usuarios": usuarios}, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"Usuarios atualizados: {len(usuarios)}")


if __name__ == "__main__":
    main()
