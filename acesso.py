"""Login do CGAPE (pedido do usuário em 25/11/2026).

Tela de usuário e senha antes da tela inicial, com um botão ACESSO LIVRE:
  - acesso livre: só PAC e AÇÕES EM MONITORAMENTO;
  - usuários cadastrados abaixo: todas as seções (PAC, MONITORAMENTO,
    CONTROLE DE PRAZOS e o DIÁRIO OFICIAL — DOE, desde 26/11/2026), salvo
    quem tem lista própria em SECOES_POR_USUARIO.

As senhas NÃO ficam gravadas: só um hash PBKDF2-SHA256 com sal próprio por
usuário (o repositório do GitHub é público). Para cadastrar/trocar uma senha:

    python -c "import acesso; print(acesso.gerar_registro('nome.usuario', 'senha'))"

e cole a linha devolvida em USUARIOS.
"""
from __future__ import annotations

import hashlib
import hmac
import os

ITERACOES = 310_000
# "DOE": Diário Oficial do Estado (doe_monitor.py) — só para os usuários
# cadastrados, pedido de 26/11/2026
SECOES_TODAS = ("PAC", "MON", "PRZ", "DOE")
SECOES_LIVRE = ("PAC", "MON")

# usuario: (nome de exibição, sal hex, hash hex)
USUARIOS = {
    "diego.figueiredo": ("Diego Figueiredo", "efeedbe0ffdd69a97c02dd2a2fae3a82",
                         "6de49d6eedd5f21b3e2475afedfd6bcc58871578e731bccc08d05e9caf40a3c5"),
    "silvana.silva": ("Silvana Silva", "7627fa92396e9ed3f0d638715835c9ee",
                      "64532dd6ec57647f0106ed252cb6d25f9d8cbccf4052191f263ca6cf3047bab5"),
    "lucitania.jesus": ("Lucitania Jesus", "879bb52e91635f9e68685fcd5eac13a7",
                        "58d1a0c72c707ccc9c4c7db03ecfae23179b39c0077f62ed8ddce46c23272d09"),
    "henrique.santana": ("Henrique Santana", "a6130c0983c471dd350007c61a596b06",
                         "6f886c29a6d5879ae7b2f7550199a86aede833b4baec9db576f7f28fd20f629f"),
}

# Seções por usuário quando NÃO são todas (quem não está aqui vê
# SECOES_TODAS). Henrique: só PAC, MONITORAMENTO e DOE — pedido de 26/11/2026.
SECOES_POR_USUARIO = {
    "henrique.santana": ("PAC", "MON", "DOE"),
}


# Quem pode criar/editar/excluir ações do CONTROLE DE PRAZOS (cadastro no
# próprio sistema, controle_prazos.py) — pedido do usuário em 25/11/2026:
# "apenas Silvana, os outros apenas a visualização".
EDITORES_PRAZOS = ("silvana.silva",)


def pode_editar_prazos(sessao) -> bool:
    return str((sessao or {}).get("usuario") or "").lower() in EDITORES_PRAZOS


def _hash(senha: str, sal_hex: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", senha.encode("utf-8"), bytes.fromhex(sal_hex), ITERACOES).hex()


def gerar_registro(usuario: str, senha: str, nome: str = "") -> str:
    sal = os.urandom(16).hex()
    return f'    "{usuario}": ("{nome or usuario}", "{sal}", "{_hash(senha, sal)}"),'


def autenticar(usuario, senha):
    """Sessão {usuario, nome, secoes} ou None. Usuário sem diferenciar
    maiúsculas; a comparação do hash é em tempo constante."""
    chave = str(usuario or "").strip().lower()
    registro = USUARIOS.get(chave)
    if not registro or not senha:
        # mesmo custo de um usuário válido (não revela quem existe)
        _hash(str(senha or "x"), "00" * 16)
        return None
    nome, sal, esperado = registro
    if not hmac.compare_digest(_hash(str(senha), sal), esperado):
        return None
    return sessao_usuario(chave)


def sessao_usuario(chave):
    registro = USUARIOS.get(chave)
    if not registro:
        return None
    return {"usuario": chave, "nome": registro[0],
            "secoes": list(SECOES_POR_USUARIO.get(chave, SECOES_TODAS))}


def sessao_livre():
    return {"usuario": "", "nome": "Acesso livre", "secoes": list(SECOES_LIVRE)}
