"""Leitura dos documentos e motor da simulação IBS/CBS de 2027.

O módulo não depende do Streamlit. Isso permite validar os layouts e as regras
de cálculo por testes automatizados e reutilizar a mesma lógica em exportações.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO, StringIO
from pathlib import Path
import re
import unicodedata
import xml.etree.ElementTree as ET
import zipfile

import pandas as pd
import xlrd
from xlrd.compdoc import CompDoc
from xlrd.biffh import XLRDError


SIMPLES_PRIMEIRA_FAIXA_2027 = pd.DataFrame(
    [
        ("Anexo I", "Comércio", 4.00, 15.33, 0.17),
        ("Anexo II", "Indústria", 4.50, 13.85, 0.15),
        ("Anexo III", "Serviços (regra geral/fator R)", 6.00, 15.43, 0.17),
        ("Anexo IV", "Serviços do § 5º-C", 4.50, 21.26, 0.24),
        ("Anexo V", "Serviços intelectuais/fator R", 15.50, 16.96, 0.19),
    ],
    columns=[
        "anexo",
        "descricao",
        "aliquota_simples_pct",
        "partilha_cbs_pct",
        "partilha_ibs_pct",
    ],
)
SIMPLES_PRIMEIRA_FAIXA_2027["credito_cbs_pct"] = (
    SIMPLES_PRIMEIRA_FAIXA_2027["aliquota_simples_pct"]
    * SIMPLES_PRIMEIRA_FAIXA_2027["partilha_cbs_pct"]
    / 100
)
SIMPLES_PRIMEIRA_FAIXA_2027["credito_ibs_pct"] = (
    SIMPLES_PRIMEIRA_FAIXA_2027["aliquota_simples_pct"]
    * SIMPLES_PRIMEIRA_FAIXA_2027["partilha_ibs_pct"]
    / 100
)


@dataclass(frozen=True)
class ParametrosSimulacao:
    ibs_pct: float = 0.10
    cbs_pct: float = 8.50
    reducao_saude_pct: float = 60.0
    ir_pct: float = 1.50
    csll_pct: float = 1.00
    anexo_simples_padrao: str = "Anexo III"
    indefinido_como_regular: bool = False


def somente_digitos(valor: object) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = str(int(valor))
    return re.sub(r"\D", "", str(valor))


def normalizar_documento(valor: object) -> str:
    digitos = somente_digitos(valor)
    if not digitos:
        return ""
    if len(digitos) <= 11:
        return digitos.zfill(11)
    if len(digitos) <= 14:
        return digitos.zfill(14)
    return digitos


def formatar_documento(valor: object) -> str:
    digitos = normalizar_documento(valor)
    if len(digitos) == 14:
        return f"{digitos[:2]}.{digitos[2:5]}.{digitos[5:8]}/{digitos[8:12]}-{digitos[12:]}"
    if len(digitos) == 11:
        return f"{digitos[:3]}.{digitos[3:6]}.{digitos[6:9]}-{digitos[9:]}"
    return digitos


def _texto(valor: object) -> str:
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ""
    return str(valor).strip()


def _codigo(valor: object) -> str:
    if isinstance(valor, float) and valor.is_integer():
        return str(int(valor))
    return _texto(valor)


def _numero(valor: object, padrao: float = 0.0) -> float:
    if valor is None or valor == "":
        return padrao
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("R$", "").replace(" ", "")
    if not texto:
        return padrao
    if "," in texto:
        texto = texto.replace(".", "").replace(",", ".")
    try:
        return float(texto)
    except ValueError:
        return padrao


def _data_excel(valor: object) -> pd.Timestamp | pd.NaT:
    if valor in (None, ""):
        return pd.NaT
    if isinstance(valor, (int, float)):
        return pd.Timestamp(datetime(1899, 12, 30) + timedelta(days=float(valor)))
    return pd.to_datetime(valor, dayfirst=True, errors="coerce")


def _sem_acento(texto: object) -> str:
    valor = unicodedata.normalize("NFKD", _texto(texto))
    valor = "".join(c for c in valor if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", valor.upper()).strip()


def _nome_local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _filho(elemento: ET.Element | None, nome: str) -> ET.Element | None:
    if elemento is None:
        return None
    return next((item for item in list(elemento) if _nome_local(item.tag) == nome), None)


def _descendente(elemento: ET.Element | None, nome: str) -> ET.Element | None:
    if elemento is None:
        return None
    return next((item for item in elemento.iter() if _nome_local(item.tag) == nome), None)


def _valor(elemento: ET.Element | None, nome: str, descendente: bool = True) -> str:
    item = _descendente(elemento, nome) if descendente else _filho(elemento, nome)
    return _texto(item.text if item is not None else "")


def _mapa_filhos(elemento: ET.Element | None) -> dict[str, str]:
    if elemento is None:
        return {}
    return {_nome_local(item.tag): _texto(item.text) for item in list(elemento)}


def _corrigir_stream_biff(dados: bytes) -> bytes:
    """Corrige, em memória, o offset BOUNDSHEET dos relatórios Domínio.

    Alguns relatórios entregues usam um OLE válido, porém o BOUNDSHEET aponta
    para uma posição anterior ao BOF real da planilha. Excel costuma reparar o
    arquivo silenciosamente; xlrd não. A correção abaixo não grava no original.
    """

    log_silencioso = StringIO()
    stream = bytearray(
        CompDoc(
            dados,
            logfile=log_silencioso,
            ignore_workbook_corruption=True,
        ).get_named_stream("Workbook")
    )
    limites: list[int] = []
    planilhas: list[int] = []
    posicao = 0
    while posicao + 4 <= len(stream):
        opcode = int.from_bytes(stream[posicao : posicao + 2], "little")
        tamanho = int.from_bytes(stream[posicao + 2 : posicao + 4], "little")
        fim = posicao + 4 + tamanho
        if fim > len(stream):
            break
        if opcode == 0x0085:  # BOUNDSHEET
            limites.append(posicao)
        elif opcode == 0x0809 and posicao > 0 and tamanho >= 4:
            tipo_stream = int.from_bytes(stream[posicao + 6 : posicao + 8], "little")
            if tipo_stream == 0x0010:  # worksheet
                planilhas.append(posicao)
        posicao = fim
    if not limites or len(limites) != len(planilhas):
        raise XLRDError("Não foi possível localizar as planilhas no arquivo XLS.")
    for limite, planilha in zip(limites, planilhas):
        stream[limite + 4 : limite + 8] = int(planilha).to_bytes(4, "little")
    return bytes(stream)


def abrir_xls(dados_ou_caminho: bytes | Path | str) -> xlrd.book.Book:
    dados = (
        Path(dados_ou_caminho).read_bytes()
        if isinstance(dados_ou_caminho, (Path, str))
        else dados_ou_caminho
    )
    log_silencioso = StringIO()
    try:
        return xlrd.open_workbook(
            file_contents=dados,
            logfile=log_silencioso,
            ignore_workbook_corruption=True,
        )
    except XLRDError:
        return xlrd.open_workbook(
            file_contents=_corrigir_stream_biff(dados),
            logfile=StringIO(),
            ignore_workbook_corruption=True,
        )


def _sheet_dataframe(book: xlrd.book.Book, nome: str) -> pd.DataFrame:
    sheet = book.sheet_by_name(nome)
    return pd.DataFrame([sheet.row_values(i) for i in range(sheet.nrows)])


def ler_fornecedores(caminho: Path | str) -> pd.DataFrame:
    book = abrir_xls(caminho)
    mapa_simples: dict[str, str] = {}
    if "Planilha1" in book.sheet_names():
        apoio = _sheet_dataframe(book, "Planilha1")
        for _, linha in apoio.iloc[1:].iterrows():
            documento = normalizar_documento(linha.iloc[0] if len(linha) else "")
            if documento:
                mapa_simples[documento] = _texto(linha.iloc[1] if len(linha) > 1 else "")

    nome_sheet = "Fornecedores" if "Fornecedores" in book.sheet_names() else book.sheet_names()[0]
    bruto = _sheet_dataframe(book, nome_sheet)
    registros: list[dict[str, object]] = []
    for _, linha in bruto.iterrows():
        if len(linha) < 10:
            continue
        codigo = _codigo(linha.iloc[0])
        razao = _texto(linha.iloc[6])
        documento = normalizar_documento(linha.iloc[8])
        if not codigo.isdigit() or not razao or not documento:
            continue
        simples_texto = _texto(linha.iloc[9]) or mapa_simples.get(documento, "")
        simples_norm = _sem_acento(simples_texto)
        simples = True if simples_norm == "SIM" else False if simples_norm in {"NAO", "NÃO"} else pd.NA
        registros.append(
            {
                "codigo_fornecedor": codigo,
                "fornecedor_cadastro": razao,
                "documento_fornecedor": documento,
                "documento_formatado": formatar_documento(documento),
                "simples_nacional": simples,
                "simples_informado": simples_texto,
                "regime_apuracao": _texto(linha.iloc[13] if len(linha) > 13 else ""),
            }
        )
    if not registros:
        return pd.DataFrame()
    fornecedores = pd.DataFrame(registros)
    fornecedores = fornecedores.drop_duplicates("codigo_fornecedor", keep="last")
    return fornecedores.reset_index(drop=True)


_PADRAO_FORNECEDOR = re.compile(r"^\s*(\d+)\s*-\s*(.+?)\s*$")


def _separar_fornecedor(valor: object) -> tuple[str, str]:
    match = _PADRAO_FORNECEDOR.match(_texto(valor))
    return (match.group(1), match.group(2)) if match else ("", "")


def ler_contas_pagar(caminho: Path | str) -> pd.DataFrame:
    path = Path(caminho)
    book = abrir_xls(path)
    sheet = book.sheet_by_index(0)
    linhas = [sheet.row_values(i) for i in range(sheet.nrows)]
    empresa_cnpj = normalizar_documento(linhas[1][2] if len(linhas) > 1 and len(linhas[1]) > 2 else "")
    empresa = _texto(linhas[0][2] if linhas and len(linhas[0]) > 2 else "")
    registros: list[dict[str, object]] = []
    fornecedor_grupo = ("", "")
    for linha in linhas:
        linha = list(linha) + [""] * max(0, 22 - len(linha))
        if _sem_acento(linha[0]) == "FORNECEDOR:":
            fornecedor_grupo = _separar_fornecedor(linha[4])
            continue
        situacao = _texto(linha[21])
        if situacao not in {"Aberta", "Paga"}:
            continue
        codigo, fornecedor = _separar_fornecedor(linha[0])
        modo_grupo = not codigo and bool(fornecedor_grupo[0])
        if modo_grupo:
            codigo, fornecedor = fornecedor_grupo
            documento = _texto(linha[4])
            emissao = _data_excel(linha[5])
            entrada = _data_excel(linha[6])
        else:
            documento = _texto(linha[5])
            emissao = _data_excel(linha[6])
            entrada = pd.NaT
        valor = _numero(linha[10])
        if not codigo or valor <= 0:
            continue
        registros.append(
            {
                "arquivo": path.name,
                "empresa": empresa,
                "empresa_cnpj": empresa_cnpj,
                "codigo_fornecedor": codigo,
                "fornecedor": fornecedor,
                "documento": documento,
                "emissao": emissao,
                "entrada": entrada,
                "vencimento": _data_excel(linha[8]),
                "valor_base_credito": valor,
                "valor_pago": _numero(linha[11]),
                "saldo": _numero(linha[19]),
                "situacao": situacao,
            }
        )
    return pd.DataFrame(registros)


def ler_todas_contas_pagar(pasta: Path | str) -> pd.DataFrame:
    quadros = [ler_contas_pagar(path) for path in sorted(Path(pasta).glob("*.xls"))]
    quadros = [quadro for quadro in quadros if not quadro.empty]
    return pd.concat(quadros, ignore_index=True) if quadros else pd.DataFrame()


def _registro_servico_base(**valores: object) -> dict[str, object]:
    base = {
        "arquivo": "",
        "formato": "",
        "empresa_cnpj": "",
        "numero_nota": "",
        "data_emissao": pd.NaT,
        "tomador_documento": "",
        "tomador": "",
        "codigo_servico": "",
        "descricao_servico": "",
        "valor_servico": 0.0,
        "base_calculo": 0.0,
        "base_retencao": 0.0,
        "ir_original": 0.0,
        "csll_original": 0.0,
        "pis_original": 0.0,
        "cofins_original": 0.0,
        "retencao_federal_original": 0.0,
        "ir_destacado": False,
        "csll_destacado": False,
        "situacao_nota": "Ativa",
        "considerar": True,
    }
    base.update(valores)
    return base


def ler_xml_exportacao(caminho: Path | str) -> pd.DataFrame:
    path = Path(caminho)
    root = ET.parse(path).getroot()
    registros: list[dict[str, object]] = []
    for nota in (item for item in root.iter() if _nome_local(item.tag) == "Reg20Item"):
        mapa = _mapa_filhos(nota)
        tributos: dict[str, float] = {}
        for item in (x for x in nota.iter() if _nome_local(x.tag) == "Reg30Item"):
            tributo = _mapa_filhos(item)
            if tributo.get("TributoIdRet") != "1":
                continue
            sigla = _sem_acento(tributo.get("TributoSigla"))
            tributos[sigla] = tributos.get(sigla, 0.0) + _numero(tributo.get("TributoValor"))
        ir = sum(valor for sigla, valor in tributos.items() if sigla in {"IR", "IRRF"})
        csll = sum(valor for sigla, valor in tributos.items() if "CSLL" in sigla)
        pis = sum(valor for sigla, valor in tributos.items() if "PIS" in sigla)
        cofins = sum(valor for sigla, valor in tributos.items() if "COFINS" in sigla)
        valor = _numero(mapa.get("VlNFS"))
        base = _numero(mapa.get("VlBasCalc"), valor)
        ativa = mapa.get("SitNf", "1") in {"", "1", "100"}
        registros.append(
            _registro_servico_base(
                arquivo=path.name,
                formato="XML municipal",
                empresa_cnpj=normalizar_documento(mapa.get("CpfCnpjPre")),
                numero_nota=mapa.get("NumNf", ""),
                data_emissao=pd.to_datetime(mapa.get("DtEmiNf"), dayfirst=True, errors="coerce"),
                tomador_documento=normalizar_documento(mapa.get("CpfCnpjTom")),
                tomador=mapa.get("RazSocTom", ""),
                codigo_servico=mapa.get("CodSrv", ""),
                descricao_servico=mapa.get("DiscrSrv", ""),
                valor_servico=valor,
                base_calculo=base,
                base_retencao=base,
                ir_original=ir,
                csll_original=csll,
                pis_original=pis,
                cofins_original=cofins,
                retencao_federal_original=ir + csll + pis + cofins,
                ir_destacado=ir > 0,
                csll_destacado=csll > 0,
                situacao_nota="Ativa" if ativa else "Cancelada",
                considerar=ativa,
            )
        )
    return pd.DataFrame(registros)


def _ler_nfse_nacional(dados: bytes, arquivo: str) -> dict[str, object]:
    root = ET.fromstring(dados)
    if _nome_local(root.tag) != "NFSe":
        raise ValueError("O XML não é uma NFS-e de produção.")
    inf_nfse = _descendente(root, "infNFSe")
    inf_dps = _descendente(root, "infDPS")
    emitente = _filho(inf_nfse, "emit")
    prestador = _filho(inf_dps, "prest")
    tomador = _filho(inf_dps, "toma")
    valores_dps = _filho(inf_dps, "valores")
    servico = _filho(inf_dps, "serv")
    tributos = _descendente(valores_dps, "tribFed")
    pis_cofins = _descendente(tributos, "piscofins")
    valor = _numero(_valor(valores_dps, "vServ"))
    base_ret = _numero(_valor(pis_cofins, "vBCPisCofins"), valor)
    pis = _numero(_valor(pis_cofins, "vPis"))
    cofins = _numero(_valor(pis_cofins, "vCofins"))
    ir = _numero(_valor(tributos, "vRetIRRF"))
    csll_bruto = _numero(_valor(tributos, "vRetCSLL"))
    # Alguns emissores gravam em vRetCSLL o total PIS+COFINS+CSLL.
    csll = max(csll_bruto - pis - cofins, 0.0) if csll_bruto else 0.0
    valores_nfse = _filho(inf_nfse, "valores")
    total_ret = _numero(_valor(valores_nfse, "vTotalRet"), ir + csll + pis + cofins)
    empresa_cnpj = normalizar_documento(_valor(emitente, "CNPJ") or _valor(prestador, "CNPJ"))
    tomador_doc = normalizar_documento(_valor(tomador, "CNPJ") or _valor(tomador, "CPF"))
    situacao = _valor(inf_nfse, "cStat")
    ativa = situacao in {"", "100"}
    return _registro_servico_base(
        arquivo=arquivo,
        formato="XML NFS-e nacional",
        empresa_cnpj=empresa_cnpj,
        numero_nota=_valor(inf_nfse, "nNFSe"),
        data_emissao=pd.to_datetime(_valor(inf_dps, "dCompet") or _valor(inf_dps, "dhEmi"), errors="coerce"),
        tomador_documento=tomador_doc,
        tomador=_valor(tomador, "xNome"),
        codigo_servico=_valor(servico, "cTribNac"),
        descricao_servico=_valor(servico, "xDescServ"),
        valor_servico=valor,
        base_calculo=valor,
        base_retencao=base_ret or valor,
        ir_original=ir,
        csll_original=csll,
        pis_original=pis,
        cofins_original=cofins,
        retencao_federal_original=total_ret,
        ir_destacado=ir > 0,
        csll_destacado=csll_bruto > 0,
        situacao_nota="Ativa" if ativa else f"Status {situacao}",
        considerar=ativa,
    )


def ler_zip_nfse(caminho: Path | str) -> pd.DataFrame:
    path = Path(caminho)
    registros: list[dict[str, object]] = []
    with zipfile.ZipFile(path) as pacote:
        for info in pacote.infolist():
            if info.is_dir() or not info.filename.lower().endswith(".xml"):
                continue
            try:
                registros.append(_ler_nfse_nacional(pacote.read(info), f"{path.name}::{info.filename}"))
            except (ET.ParseError, ValueError):
                continue
    return pd.DataFrame(registros)


def ler_txt_fixo(caminho: Path | str) -> pd.DataFrame:
    path = Path(caminho)
    dados = path.read_bytes()
    for encoding in ("utf-8-sig", "cp1252", "latin1"):
        try:
            linhas = dados.decode(encoding).splitlines()
            break
        except UnicodeDecodeError:
            continue
    registros: list[dict[str, object]] = []
    for indice, linha in enumerate(linhas, start=1):
        if len(linha) < 171:
            continue
        try:
            monetarios = [int(linha[65 + i * 12 : 77 + i * 12]) / 100 for i in range(7)]
        except ValueError:
            continue
        valor, base, _iss, _iss_ret, ir, _campo6, _campo7 = monetarios
        descricao = re.sub(r"<br\s*/?>", " | ", linha[433:1469], flags=re.I).strip()
        registros.append(
            _registro_servico_base(
                arquivo=path.name,
                formato="TXT fixo municipal",
                empresa_cnpj=normalizar_documento(linha[0:14]),
                numero_nota=linha[22:32].lstrip("0") or "0",
                data_emissao=pd.to_datetime(linha[14:22], format="%Y%m%d", errors="coerce"),
                tomador_documento=normalizar_documento(linha[157:171]),
                tomador=linha[171:271].strip(),
                codigo_servico=linha[40:60].strip(),
                descricao_servico=descricao,
                valor_servico=valor,
                base_calculo=base or valor,
                base_retencao=base or valor,
                ir_original=ir,
                retencao_federal_original=ir,
                ir_destacado=ir > 0,
            )
        )
    return pd.DataFrame(registros)


def ler_servicos(pasta: Path | str) -> pd.DataFrame:
    pasta = Path(pasta)
    quadros: list[pd.DataFrame] = []
    for path in sorted(pasta.iterdir()):
        sufixo = path.suffix.lower()
        if sufixo == ".xml":
            quadros.append(ler_xml_exportacao(path))
        elif sufixo == ".zip":
            quadros.append(ler_zip_nfse(path))
        elif sufixo == ".txt":
            quadros.append(ler_txt_fixo(path))
    quadros = [quadro for quadro in quadros if not quadro.empty]
    return pd.concat(quadros, ignore_index=True) if quadros else pd.DataFrame()


def cruzar_contas_fornecedores(contas: pd.DataFrame, fornecedores: pd.DataFrame) -> pd.DataFrame:
    if contas.empty:
        return contas.copy()
    resultado = contas.merge(fornecedores, how="left", on="codigo_fornecedor")
    if fornecedores.empty:
        return resultado
    por_nome = fornecedores.assign(_nome=fornecedores["fornecedor_cadastro"].map(_sem_acento))
    por_nome = por_nome.drop_duplicates("_nome").set_index("_nome")
    faltantes = resultado["documento_fornecedor"].isna()
    for idx in resultado.index[faltantes]:
        nome = _sem_acento(resultado.at[idx, "fornecedor"])
        if nome in por_nome.index:
            for coluna in [
                "fornecedor_cadastro",
                "documento_fornecedor",
                "documento_formatado",
                "simples_nacional",
                "simples_informado",
                "regime_apuracao",
            ]:
                resultado.at[idx, coluna] = por_nome.at[nome, coluna]
    return resultado


def simular_servicos(servicos: pd.DataFrame, parametros: ParametrosSimulacao) -> pd.DataFrame:
    resultado = servicos.copy()
    if resultado.empty:
        return resultado
    fator = 1 - parametros.reducao_saude_pct / 100
    resultado["aliquota_ibs_efetiva_pct"] = parametros.ibs_pct * fator
    resultado["aliquota_cbs_efetiva_pct"] = parametros.cbs_pct * fator
    resultado["debito_ibs_2027"] = (
        resultado["base_calculo"] * resultado["aliquota_ibs_efetiva_pct"] / 100
    ).round(2)
    resultado["debito_cbs_2027"] = (
        resultado["base_calculo"] * resultado["aliquota_cbs_efetiva_pct"] / 100
    ).round(2)
    resultado["debito_total_2027"] = resultado["debito_ibs_2027"] + resultado["debito_cbs_2027"]
    resultado["ir_2027"] = (
        resultado["base_retencao"]
        * parametros.ir_pct
        / 100
        * resultado["ir_destacado"].astype(float)
    ).round(2)
    resultado["csll_2027"] = (
        resultado["base_retencao"]
        * parametros.csll_pct
        / 100
        * resultado["csll_destacado"].astype(float)
    ).round(2)
    resultado["retencao_federal_2027"] = resultado["ir_2027"] + resultado["csll_2027"]
    resultado["reducao_retencoes"] = (
        resultado["retencao_federal_original"] - resultado["retencao_federal_2027"]
    ).round(2)
    resultado["liquido_apos_retencoes_2027"] = (
        resultado["valor_servico"] - resultado["retencao_federal_2027"]
    ).round(2)
    for coluna in ("empresa_cnpj", "tomador_documento"):
        resultado[f"{coluna}_formatado"] = resultado[coluna].map(formatar_documento)
    return resultado


def simular_creditos(contas: pd.DataFrame, parametros: ParametrosSimulacao) -> pd.DataFrame:
    resultado = contas.copy()
    if resultado.empty:
        return resultado
    tabela = SIMPLES_PRIMEIRA_FAIXA_2027.set_index("anexo")
    if parametros.anexo_simples_padrao not in tabela.index:
        raise ValueError("Anexo do Simples inválido.")
    simples = resultado["simples_nacional"].eq(True)
    regular = resultado["simples_nacional"].eq(False)
    regime_indefinido = resultado["simples_nacional"].isna()
    cnpj_valido = resultado["documento_fornecedor"].fillna("").astype(str).str.len().eq(14)
    sem_cadastro = resultado["documento_fornecedor"].isna() | resultado["documento_fornecedor"].eq("")
    if parametros.indefinido_como_regular:
        regular = regular | regime_indefinido | sem_cadastro
        cnpj_valido = cnpj_valido | sem_cadastro

    resultado["anexo_simples"] = ""
    resultado.loc[simples, "anexo_simples"] = parametros.anexo_simples_padrao
    resultado["aliquota_credito_cbs_pct"] = 0.0
    resultado["aliquota_credito_ibs_pct"] = 0.0
    resultado.loc[regular & cnpj_valido, "aliquota_credito_cbs_pct"] = parametros.cbs_pct
    resultado.loc[regular & cnpj_valido, "aliquota_credito_ibs_pct"] = parametros.ibs_pct
    taxa_simples = tabela.loc[parametros.anexo_simples_padrao]
    resultado.loc[simples & cnpj_valido, "aliquota_credito_cbs_pct"] = taxa_simples["credito_cbs_pct"]
    resultado.loc[simples & cnpj_valido, "aliquota_credito_ibs_pct"] = taxa_simples["credito_ibs_pct"]
    resultado["credito_cbs_2027"] = (
        resultado["valor_base_credito"] * resultado["aliquota_credito_cbs_pct"] / 100
    ).round(2)
    resultado["credito_ibs_2027"] = (
        resultado["valor_base_credito"] * resultado["aliquota_credito_ibs_pct"] / 100
    ).round(2)
    resultado["credito_total_2027"] = resultado["credito_cbs_2027"] + resultado["credito_ibs_2027"]
    resultado["tratamento_credito"] = "Sem crédito — cadastro/regime não identificado"
    resultado.loc[regular & cnpj_valido, "tratamento_credito"] = "Regime regular — alíquota integral"
    resultado.loc[
        parametros.indefinido_como_regular & regime_indefinido & cnpj_valido,
        "tratamento_credito",
    ] = "Regime regular — premissa para cadastro indefinido"
    resultado.loc[simples & cnpj_valido, "tratamento_credito"] = (
        "Simples Nacional — 1ª faixa " + parametros.anexo_simples_padrao
    )
    resultado.loc[~cnpj_valido & ~sem_cadastro, "tratamento_credito"] = "Pessoa física/sem CNPJ — sem crédito presumido"
    return resultado


def resumo_simulacao(servicos: pd.DataFrame, creditos: pd.DataFrame) -> dict[str, float]:
    ativos = servicos[servicos.get("considerar", True).astype(bool)] if not servicos.empty else servicos
    debito_ibs = ativos.get("debito_ibs_2027", pd.Series(dtype=float)).sum()
    debito_cbs = ativos.get("debito_cbs_2027", pd.Series(dtype=float)).sum()
    credito_ibs = creditos.get("credito_ibs_2027", pd.Series(dtype=float)).sum()
    credito_cbs = creditos.get("credito_cbs_2027", pd.Series(dtype=float)).sum()
    return {
        "receita_bruta": ativos.get("valor_servico", pd.Series(dtype=float)).sum(),
        "debitos_ibs": debito_ibs,
        "debitos_cbs": debito_cbs,
        "creditos_ibs": credito_ibs,
        "creditos_cbs": credito_cbs,
        "saldo_ibs": debito_ibs - credito_ibs,
        "saldo_cbs": debito_cbs - credito_cbs,
        "retencoes_originais": ativos.get("retencao_federal_original", pd.Series(dtype=float)).sum(),
        "retencoes_2027": ativos.get("retencao_federal_2027", pd.Series(dtype=float)).sum(),
    }


def carregar_base(pasta_arquivos: Path | str) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    pasta = Path(pasta_arquivos)
    fornecedores = ler_fornecedores(pasta / "Fornecedores.xls")
    pasta_contas = next((p for p in pasta.iterdir() if p.is_dir() and "CONTAS" in _sem_acento(p.name)), None)
    pasta_servicos = next((p for p in pasta.iterdir() if p.is_dir() and "SERVICOS" in _sem_acento(p.name)), None)
    if pasta_contas is None or pasta_servicos is None:
        raise FileNotFoundError("As pastas de contas a pagar e serviços prestados não foram localizadas.")
    contas = cruzar_contas_fornecedores(ler_todas_contas_pagar(pasta_contas), fornecedores)
    servicos = ler_servicos(pasta_servicos)
    return servicos, contas, fornecedores


def gerar_excel(
    servicos: pd.DataFrame,
    creditos: pd.DataFrame,
    fornecedores: pd.DataFrame,
    parametros: ParametrosSimulacao,
) -> bytes:
    resumo = resumo_simulacao(servicos, creditos)
    premissas = pd.DataFrame(
        [
            ("IBS padrão (%)", parametros.ibs_pct),
            ("CBS padrão (%)", parametros.cbs_pct),
            ("Redução saúde (%)", parametros.reducao_saude_pct),
            ("IR retido quando já destacado (%)", parametros.ir_pct),
            ("CSLL retida quando já destacada (%)", parametros.csll_pct),
            ("Anexo padrão do Simples", parametros.anexo_simples_padrao),
            ("Cadastro/regime indefinido tratado como regular", parametros.indefinido_como_regular),
        ],
        columns=["Premissa", "Valor"],
    )
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        pd.DataFrame([resumo]).to_excel(writer, sheet_name="Resumo", index=False)
        servicos.to_excel(writer, sheet_name="Servicos", index=False)
        creditos.to_excel(writer, sheet_name="Creditos", index=False)
        fornecedores.to_excel(writer, sheet_name="Fornecedores", index=False)
        SIMPLES_PRIMEIRA_FAIXA_2027.to_excel(writer, sheet_name="Simples 2027", index=False)
        premissas.to_excel(writer, sheet_name="Premissas", index=False)
        for sheet_name, worksheet in writer.sheets.items():
            worksheet.freeze_panes(1, 0)
            worksheet.autofilter(0, 0, 0, max(0, writer.sheets[sheet_name].dim_colmax))
            worksheet.set_column(0, max(0, worksheet.dim_colmax), 16)
    return buffer.getvalue()
