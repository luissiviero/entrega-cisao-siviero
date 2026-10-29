"""Check SEI lots in Farroupilha against the municipal APP map (Lei 4.814/2023).

Source: public ArcGIS services of the Farroupilha city hall (GEOFarroupilha):
  cadastro  .../_GEO_Farroupilha_Publico_Cadastro_Imobiliario/MapServer (272 Quadras, 404 Lotes)
  APP map   .../Marco_Regulatorio/MapServer (0 Nascentes, 2 APP_50m, 3 APP_30m,
            4 Faixa_nao_edificante_15m, 5 Faixa_nao_edificante_5m,
            10 Sensibilidade_Ambiental, 11 Rev_MP_APPs)
For every Farroupilha lot of the SEI 2025 inventory (own stock, possible APP
and third-party lots) the script finds the lot by quadra and lot number,
measures the part of the lot inside each APP layer and saves the raw
responses to 01_fontes/geofarroupilha/ (the map is today's version, not the
31.12.2024 one).

Usage: python geofarroupilha_app.py
Writes 02_trabalho/geofarroupilha-app-lotes.xlsx and 01_fontes/geofarroupilha/*.json.
"""

from __future__ import annotations

import datetime as dt
import json
import ssl
from pathlib import Path

import openpyxl
import requests

try:  # the city hall server omits its intermediate certificate; use the Windows store
    import truststore
except ImportError:
    from pip._vendor import truststore


class _WinStore(requests.adapters.HTTPAdapter):
    def init_poolmanager(self, *args, **kwargs):
        kwargs["ssl_context"] = truststore.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
        return super().init_poolmanager(*args, **kwargs)


HTTP = requests.Session()
HTTP.mount("https://", _WinStore())
from shapely.geometry import Point, Polygon
from shapely.ops import unary_union

BASE = Path("C:/entrega-cisao-siviero")
W = BASE / "02_trabalho"
RAW = BASE / "01_fontes" / "geofarroupilha"
OUT = W / "geofarroupilha-app-lotes.xlsx"
SRV = "https://geo.farroupilha.rs.gov.br/server/rest/services/Geo_Farroupilha"
CAD = f"{SRV}/_GEO_Farroupilha_Publico_Cadastro_Imobiliario/MapServer"
APP = f"{SRV}/Marco_Regulatorio/MapServer"
SR = 31982  # SIRGAS 2000 / UTM 22S, metres
# Sector of each development where a block number repeats (checked on the map)
SETOR = {"Jardim Itália I – II – III": "24", "Firminio Dalzóchio": "06"}
RESTRICTING = {2, 3, 4, 5}  # APP 50 m, APP 30 m, non-building strips 15 m and 5 m
WITH_SENS = RESTRICTING | {10}  # plus "Sensibilidade_Ambiental" (30 m strip along all watercourses)
APP_LAYERS = {
    2: "APP 50 m",
    3: "APP 30 m",
    4: "Faixa não edificante 15 m",
    5: "Faixa não edificante 5 m",
    10: "Sensibilidade ambiental",
    11: "Revisão MP (APPs)",
}


def query(url: str, **params) -> dict:
    params = {"f": "json", "outSR": SR, **params}
    r = HTTP.post(f"{url}/query", data=params, timeout=60)
    r.raise_for_status()
    data = r.json()
    if "error" in data:
        raise RuntimeError(f"{url}: {data['error']}")
    return data


def poly(geom: dict) -> Polygon:
    rings = [Polygon(r) for r in geom["rings"]]
    return unary_union(rings).buffer(0)


def esri(p) -> str:
    rings = []
    for g in getattr(p, "geoms", [p]):
        rings.append(list(g.exterior.coords))
        rings += [list(i.coords) for i in g.interiors]
    return json.dumps({"rings": rings, "spatialReference": {"wkid": SR}})


def sei_lots():
    ws = openpyxl.load_workbook(
        W / "sei-2025-itens.xlsx", read_only=True, data_only=True
    )["Lotes"]
    it = ws.iter_rows(values_only=True)
    head = next(it)
    for r in it:
        d = dict(zip(head, r))
        if d["cidade"] == "Farroupilha":
            yield d


def categoria(d) -> str:
    if str(d["bloco"]).startswith("Lotes de terceiros"):
        return "Terceiro indenizável"
    if d["observação"] == "ÁREA DE POSSÍVEL APP":
        return "Próprio, possível APP"
    return "Próprio, valor cheio"


def lot_code(lote: str) -> str:
    s = str(lote).strip().upper().replace("O", "0")
    return f"{int(s):05d}"


def main() -> None:
    RAW.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now(dt.UTC).date().isoformat()
    quadras: dict[str, dict] = {}
    rows = []
    raw_lots = {}
    for d in sei_lots():
        q = str(d["quadra"]).strip()
        q4 = f"{int(q):04d}" if q.isdigit() else q
        status, area_lote, inscr, hits, nasc = "lote não encontrado", None, None, {}, 0
        restr = []
        restr_s = []
        res = query(
            f"{CAD}/404",
            where=f"inscricao LIKE '%.{q4}.{lot_code(d['lote'])}'",
            outFields="cod_lote,inscricao,propriedade",
            returnGeometry="true",
        )
        feats = res["features"]
        if len(feats) > 1:
            # same block number in two sectors: keep the sector of the development
            setor = SETOR.get(d["loteamento"])
            feats = [f for f in feats if f["attributes"]["inscricao"].split(".")[0] == setor]
        quadras[q] = [f["attributes"]["inscricao"] for f in res["features"]]
        if True:
            raw_lots[f"{d['loteamento']}|{q}|{d['lote']}"] = feats
            if len(feats) == 1:
                lp = poly(feats[0]["geometry"])
                area_lote = lp.area
                inscr = feats[0]["attributes"]["inscricao"]
                status = "ok"
                restr = []  # parts of the lot inside APP or non-building strips
                for lid, name in APP_LAYERS.items():
                    a = query(
                        f"{APP}/{lid}",
                        geometry=esri(lp),
                        geometryType="esriGeometryPolygon",
                        spatialRel="esriSpatialRelIntersects",
                        inSR=SR,
                        outFields="*",
                        returnGeometry="true",
                    )
                    parts = [lp.intersection(poly(f["geometry"])) for f in a["features"]]
                    inter = sum(g.area for g in parts)
                    if lid in RESTRICTING:
                        restr += parts
                    if lid in WITH_SENS:
                        restr_s += parts
                    if inter > 0.5:
                        extra = ""
                        if lid == 11:
                            extra = "; ".join(
                                str(f["attributes"].get("processo_adm"))
                                for f in a["features"]
                            )
                        hits[name] = (inter, extra)
                n = query(
                    f"{APP}/0",
                    geometry=esri(lp.buffer(50)),
                    geometryType="esriGeometryPolygon",
                    spatialRel="esriSpatialRelIntersects",
                    inSR=SR,
                    outFields="objectid",
                    returnGeometry="true",
                )
                nasc = sum(
                    1
                    for f in n["features"]
                    if lp.distance(Point(f["geometry"]["x"], f["geometry"]["y"])) <= 50
                )
            else:
                status = f"{len(feats)} lotes com esse código"
        restr_area = unary_union(restr).area if area_lote and restr else 0.0
        restr_s_area = unary_union(restr_s).area if area_lote and restr_s else 0.0
        rows.append(
            (
                categoria(d),
                d["loteamento"],
                q,
                d["lote"],
                d["valor (R$)"],
                inscr,
                status,
                round(area_lote, 1) if area_lote else None,
                *(
                    round(hits[n][0], 1) if n in hits else None
                    for n in APP_LAYERS.values()
                ),
                round(100 * restr_area / area_lote, 1) if area_lote else None,
                round(100 * restr_s_area / area_lote, 1) if area_lote else None,
                nasc,
                hits.get("Revisão MP (APPs)", (0, ""))[1],
            )
        )
    (RAW / f"lotes-sei-{stamp}.json").write_text(
        json.dumps(raw_lots, ensure_ascii=False), encoding="utf-8"
    )
    (RAW / f"quadras-sei-{stamp}.json").write_text(
        json.dumps(quadras, ensure_ascii=False), encoding="utf-8"
    )

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Lotes"
    ws.append(
        [
            "Categoria (doc 11)",
            "Loteamento",
            "Quadra",
            "Lote",
            "Valor doc 11 (R$)",
            "Inscrição municipal",
            "Situação da busca",
            "Área do lote (m²)",
            *(f"{n} (m²)" for n in APP_LAYERS.values()),
            "% do lote em APP ou faixa não edificante (união das camadas 2 a 5)",
            "% do lote em APP, faixa ou sensibilidade (união das camadas 2 a 5 e 10)",
            "Nascentes a até 50 m",
            "Processo (camada Revisão MP)",
        ]
    )
    for r in rows:
        ws.append(list(r))
    ws2 = wb.create_sheet("Fonte")
    ws2.append(["Consulta", stamp])
    ws2.append(["Cadastro", CAD])
    ws2.append(["Mapa do Marco Regulatório (Lei 4.814/2023)", APP])
    ws2.append(
        [
            "Ressalva",
            "Mapa na versão do dia da consulta, não na de 31.12.2024. Áreas calculadas no SIRGAS 2000 / UTM 22S.",
        ]
    )
    wb.save(OUT)
    ok = sum(1 for r in rows if r[6] == "ok")
    print(f"{OUT.name}: {len(rows)} lotes de Farroupilha; {ok} localizados")
    for r in rows:
        if r[0] != "Próprio, valor cheio" or any(r[8:14]):
            print(
                r[0][:12],
                r[1][:14],
                r[2],
                r[3],
                r[6],
                r[7],
                r[8:14],
                r[14],
                r[15],
                r[16],
            )


if __name__ == "__main__":
    main()
