#!/usr/bin/env python3
"""
PiP Studio - Modulo de prospeccion para Foglia Valvulas
Busca revendedores calificados de valvula de diafragma para bebederos.
Perfil objetivo: negocios fisicos serios con capacidad de asesoramiento tecnico.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Any

import scraper

# ---------------------------------------------------------------------------
# Perfil de revendedor objetivo
# ---------------------------------------------------------------------------
# La valvula de diafragma para bebederos es un insumo ganadero/tambo.
# El revendedor ideal es quien ya le vende al productor agropecuario:
#   - Agropecuarias con mostrador fisico
#   - Cooperativas rurales con seccion de insumos
#   - Veterinarias rurales con venta de insumos de campo
#   - Distribuidoras de materiales para establecimientos ganaderos
#   - Ferreterias de pueblo con rubro agro
#   - Casas de riego y sistemas de agua rural
# ---------------------------------------------------------------------------

FOGLIA_CATEGORIES = (

    scraper.Category(
        "Agropecuaria con mostrador",
        queries=(
            "agropecuaria",
            "casa agropecuaria",
            "insumos agropecuarios",
            "agropecuaria ganadera",
            "venta de insumos rurales",
            "agropecuaria campo",
        ),
    ),

    scraper.Category(
        "Cooperativa rural con insumos",
        queries=(
            "cooperativa agropecuaria",
            "cooperativa rural insumos",
            "cooperativa ganadera",
            "cooperativa tambera",
            "acopio y suministros rurales",
            "cooperativa campo insumos",
        ),
    ),

    scraper.Category(
        "Veterinaria rural con insumos",
        place_type="veterinary_care",
        queries=(
            "veterinaria rural",
            "veterinaria campo",
            "veterinaria ganadera",
            "veterinaria tambos",
            "veterinaria hacienda",
            "clinica veterinaria rural",
        ),
    ),

    scraper.Category(
        "Ferreteria rural y agro",
        place_type="hardware_store",
        queries=(
            "ferreteria agro",
            "ferreteria rural",
            "ferreteria campo",
            "ferreteria y agropecuaria",
            "materiales y agro",
            "ferreteria insumos rurales",
        ),
    ),

    scraper.Category(
        "Distribuidora materiales ganaderos",
        queries=(
            "distribuidora ganadera",
            "insumos para tambos",
            "materiales para hacienda",
            "suministros ganaderos",
            "distribuidor agropecuario",
            "materiales para bebederos",
        ),
    ),

    scraper.Category(
        "Riego y sistemas de agua rural",
        queries=(
            "sistemas de riego",
            "riego rural",
            "bombas y bebederos",
            "instalacion de bebederos",
            "sistemas de agua para campo",
            "riego y bebederos ganaderos",
        ),
    ),

    scraper.Category(
        "Casa de campo y semillas",
        queries=(
            "casa de campo",
            "semillas y campo",
            "almacen rural",
            "venta semillas insumos campo",
            "almacen agropecuario",
        ),
    ),
)


# ---------------------------------------------------------------------------
# Ciudades por provincia
# Foco: zonas ganaderas y tamberas reales
# Ciudades chicas + medianas por igual
# ---------------------------------------------------------------------------

FOGLIA_CITIES: dict[str, tuple[str, ...]] = {

    "Santa Fe": (
        "Sunchales", "Rafaela", "Reconquista", "Venado Tuerto", "Esperanza",
        "San Cristobal", "Galvez", "Las Rosas", "Firmat", "Casilda",
        "Ceres", "Tostado", "Vera", "San Justo", "Morteros",
        "Sastre", "Totoras", "Coronda", "Laguna Paiva", "Rufino",
        "Avellaneda", "Calchaqui", "San Lorenzo", "Canada de Gomez", "Villa Constitucion",
    ),

    "Cordoba": (
        "Rio Cuarto", "Villa Maria", "San Francisco", "Marcos Juarez", "Bell Ville",
        "La Carlota", "Laboulaye", "General Cabrera", "Oncativo", "Morrison",
        "Corral de Bustos", "Leones", "Arroyito", "Jesus Maria", "Villa del Rosario",
        "Hernando", "Villa Nueva", "Buchardo", "General Deheza", "Vicuna Mackenna",
        "Huinca Renanco", "Berrotaran", "Almafuerte", "Noetinger", "Justiniano Posse",
    ),

    "Buenos Aires": (
        "Tandil", "Olavarria", "Azul", "Junin", "Pergamino",
        "Saladillo", "Bolivar", "Rauch", "Las Flores", "Dolores",
        "Chascomus", "Lobos", "Bragado", "Chivilcoy", "Nueve de Julio",
        "Lincoln", "General Viamonte", "Pehuajo", "Trenque Lauquen", "Rivadavia",
        "Daireaux", "Tapalque", "General Belgrano", "Monte", "Roque Perez",
    ),

    "Entre Rios": (
        "Concordia", "Gualeguaychu", "Parana", "Concepcion del Uruguay", "Villaguay",
        "Colon", "Federal", "La Paz", "Crespo", "Nogoya",
        "Victoria", "Diamante", "Gualeguay", "Rosario del Tala", "San Jose",
        "Basavilbaso", "Urdinarrain", "Chajarí", "Viale", "Cerrito",
    ),

    "La Pampa": (
        "Santa Rosa", "General Pico", "Realico", "Quemu Quemu", "Eduardo Castex",
        "Toay", "Intendente Alvear", "General Acha", "Macachín", "Victorica",
        "Winifreda", "Doblas", "Rancul", "Trenel", "Ingeniero Luiggi",
    ),

    "Mendoza": (
        "San Rafael", "Rivadavia", "Junin", "San Martin", "Lujan de Cuyo",
        "Tupungato", "Tunuyan", "La Paz", "Lavalle", "Santa Rosa",
        "General Alvear", "Malargue",
    ),

    "Santiago del Estero": (
        "Santiago del Estero", "La Banda", "Frías", "Añatuya", "Loreto",
        "Fernandez", "Monte Quemado", "Quimili", "Sumampa", "Clodomira",
        "Bandera", "Suncho Corral", "Villa Ojo de Agua",
    ),
}


# ---------------------------------------------------------------------------
# Dataclass resultado Foglia
# ---------------------------------------------------------------------------

@dataclass
class FogliaProspectResult:
    filename: str
    path: Path
    count: int
    province: str
    emails_found: int = 0


# ---------------------------------------------------------------------------
# Funcion principal de prospeccion Foglia
# ---------------------------------------------------------------------------

def prospect_foglia_province(
    province: str,
    *,
    output_dir: Path | None = None,
    verbose: bool = True,
    fetch_emails: bool = True,
    progress=None,
) -> FogliaProspectResult:
    """
    Corre todas las categorias Foglia en todas las ciudades de una provincia.
    Devuelve un Excel consolidado con todos los revendedores potenciales.
    """
    if province not in FOGLIA_CITIES:
        raise ValueError("Provincia no disponible: {}. Opciones: {}".format(
            province, ", ".join(FOGLIA_CITIES.keys())
        ))

    cities      = FOGLIA_CITIES[province]
    all_leads   = []
    n_categories = len(FOGLIA_CATEGORIES)
    n_cities     = len(cities)
    total_queries = sum(len(cat.queries) for cat in FOGLIA_CATEGORIES) * n_cities

    if verbose:
        print("=" * 65)
        print("  Foglia Valvulas - Prospeccion de Revendedores")
        print("=" * 65)
        print("  Provincia  : {}".format(province))
        print("  Ciudades   : {}".format(n_cities))
        print("  Categorias : {}".format(n_categories))
        print("  Queries    : {} en total".format(total_queries))
        print("-" * 65)

    query_count = 0

    for city in cities:
        if verbose:
            print("\n  >> {}".format(city))

        for cat in FOGLIA_CATEGORIES:
            city_leads = scraper.search_places_for_city(
                cat, city,
                progress=progress,
                query_offset=query_count,
                total_queries=total_queries,
            )
            all_leads.extend(city_leads)
            query_count += len(cat.queries)

    # Deduplicar global
    all_leads = scraper.deduplicate(all_leads)

    if verbose:
        print("\n  Total unico: {} negocios".format(len(all_leads)))

    # Enriquecer con emails
    if fetch_emails:
        with_web = sum(1 for l in all_leads if l.get("sitio_web"))
        if verbose:
            print("  Buscando emails en {} sitios web...".format(with_web))
        all_leads = scraper.enrich_with_emails(
            all_leads,
            progress=progress,
            email_offset=total_queries,
            total_emails=total_queries + with_web,
        )

    emails_found = sum(1 for l in all_leads if l.get("email"))

    # Exportar
    target_dir = output_dir or scraper.OUTPUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)

    from datetime import datetime
    ts          = datetime.now().strftime("%Y%m%d_%H%M%S")
    prov_slug   = scraper.slugify(province)
    output_path = target_dir / "foglia_revendedores_{}_{}.xlsx".format(prov_slug, ts)

    if all_leads:
        scraper.export_to_excel(all_leads, output_path)
        if verbose:
            print("\n  Listo -> {}".format(output_path))
            print("  {} revendedores, {} con email".format(len(all_leads), emails_found))
            print("=" * 65)

    if progress:
        progress("Listo! Preparando descarga...", 1, 1)

    return FogliaProspectResult(
        filename=output_path.name,
        path=output_path,
        count=len(all_leads),
        province=province,
        emails_found=emails_found,
    )


def prospect_foglia_city(
    province: str,
    city: str,
    *,
    output_dir: Path | None = None,
    verbose: bool = True,
    fetch_emails: bool = True,
    progress=None,
) -> FogliaProspectResult:
    """
    Corre todas las categorias Foglia en una sola ciudad.
    Util para busquedas rapidas o pruebas.
    """
    scraper.check_api_key()
    all_leads    = []
    total_queries = sum(len(cat.queries) for cat in FOGLIA_CATEGORIES)
    query_count  = 0

    if verbose:
        print("=" * 65)
        print("  Foglia Valvulas - Prospeccion ciudad: {}".format(city))
        print("=" * 65)

    for cat in FOGLIA_CATEGORIES:
        city_leads = scraper.search_places_for_city(
            cat, city,
            progress=progress,
            query_offset=query_count,
            total_queries=total_queries,
        )
        all_leads.extend(city_leads)
        query_count += len(cat.queries)

    all_leads = scraper.deduplicate(all_leads)

    if fetch_emails:
        all_leads = scraper.enrich_with_emails(
            all_leads,
            progress=progress,
            email_offset=total_queries,
            total_emails=total_queries + sum(1 for l in all_leads if l.get("sitio_web")),
        )

    emails_found = sum(1 for l in all_leads if l.get("email"))

    target_dir = output_dir or scraper.OUTPUT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)

    from datetime import datetime
    ts          = datetime.now().strftime("%Y%m%d_%H%M%S")
    city_slug   = scraper.slugify(city)
    output_path = target_dir / "foglia_revendedores_{}_{}.xlsx".format(city_slug, ts)

    if all_leads:
        scraper.export_to_excel(all_leads, output_path)
        if verbose:
            print("\n  {} revendedores encontrados, {} con email.".format(len(all_leads), emails_found))
            print("  Archivo: {}".format(output_path))

    if progress:
        progress("Listo! Preparando descarga...", 1, 1)

    return FogliaProspectResult(
        filename=output_path.name,
        path=output_path,
        count=len(all_leads),
        province=province,
        emails_found=emails_found,
    )


def get_provinces():
    return list(FOGLIA_CITIES.keys())


def get_cities_for_province(province: str):
    return list(FOGLIA_CITIES.get(province, []))


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    import sys
    print("\n--- Foglia Valvulas - Prospeccion de Revendedores ---")
    print("\nProvincias disponibles:\n")
    provinces = get_provinces()
    for i, p in enumerate(provinces, start=1):
        n = len(FOGLIA_CITIES[p])
        print("  {:2}. {} ({} ciudades)".format(i, p, n))

    while True:
        choice = input("\nElegi una provincia (1-{}): ".format(len(provinces))).strip()
        if not choice.isdigit() or not (1 <= int(choice) <= len(provinces)):
            print("Numero invalido.")
            continue
        province = provinces[int(choice) - 1]
        break

    print("\nOpciones:")
    print("  1. Prospectar toda la provincia ({} ciudades)".format(len(FOGLIA_CITIES[province])))
    print("  2. Prospectar una ciudad especifica")

    mode = input("\nElegi (1 o 2): ").strip()

    if mode == "2":
        cities = get_cities_for_province(province)
        print("\nCiudades de {}:\n".format(province))
        for i, c in enumerate(cities, start=1):
            print("  {:2}. {}".format(i, c))
        while True:
            cc = input("\nElegi ciudad (1-{}): ".format(len(cities))).strip()
            if not cc.isdigit() or not (1 <= int(cc) <= len(cities)):
                print("Numero invalido.")
                continue
            city = cities[int(cc) - 1]
            break
        try:
            prospect_foglia_city(province, city, verbose=True)
        except Exception as exc:
            print("Error: {}".format(exc))
            sys.exit(1)
    else:
        try:
            prospect_foglia_province(province, verbose=True)
        except Exception as exc:
            print("Error: {}".format(exc))
            sys.exit(1)


if __name__ == "__main__":
    main()

# ---------------------------------------------------------------------------
# Uso:
#   python foglia.py                    (CLI interactivo)
#   python server.py                    (interfaz web, incluye modo Foglia)
# ---------------------------------------------------------------------------