import re
from io import StringIO
from bs4 import BeautifulSoup
from curl_cffi import requests
import openpyxl

print("=== INICIANDO EXTRACCIÓN Y AUTOMATIZACIÓN DE CALENDARIO ===")

URL = "https://es.investing.com/economic-calendar/"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9",
}

# Diccionario para mapear código de divisa/país a Bandera oficial
MAPA_DIVISAS = {
    "US": "USD 🇺🇸",
    "USD": "USD 🇺🇸",
    "EUR": "EUR 🇪🇺",
    "DE": "EUR 🇩🇪",
    "FR": "EUR 🇫🇷",
    "IT": "EUR 🇮🇹",
    "ES": "EUR 🇪🇸",
    "GBP": "GBP 🇬🇧",
    "UK": "GBP 🇬🇧",
    "CAD": "CAD 🇨🇦",
    "AUD": "AUD 🇦🇺",
    "NZD": "NZD 🇳🇿",
    "JPY": "JPY 🇯🇵",
    "JP": "JPY 🇯🇵",
    "CNY": "CNY 🇨🇳",
    "CN": "CNY 🇨🇳",
    "CHF": "CHF 🇨🇭",
    "MXN": "MXN 🇲🇽",
    "BR": "BRL 🇧🇷",
}

try:
    print("1. Conectando con Investing.com (TLS Impersonate)...")
    response = requests.get(
        URL, headers=HEADERS, impersonate="chrome124", timeout=15
    )

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")

        # Crear libro de Excel estructurado
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Histórico Eventos"

        # Encabezado informativo exacto de tu plantilla
        ws.append([
            "Nota:",
            "Este archivo constituye la base histórica de indicadores macroeconómicos. Datos extraídos de investing.com, e intenta recopilar los eventos o indicadores publicados.",
        ])
        ws.append([])

        filas_tabla = soup.find_all("tr")
        total_eventos = 0

        for tr in filas_tabla:
            classes = tr.get("class", [])

            # Detectar encabezado de FECHA (ej. Lunes, 14 de septiembre)
            if "theader" in classes or tr.find(
                "td", class_=lambda c: c and "theader" in c
            ):
                texto_fecha = tr.text.strip()
                if texto_fecha:
                    ws.append([])
                    ws.append([texto_fecha])  # Fila con el día
                    ws.append([
                        "Hora",
                        "Divisa",
                        "Impacto",
                        "Evento",
                        "Actual",
                        "Pronóstico",
                        "Anterior",
                    ])
                continue

            # Detectar filas de datos de eventos
            tds = tr.find_all("td")
            if len(tds) >= 6:
                # 1. HORA
                hora = tds[0].text.strip()
                if not hora or hora.lower() == "hora":
                    continue

                # 2. DIVISA
                divisa_raw = tds[1].text.strip()
                flag_span = tds[1].find("span")
                if flag_span and flag_span.get("title"):
                    divisa_raw = flag_span.get("title").upper()
                divisa = MAPA_DIVISAS.get(divisa_raw, divisa_raw)

                # 3. IMPACTO (Estrellas)
                bulls = tds[2].find_all(
                    "i", class_=re.compile(r"grayFullBullishIcon|grayBull")
                )
                num_stars = len(bulls) or tds[2].text.count("⭐") or 1
                impacto = "⭐" * num_stars

                # 4. EVENTO (Limpio de sufijos de la web)
                evento_raw = tds[3].text.strip()
                evento = re.sub(
                    r"Act:.*", "", evento_raw, flags=re.IGNORECASE
                ).strip()

                # 5, 6, 7. VALORES
                actual = tds[4].text.strip() if len(tds) > 4 else "—"
                pronostico = tds[5].text.strip() if len(tds) > 5 else "—"
                anterior = tds[6].text.strip() if len(tds) > 6 else "—"

                # Guardar fila formateada
                ws.append([
                    hora,
                    divisa,
                    impacto,
                    evento,
                    actual or "—",
                    pronostico or "—",
                    anterior or "—",
                ])
                total_eventos += 1

        archivo_salida = "Historico_Eventos_Automatizado.xlsx"
        wb.save(archivo_salida)
        print(
            f"\n¡ÉXITO! Se guardaron {total_eventos} eventos formateados en '{archivo_salida}'."
        )

    else:
        print(f"Error HTTP: {response.status_code}")

except Exception as e:
    print(f"Error en la ejecución: {e}")