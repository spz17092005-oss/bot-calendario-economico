import re
from bs4 import BeautifulSoup
from curl_cffi import requests
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

print("=== INICIANDO EXTRACCIÓN CON FILTRO EXACTO DE PAÍSES Y DISEÑO ===")

URL = "https://es.investing.com/economic-calendar/"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "es-ES,es;q=0.9",
}

# -------------------------------------------------------------
# FILTRO DE PAÍSES EXACTO DE TUS CAPTURAS:
# EU, GB, CH, CA, US, AU, CN, JP, NZ
# -------------------------------------------------------------
CODIGOS_PERMITIDOS = [
    "EU",
    "EUR",
    "GB",
    "GBP",
    "UK",
    "CH",
    "CHF",
    "CA",
    "CAD",
    "US",
    "USD",
    "AU",
    "AUD",
    "CN",
    "CNY",
    "JP",
    "JPY",
    "NZ",
    "NZD",
]

MAPA_DIVISAS = {
    "US": "USD",
    "USD": "USD",
    "EU": "EUR",
    "EUR": "EUR",
    "DE": "EUR",
    "FR": "EUR",
    "IT": "EUR",
    "ES": "EUR",
    "GB": "GBP",
    "GBP": "GBP",
    "UK": "GBP",
    "CA": "CAD",
    "CAD": "CAD",
    "AU": "AUD",
    "AUD": "AUD",
    "NZ": "NZD",
    "NZD": "NZD",
    "JP": "JPY",
    "JPY": "JPY",
    "CN": "CNY",
    "CNY": "CNY",
    "CH": "CHF",
    "CHF": "CHF",
}

# Estilos idénticos a la foto del cliente (banners azules y bordes)
FILL_BANNER = PatternFill(
    start_color="1B4F72", end_color="1B4F72", fill_type="solid"
)  # Azul oscuro
FONT_BANNER = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
FONT_HEADER = Font(name="Calibri", size=10, bold=True)
ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
BORDER_THIN = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)

try:
    response = requests.get(
        URL, headers=HEADERS, impersonate="chrome124", timeout=15
    )

    if response.status_code == 200:
        soup = BeautifulSoup(response.text, "html.parser")

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Histórico Eventos"

        # Nota informativa
        ws.merge_cells("B1:G1")
        ws["A1"] = "Nota:"
        ws["A1"].font = Font(bold=True)
        ws["B1"] = (
            "Este archivo constituye la base histórica de indicadores macroeconómicos. "
            "Datos extraídos de investing.com, e intenta recopilar los eventos o indicadores publicados."
        )
        ws["B1"].font = Font(color="595959")

        filas = soup.find_all("tr")
        total_eventos = 0

        for tr in filas:
            classes = tr.get("class", [])

            # Encabezado de FECHA (Banner Azul)
            if "theader" in classes or tr.find(
                "td", class_=lambda c: c and "theader" in c
            ):
                texto_fecha = tr.text.strip()
                if texto_fecha:
                    ws.append([])  # Separador

                    current_row = ws.max_row + 1
                    ws.append([texto_fecha, "", "", "", "", "", ""])
                    ws.merge_cells(
                        start_row=current_row,
                        start_column=1,
                        end_row=current_row,
                        end_column=7,
                    )

                    cell_banner = ws.cell(row=current_row, column=1)
                    cell_banner.fill = FILL_BANNER
                    cell_banner.font = FONT_BANNER
                    cell_banner.alignment = ALIGN_CENTER

                    # Fila Títulos
                    ws.append([
                        "Hora",
                        "Divisa",
                        "Impacto",
                        "Evento",
                        "Actual",
                        "Pronóstico",
                        "Anterior",
                    ])
                    h_row = ws.max_row
                    for c_idx in range(1, 8):
                        c = ws.cell(row=h_row, column=c_idx)
                        c.font = FONT_HEADER
                        c.alignment = ALIGN_CENTER
                        c.border = BORDER_THIN
                continue

            tds = tr.find_all("td")
            if len(tds) >= 6:
                # 1. HORA
                td_hora = tds[0].text.strip()
                if not td_hora or td_hora.lower() == "hora":
                    continue

                # 2. DIVISA / PAÍS Y FILTRO
                td_divisa_raw = tds[1].text.strip().upper()
                flag_span = tds[1].find("span")
                if flag_span and flag_span.get("title"):
                    td_divisa_raw = flag_span.get("title").upper()

                # Normalizar código
                codigo_encontrado = None
                for cod in CODIGOS_PERMITIDOS:
                    if cod in td_divisa_raw or td_divisa_raw in cod:
                        codigo_encontrado = cod
                        break

                # Si el país no está entre los 9 de la foto, se ignora
                if not codigo_encontrado:
                    continue

                divisa_final = MAPA_DIVISAS.get(
                    codigo_encontrado, codigo_encontrado
                )

                # 3. IMPACTO (Estrellas)
                bulls = tds[2].find_all(
                    "i", class_=re.compile(r"grayFullBullishIcon|grayBull")
                )
                num_stars = len(bulls) or tds[2].text.count("⭐") or 1
                impacto = "⭐" * num_stars

                # 4. EVENTO
                evento_raw = tds[3].text.strip()
                evento = re.sub(
                    r"Act:.*", "", evento_raw, flags=re.IGNORECASE
                ).strip()

                # 5, 6, 7. VALORES
                actual = tds[4].text.strip() if len(tds) > 4 else "—"
                pronostico = tds[5].text.strip() if len(tds) > 5 else "—"
                anterior = tds[6].text.strip() if len(tds) > 6 else "—"

                # Guardar en Excel con formato
                ws.append([
                    td_hora,
                    divisa_final,
                    impacto,
                    evento,
                    actual or "—",
                    pronostico or "—",
                    anterior or "—",
                ])
                r_idx = ws.max_row

                for c_idx in range(1, 8):
                    c = ws.cell(row=r_idx, column=c_idx)
                    c.border = BORDER_THIN
                    c.alignment = ALIGN_CENTER if c_idx != 4 else ALIGN_LEFT

                total_eventos += 1

        # Ancho automático de columnas
        for col in ws.columns:
            max_len = max(len(str(cell.value or "")) for cell in col)
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        archivo_salida = "Historico_Eventos_Automatizado.xlsx"
        wb.save(archivo_salida)
        print(
            f"\n¡ÉXITO TOTAL! Se procesaron {total_eventos} eventos para los 9 países filtrados."
        )

    else:
        print(f"Error HTTP: {response.status_code}")

except Exception as e:
    print(f"\nError durante la ejecución: {e}")