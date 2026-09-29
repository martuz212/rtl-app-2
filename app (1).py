import io
import os
import math
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt

st.set_page_config(page_title="RTL–MC PRECISO PRO", layout="wide")
st.title("🧭 RTL–MC PRECISO PRO")

COLS_PUNTOS = ["CONSECUTIVO", "ORDEN", "X", "Y"]
COLS_LINEAS = ["CONSECUTIVO", "ORDEN", "LONGITUD", "NOM_COLINDANTE", "CARDINALDIAD",
               "OBSERVACIONES", "NPN_COLINDANTE", "FMI_COLINDANTE", "NOMBRE_PREDIO_COL"]

# =========================================================
# FUNCIONES
# =========================================================

def cargar_archivo(file, hoja=None):
    """Lee xlsx o csv (detecta separador , o ; y codificación)."""
    nombre = file.name.lower()
    if nombre.endswith(".xlsx"):
        xl = pd.ExcelFile(file)
        nombres = {n.strip().upper(): n for n in xl.sheet_names}
        df = pd.read_excel(xl, sheet_name=nombres.get(hoja, xl.sheet_names[0]), dtype=str)
    else:
        raw = file.read()
        for enc in ("utf-8-sig", "latin-1"):
            try:
                texto = raw.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        df = pd.read_csv(io.StringIO(texto), dtype=str, sep=None, engine="python")
    df.columns = [str(c).strip().upper() for c in df.columns]
    df = df.dropna(how="all").reset_index(drop=True)
    df["_FILA"] = df.index + 2  # fila real en Excel (fila 1 = encabezado)
    return df


def a_num(v):
    """Convierte texto a número aceptando coma o punto decimal. None si no se puede."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return None
    s = str(v).strip().replace(" ", "")
    if s == "" or s.lower() == "nan":
        return None
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def f(v):
    return f"{v:.2f}".replace(".", ",")


def clasificar_sentido(dx, dy):
    dx_abs, dy_abs = abs(dx), abs(dy)
    total = dx_abs + dy_abs
    if total == 0:
        return "norte"
    px, py = dx_abs / total, dy_abs / total
    tol = 0.02
    if px <= tol:
        return "norte" if dy > 0 else "sur"
    if py <= tol:
        return "este" if dx > 0 else "oeste"
    if dx > 0 and dy > 0:
        return "noreste"
    elif dx > 0 and dy < 0:
        return "sureste"
    elif dx < 0 and dy < 0:
        return "suroeste"
    else:
        return "noroeste"


def lista_filas(filas, max_n=10):
    filas = list(filas)
    txt = ", ".join(str(x) for x in filas[:max_n])
    if len(filas) > max_n:
        txt += f" … (y {len(filas) - max_n} más)"
    return txt


def crear_plantilla():
    """Genera en memoria un Excel de ejemplo con las dos hojas."""
    puntos = pd.DataFrame({
        "CONSECUTIVO": ["001"] * 4,
        "ORDEN": [1, 2, 3, 4],
        "X": ["1000100,50", "1000200,50", "1000200,50", "1000100,50"],
        "Y": ["1200100,00", "1200100,00", "1200000,00", "1200000,00"],
    })
    lineas = pd.DataFrame({
        "CONSECUTIVO": ["001"] * 4,
        "ORDEN": [1, 2, 3, 4],
        "LONGITUD": ["100,0", "100,0", "100,0", "100,0"],
        "NOM_COLINDANTE": ["Predio Los Robles", "Vía veredal", "Predio El Mirador", "Quebrada La Honda"],
        "CARDINALDIAD": ["NORTE", "ESTE", "SUR", "OESTE"],
        "OBSERVACIONES": ["CORRESPONDE", "", "TRASLAPA", ""],
        "NPN_COLINDANTE": ["250000000000000010001000000000", "", "250000000000000010002000000000", ""],
        "FMI_COLINDANTE": ["050-123456", "", "050-654321", ""],
        "NOMBRE_PREDIO_COL": ["Juan Pérez", "Municipio", "María Gómez", "N/A"],
    })
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        puntos.to_excel(w, sheet_name="PUNTOS", index=False)
        lineas.to_excel(w, sheet_name="LINEAS", index=False)
    return puntos, lineas, buf.getvalue()


def validar(df_p, df_l):
    """Devuelve (errores, advertencias) en lenguaje claro. Filas = filas del Excel."""
    errores, avisos = [], []

    # ---- Columnas obligatorias
    for nombre, df, req in (("Tabla de puntos", df_p, COLS_PUNTOS), ("Tabla de líneas", df_l, COLS_LINEAS)):
        faltan = [c for c in req if c not in df.columns]
        if faltan:
            errores.append(
                f"**{nombre}:** faltan las columnas {', '.join(faltan)}. "
                f"Los encabezados deben llamarse exactamente: {', '.join(req)}. "
                f"Columnas encontradas en tu archivo: {', '.join(c for c in df.columns if c != '_FILA')}."
            )
    if errores:
        return errores, avisos

    if df_p.empty:
        errores.append("**Tabla de puntos:** no tiene filas con datos.")
    if df_l.empty:
        errores.append("**Tabla de líneas:** no tiene filas con datos.")
    if errores:
        return errores, avisos

    # ---- CONSECUTIVO vacío
    for nombre, df in (("puntos", df_p), ("líneas", df_l)):
        vacias = df.loc[df["CONSECUTIVO"].isna() | (df["CONSECUTIVO"].astype(str).str.strip() == ""), "_FILA"]
        if len(vacias):
            errores.append(f"**Tabla de {nombre}:** CONSECUTIVO vacío en las filas {lista_filas(vacias)}. Cada fila debe indicar a qué predio pertenece.")
    return errores, avisos


def validar_predio(dp, dl, cons):
    errores, avisos = [], []

    if dl.empty:
        errores.append(f"El CONSECUTIVO **{cons}** existe en la tabla de puntos pero **no aparece en la tabla de líneas**. "
                       "Revisa que se escriba igual en ambos archivos (ojo con ceros a la izquierda y espacios).")
        return errores, avisos

    # ---- PUNTOS: ORDEN
    ordenes = dp["ORDEN"].apply(a_num)
    malos = dp.loc[ordenes.isna() | (ordenes.fillna(0) % 1 != 0), "_FILA"]
    if len(malos):
        errores.append(f"**Puntos:** ORDEN no es un número entero en las filas {lista_filas(malos)}. Use 1, 2, 3, …")
    else:
        if ordenes.duplicated().any():
            dup = sorted(set(ordenes[ordenes.duplicated()].astype(int)))
            errores.append(f"**Puntos:** el ORDEN está repetido ({lista_filas(dup)}) dentro del consecutivo {cons}. Cada punto debe tener un ORDEN único.")

    # ---- PUNTOS: coordenadas
    for col, etiqueta in (("X", "X (Este)"), ("Y", "Y (Norte)")):
        vals = dp[col].apply(a_num)
        malos = dp.loc[vals.isna(), "_FILA"]
        if len(malos):
            errores.append(f"**Puntos:** la columna {etiqueta} está vacía o tiene texto no numérico en las filas {lista_filas(malos)}. "
                           "Ejemplo válido: 1000100,50 o 1000100.50 (sin letras ni unidades).")
    if len(dp) < 3:
        errores.append(f"**Puntos:** el consecutivo {cons} tiene solo {len(dp)} punto(s). Un polígono necesita mínimo 3.")

    # ---- LÍNEAS: ORDEN
    ords_l = dl["ORDEN"].apply(a_num)
    malos = dl.loc[ords_l.isna() | (ords_l.fillna(0) % 1 != 0), "_FILA"]
    if len(malos):
        errores.append(f"**Líneas:** ORDEN no es un número entero en las filas {lista_filas(malos)}.")
    elif ords_l.duplicated().any():
        errores.append(f"**Líneas:** el ORDEN está repetido en el consecutivo {cons}.")

    # ---- LÍNEAS: cantidad
    if len(dl) != len(dp):
        errores.append(f"**Cantidad:** el consecutivo {cons} tiene **{len(dp)} puntos** pero **{len(dl)} líneas**. "
                       "Debe haber una línea por cada punto (incluida la que cierra del último punto al primero).")

    # ---- LÍNEAS: longitud
    lon = dl["LONGITUD"].apply(a_num)
    malos = dl.loc[lon.isna(), "_FILA"]
    if len(malos):
        errores.append(f"**Líneas:** LONGITUD vacía o no numérica en las filas {lista_filas(malos)}. Ejemplo válido: 125,40")
    neg = dl.loc[lon.notna() & (lon <= 0), "_FILA"]
    if len(neg):
        errores.append(f"**Líneas:** LONGITUD debe ser mayor que cero (filas {lista_filas(neg)}).")

    # ---- LÍNEAS: textos obligatorios
    for col in ("NOM_COLINDANTE", "CARDINALDIAD"):
        vac = dl.loc[dl[col].isna() | (dl[col].astype(str).str.strip() == ""), "_FILA"]
        if len(vac):
            errores.append(f"**Líneas:** la columna {col} está vacía en las filas {lista_filas(vac)}.")

    # ---- Avisos (no bloquean)
    obs = dl["OBSERVACIONES"].fillna("").astype(str).str.strip().str.upper()
    raros = dl.loc[~obs.isin(["", "TRASLAPA", "CORRESPONDE"]), "_FILA"]
    if len(raros):
        avisos.append(f"OBSERVACIONES tiene valores distintos de TRASLAPA / CORRESPONDE / vacío en las filas {lista_filas(raros)}; "
                      "en la redacción no se incluirá el Número Predial Nacional para esas líneas.")
    for col in ("NPN_COLINDANTE", "FMI_COLINDANTE", "NOMBRE_PREDIO_COL"):
        vac = dl.loc[dl[col].isna() | (dl[col].astype(str).str.strip() == ""), "_FILA"]
        if len(vac):
            avisos.append(f"{col} está vacío en las filas {lista_filas(vac)}; aparecerá vacío en el texto final.")
    return errores, avisos


def mostrar_errores(errores, avisos):
    if errores:
        st.error("### ❌ No se puede generar la redacción: corrige estos problemas en tus archivos y vuelve a cargarlos")
        for e in errores:
            st.markdown(f"- {e}")
        st.info("📘 Consulta el instructivo (botón de descarga arriba) para ver el formato esperado.")
        st.stop()
    for a in avisos:
        st.warning("⚠️ " + a)

# =========================================================
# AYUDA / DESCARGAS
# =========================================================

with st.expander("📘 Instructivo y plantilla de datos", expanded=False):
    st.write("Descarga el instructivo y la plantilla de ejemplo para preparar tus archivos correctamente.")
    c1, c2 = st.columns(2)
    ruta_doc = os.path.join(os.path.dirname(__file__), "instructivo_RTL_MC.docx")
    if os.path.exists(ruta_doc):
        with open(ruta_doc, "rb") as fh:
            c1.download_button("⬇️ Descargar instructivo (Word)", fh.read(), "instructivo_RTL_MC.docx",
                               "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    else:
        c1.info("Sube `instructivo_RTL_MC.docx` a la misma carpeta de app.py para habilitar la descarga.")
    _pp, _pl, _bytes = crear_plantilla()
    c2.download_button("⬇️ Descargar plantilla Excel de ejemplo", _bytes, "plantilla_RTL_MC.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.markdown("**Tabla de puntos** (ejemplo)")
    st.dataframe(_pp, hide_index=True)
    st.markdown("**Tabla de líneas** (ejemplo)")
    st.dataframe(_pl, hide_index=True)

tol_dist = st.sidebar.number_input("Tolerancia diferencia de distancia (m)", min_value=0.0, value=0.05, step=0.05,
                                   help="Diferencia máxima aceptada entre la longitud de la tabla y la calculada con coordenadas.")

# =========================================================
# CARGA
# =========================================================

puntos_file = st.file_uploader("📌 Tabla de puntos", type=["xlsx", "csv"])
lineas_file = st.file_uploader("📐 Tabla de líneas", type=["xlsx", "csv"])

# =========================================================
# PROCESO
# =========================================================

if puntos_file and lineas_file:

    try:
        df_p = cargar_archivo(puntos_file, "PUNTOS")
        df_l = cargar_archivo(lineas_file, "LINEAS")
    except Exception as e:
        st.error(f"❌ No se pudo leer alguno de los archivos ({e}). Verifica que sea un .xlsx o .csv válido y no esté dañado o protegido con contraseña.")
        st.stop()

    errores, avisos = validar(df_p, df_l)
    mostrar_errores(errores, avisos)

    df_p["CONSECUTIVO"] = df_p["CONSECUTIVO"].astype(str).str.strip()
    df_l["CONSECUTIVO"] = df_l["CONSECUTIVO"].astype(str).str.strip()

    cons = st.selectbox("🔍 CONSECUTIVO", df_p["CONSECUTIVO"].unique())

    df_p = df_p[df_p["CONSECUTIVO"] == cons].copy()
    df_l = df_l[df_l["CONSECUTIVO"] == cons].copy()

    errores, avisos = validar_predio(df_p, df_l, cons)
    mostrar_errores(errores, avisos)

    # ---------------- LIMPIEZA ----------------
    df_p["ORDEN"] = df_p["ORDEN"].apply(a_num).astype(int)
    df_p = df_p.sort_values("ORDEN")
    df_p["NORTE"] = df_p["Y"].apply(a_num)
    df_p["ESTE"] = df_p["X"].apply(a_num)
    df_p["PUNTO"] = df_p["ORDEN"].astype(str).str.zfill(2)

    df_l["ORDEN"] = df_l["ORDEN"].apply(a_num).astype(int)
    df_l = df_l.sort_values("ORDEN")
    df_l["LONGITUD"] = df_l["LONGITUD"].apply(a_num)
    df_l["COL"] = df_l["NOM_COLINDANTE"].astype(str).str.strip()
    for c in ("CARDINALDIAD", "OBSERVACIONES", "NPN_COLINDANTE", "FMI_COLINDANTE", "NOMBRE_PREDIO_COL"):
        df_l[c] = df_l[c].fillna("").astype(str).str.strip()

    puntos = df_p["PUNTO"].tolist()
    coords = {r["PUNTO"]: (r["NORTE"], r["ESTE"]) for _, r in df_p.iterrows()}

    # puntos consecutivos repetidos (distancia 0)
    repetidos = [f"{puntos[i]}→{puntos[(i+1) % len(puntos)]}" for i in range(len(puntos))
                 if coords[puntos[i]] == coords[puntos[(i + 1) % len(puntos)]]]
    if repetidos:
        mostrar_errores([f"**Puntos:** hay puntos consecutivos con las mismas coordenadas ({', '.join(repetidos)}). "
                         "Elimina el duplicado o corrige las coordenadas."], [])

    # =====================================================
    # VISUALIZACIÓN
    # =====================================================
    st.markdown("### 🗺️ Visualización del polígono")

    x = [coords[p][1] for p in puntos]
    y = [coords[p][0] for p in puntos]
    x.append(x[0])
    y.append(y[0])

    fig, ax = plt.subplots()
    ax.plot(x, y, marker='o')
    for i, p in enumerate(puntos):
        ax.text(x[i], y[i], p)
    ax.ticklabel_format(useOffset=False, style="plain")
    st.pyplot(fig)

    # =====================================================
    # TRAMOS
    # =====================================================
    tramos = []
    for i in range(len(puntos)):
        p1 = puntos[i]
        p2 = puntos[(i + 1) % len(puntos)]
        N1, E1 = coords[p1]
        N2, E2 = coords[p2]
        dx = E2 - E1
        dy = N2 - N1

        ang = math.degrees(math.atan2(dx, dy)) % 360
        sentido = clasificar_sentido(dx, dy)
        dist_calc = round(math.sqrt(dy**2 + dx**2), 1)
        dist_tab = df_l.iloc[i]["LONGITUD"]
        dif = round(abs(dist_calc - dist_tab), 1)

        tramos.append({
            "INI": p1, "FIN": p2, "ANGULO": ang, "SENTIDO": sentido,
            "DIST_CALC": dist_calc, "DIST_TAB": dist_tab, "DIF": dif,
            "ESTADO": "✅ OK" if dif <= tol_dist else "❌ ERROR",
            "CARD": df_l.iloc[i]["CARDINALDIAD"],
            "COL": df_l.iloc[i]["COL"],
            "COND": df_l.iloc[i]["OBSERVACIONES"],
            "NPN": df_l.iloc[i]["NPN_COLINDANTE"],
            "FMI": df_l.iloc[i]["FMI_COLINDANTE"],
            "TIT": df_l.iloc[i]["NOMBRE_PREDIO_COL"],
        })

    df_tramos = pd.DataFrame(tramos)

    st.subheader("📐 Tramos técnicos")
    st.dataframe(df_tramos)

    malos = df_tramos[df_tramos["ESTADO"] == "❌ ERROR"]
    if len(malos):
        detalle = "; ".join(f"tramo {r.INI}→{r.FIN}: tabla {r.DIST_TAB} m vs. calculada {r.DIST_CALC} m (dif. {r.DIF} m)"
                            for r in malos.itertuples())
        st.warning("⚠️ **La longitud de la tabla de líneas no coincide con la calculada con las coordenadas** en "
                   f"{len(malos)} tramo(s): {detalle}. La redacción usa la distancia calculada; revisa si las coordenadas "
                   "o las longitudes están mal digitadas, o si el orden de las líneas no sigue el orden de los puntos.")

    # =====================================================
    # BLOQUES
    # =====================================================
    bloques = []
    actual = [df_tramos.iloc[0]]
    for i in range(1, len(df_tramos)):
        t = df_tramos.iloc[i]
        u = actual[-1]
        if (t["CARD"] == u["CARD"] and t["COL"] == u["COL"] and
                str(t["NPN"]).strip() == str(u["NPN"]).strip() and
                str(t["FMI"]).strip() == str(u["FMI"]).strip()):
            actual.append(t)
        else:
            bloques.append(actual)
            actual = [t]
    bloques.append(actual)

    # =====================================================
    # RTL FINAL
    # =====================================================
    salida = "LINDEROS TÉCNICOS\n\n"
    orden = df_p["PUNTO"].tolist()
    card_actual = None
    contador_lindero = 1

    for b in bloques:
        card = b[0]["CARD"]
        if card != card_actual:
            salida += f"POR EL {card}:\n\n"
            card_actual = card

        salida += f"Lindero {contador_lindero}:\n"

        segmentos = []
        segmento = [b[0]]
        for t in b[1:]:
            if t["SENTIDO"] == segmento[-1]["SENTIDO"]:
                segmento.append(t)
            else:
                segmentos.append(segmento)
                segmento = [t]
        segmentos.append(segmento)

        primera = True
        for segmento in segmentos:
            p_ini = segmento[0]["INI"]
            p_fin = segmento[-1]["FIN"]
            i1 = orden.index(p_ini)
            i2 = orden.index(p_fin)
            inter = orden[i1 + 1:i2] if i2 > i1 else orden[i1 + 1:] + orden[:i2]

            texto_int = ""
            if len(inter) == 1:
                p = inter[0]
                N, E = coords[p]
                texto_int = f"pasando por el punto de coordenadas punto {p} N= {f(N)} m, E= {f(E)} m, "
            elif len(inter) > 1:
                texto_int = "pasando por los puntos de coordenadas "
                for p in inter:
                    N, E = coords[p]
                    texto_int += f"punto {p} N= {f(N)} m, E= {f(E)} m, "
            texto_int = texto_int.rstrip(", ") + ", " if texto_int else ""

            dist = round(sum(s["DIST_CALC"] for s in segmento), 1)
            dist_txt = str(dist).replace(".", ",")
            tipo = "recta" if len(segmento) == 1 else "quebrada"
            sentido = segmento[0]["SENTIDO"]
            N_ini, E_ini = coords[p_ini]
            N_fin, E_fin = coords[p_fin]

            verbo = "Inicia" if primera else "Continúa"
            salida += (
                f"{verbo} en el punto {p_ini} con coordenadas planas N= {f(N_ini)} m, E= {f(E_ini)} m, "
                f"en línea {tipo} en sentido {sentido}, "
                f"{texto_int}"
                f"en una distancia de {dist_txt} m, hasta encontrar el punto {p_fin} "
                f"con coordenadas planas N= {f(N_fin)} m, E= {f(E_fin)} m.\n"
            )
            primera = False

        fila = b[-1]
        salida += f"; colindando con {fila['COL']}"
        if str(fila["COND"]).upper() == "TRASLAPA":
            salida += f", que traslapa con el Número Predial Nacional {fila['NPN']}"
        elif str(fila["COND"]).upper() == "CORRESPONDE":
            salida += f", que corresponde con el Número Predial Nacional {fila['NPN']}"
        salida += f", Folio de matrícula inmobiliaria {fila['FMI']}"
        salida += f" y catastralmente a nombre de {fila['TIT']}.\n\n"
        contador_lindero += 1

    salida = salida.strip() + " y encierra"

    st.text_area("📄 RTL FINAL", salida, height=600)
    st.download_button("⬇️ Descargar redacción (.txt)", salida.encode("utf-8"), f"RTL_{cons}.txt", "text/plain")
