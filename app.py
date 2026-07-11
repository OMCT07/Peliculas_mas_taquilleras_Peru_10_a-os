from io import BytesIO
from pathlib import Path
import re

import pandas as pd
import plotly.express as px
import streamlit as st


# Configuración general de la página


st.set_page_config(
    page_title="Taquilla de cine en Perú",
    page_icon="🎬",
    layout="wide",
    initial_sidebar_state="expanded",
)


# Rutas y datos básicos del proyecto


CARPETA_PROYECTO = Path(__file__).resolve().parent
CARPETA_DATOS = CARPETA_PROYECTO / "data"

ANIOS = list(range(2015, 2026))

COLUMNAS_ESPERADAS = [
    "Título",
    "Recaudación",
    "Dirección",
    "Productora",
    "País",
]


# Apariencia


st.markdown(
    """
    <style>
        .block-container {
            max-width: 1400px;
            padding-top: 2rem;
            padding-bottom: 3rem;
        }

        [data-testid="stSidebar"] {
            background: #f6f3ee;
            border-right: 1px solid #e4ddd3;
        }

        .portada {
            padding: 2.2rem 2.4rem;
            margin-bottom: 1.5rem;
            border-radius: 22px;
            background:
                linear-gradient(
                    120deg,
                    rgba(40, 25, 20, 0.96),
                    rgba(116, 43, 34, 0.90)
                );
            color: white;
            box-shadow: 0 14px 34px rgba(55, 30, 20, 0.18);
        }

        .portada h1 {
            font-size: 2.7rem;
            margin: 0;
            line-height: 1.1;
        }

        .portada p {
            max-width: 760px;
            margin-top: 0.8rem;
            margin-bottom: 0;
            font-size: 1.05rem;
            color: rgba(255, 255, 255, 0.86);
        }

        .seccion-titulo {
            margin-top: 0.4rem;
            margin-bottom: 1rem;
        }

        div[data-testid="stMetric"] {
            padding: 1rem;
            border: 1px solid #e8e2da;
            border-radius: 16px;
            background: white;
            box-shadow: 0 5px 16px rgba(60, 40, 30, 0.06);
        }

        div[data-testid="stDataFrame"] {
            border: 1px solid #e8e2da;
            border-radius: 14px;
            overflow: hidden;
        }

        .ficha {
            padding: 1.5rem;
            border: 1px solid #e7dfd5;
            border-radius: 18px;
            background: #fffdf9;
        }

        .ficha h2 {
            margin-top: 0;
            color: #6f2923;
        }

        .ficha p {
            margin: 0.55rem 0;
            font-size: 1rem;
        }

        .aviso {
            padding: 1rem 1.2rem;
            margin: 1rem 0;
            border-left: 5px solid #bc7a27;
            border-radius: 8px;
            background: #fff6e6;
        }

        .pie {
            text-align: center;
            margin-top: 3rem;
            color: #7a746d;
            font-size: 0.85rem;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


# Funciones para limpiar y ordenar los datos


def limpiar_texto(valor) -> str:
    """
    Elimina espacios sobrantes.

    Algunos Excel pueden traer celdas vacías o textos con espacios
    adicionales. Esta función deja todos esos valores en un formato
    más uniforme.
    """
    if pd.isna(valor):
        return ""

    return re.sub(r"\s+", " ", str(valor).strip())


def limpiar_recaudacion(valor):
    """
    Convierte la recaudación en un número.

    Los archivos pueden contener valores como:
    $5,400,000
    US$ 5,400,000
    5400000

    Para sumar, ordenar y graficar necesitamos que el resultado sea
    numérico y no una cadena de texto.
    """
    if pd.isna(valor):
        return None

    if isinstance(valor, (int, float)):
        return float(valor)

    texto = str(valor).strip()

    if not texto:
        return None

    texto = (
        texto.replace("US$", "")
        .replace("USD", "")
        .replace("$", "")
        .replace("S/", "")
        .replace(" ", "")
    )

    texto = re.sub(r"[^0-9,.\-]", "", texto)

    if not texto:
        return None

    # En los archivos actuales la coma se utiliza principalmente
    # como separador de miles.
    if texto.count(",") >= 1 and "." not in texto:
        texto = texto.replace(",", "")

    elif "," in texto and "." in texto:
        texto = texto.replace(",", "")

    try:
        return float(texto)
    except ValueError:
        return None


def nombre_del_archivo(anio: int) -> str:
    """Construye el nombre exacto que tienen los Excel."""
    return f"Peliculas_mas_taquilleras_Peru_{anio}.xlsx"


@st.cache_data(show_spinner=False)
def cargar_anio(anio: int) -> pd.DataFrame:
    """
    Lee el Excel de un año y prepara sus columnas.

    El caché evita que Streamlit vuelva a leer el mismo archivo cada vez
    que el usuario cambia un filtro.
    """
    ruta = CARPETA_DATOS / nombre_del_archivo(anio)

    if not ruta.exists():
        raise FileNotFoundError(
            f"No se encontró {nombre_del_archivo(anio)} dentro de data."
        )

    df = pd.read_excel(ruta, engine="openpyxl")

    # Quitamos espacios accidentales de los encabezados.
    df.columns = [str(columna).strip() for columna in df.columns]

    faltantes = [
        columna
        for columna in COLUMNAS_ESPERADAS
        if columna not in df.columns
    ]

    if faltantes:
        raise ValueError(
            f"Al archivo de {anio} le faltan las columnas: "
            + ", ".join(faltantes)
        )

    df = df[COLUMNAS_ESPERADAS].copy()

    for columna in ["Título", "Dirección", "Productora", "País"]:
        df[columna] = df[columna].apply(limpiar_texto)

    df["Recaudación"] = df["Recaudación"].apply(limpiar_recaudacion)

    # Una fila sin título no aporta información útil.
    df = df[df["Título"] != ""].copy()

    # El puesto refleja el orden original del Excel.
    df.insert(0, "Puesto", range(1, len(df) + 1))
    df.insert(1, "Año", anio)

    return df.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def cargar_base_completa():
    """
    Une todos los años disponibles.

    También guarda una lista de errores para que un archivo defectuoso
    no impida abrir toda la aplicación.
    """
    tablas = []
    errores = []

    for anio in ANIOS:
        try:
            tablas.append(cargar_anio(anio))
        except Exception as error:
            errores.append(f"{anio}: {error}")

    if not tablas:
        return pd.DataFrame(), errores

    return pd.concat(tablas, ignore_index=True), errores


def mostrar_moneda(valor) -> str:
    """Presenta la recaudación de una manera más fácil de leer."""
    if pd.isna(valor):
        return "Sin dato"

    return f"US$ {valor:,.0f}"


def convertir_a_excel(df: pd.DataFrame) -> bytes:
    """
    Crea temporalmente un Excel para el botón de descarga.

    El archivo se mantiene en memoria; no modifica los Excel originales.
    """
    salida = BytesIO()

    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Películas")

        hoja = writer.sheets["Películas"]
        hoja.freeze_panes = "A2"
        hoja.auto_filter.ref = hoja.dimensions

        anchos = {
            "A": 10,
            "B": 10,
            "C": 38,
            "D": 18,
            "E": 34,
            "F": 38,
            "G": 28,
        }

        for columna, ancho in anchos.items():
            hoja.column_dimensions[columna].width = ancho

    salida.seek(0)
    return salida.getvalue()


def filtrar_texto(df: pd.DataFrame, consulta: str) -> pd.DataFrame:
    """
    Busca una palabra en título, dirección, productora o país.

    re.escape evita problemas cuando el usuario escribe símbolos
    especiales en el buscador.
    """
    if not consulta.strip():
        return df

    patron = re.escape(consulta.strip())

    mascara = (
        df["Título"].str.contains(patron, case=False, na=False)
        | df["Dirección"].str.contains(patron, case=False, na=False)
        | df["Productora"].str.contains(patron, case=False, na=False)
        | df["País"].str.contains(patron, case=False, na=False)
    )

    return df[mascara]


# Portada


st.markdown(
    """
    <div class="portada">
        <h1>🎬 Taquilla de cine en Perú</h1>
        <p>
            Una mirada a las películas con mayor recaudación registradas
            entre 2015 y 2025. Explora cada año, compara resultados y
            descarga la información que necesites.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)


# Carga inicial


with st.spinner("Preparando la información..."):
    base_completa, errores = cargar_base_completa()

if base_completa.empty:

    st.error("No fue posible cargar los archivos.")

    st.write("Ruta donde la aplicación está buscando los datos:")
    st.code(str(CARPETA_DATOS))

    st.write("¿Existe la carpeta?")
    st.write(CARPETA_DATOS.exists())

    if CARPETA_DATOS.exists():

        st.write("Archivos encontrados:")

        for archivo in CARPETA_DATOS.iterdir():
            st.write(archivo.name)

    st.write("Errores detectados:")

    if errores:
        for error in errores:
            st.error(error)

    st.stop()

    if errores:
        for error in errores:
            st.write(error)

    st.stop()

if errores:
    st.warning(
        "Algunos archivos presentan problemas."
    )

    with st.expander("Ver detalles"):
        for error in errores:
            st.write(f"• {error}")


anios_disponibles = sorted(base_completa["Año"].unique())


# Navegación lateral


st.sidebar.header("Explorar")

seccion = st.sidebar.radio(
    "¿Qué deseas revisar?",
    [
        "Ranking por año",
        "Comparación histórica",
        "Base completa",
        "Acerca de los datos",
    ],
)

st.sidebar.divider()


# SECCIÓN 1: RANKING POR AÑO


if seccion == "Ranking por año":

    anio = st.sidebar.selectbox(
        "Año",
        anios_disponibles,
        index=len(anios_disponibles) - 1,
    )

    df_anio = base_completa[
        base_completa["Año"] == anio
    ].copy()

    busqueda = st.sidebar.text_input(
        "Buscar",
        placeholder="Escribe un título, país o director",
    )

    paises = sorted(
        pais
        for pais in df_anio["País"].dropna().unique()
        if pais
    )

    paises_elegidos = st.sidebar.multiselect(
        "País de producción",
        paises,
    )

    solo_con_recaudacion = st.sidebar.checkbox(
        "Ocultar películas sin recaudación",
        value=False,
    )

    orden = st.sidebar.selectbox(
        "Orden",
        [
            "Ranking original",
            "Mayor recaudación",
            "Menor recaudación",
            "Título de A a Z",
        ],
    )

    resultados = filtrar_texto(df_anio, busqueda)

    if paises_elegidos:
        resultados = resultados[
            resultados["País"].isin(paises_elegidos)
        ]

    if solo_con_recaudacion:
        resultados = resultados[
            resultados["Recaudación"].notna()
            & (resultados["Recaudación"] > 0)
        ]

    if orden == "Mayor recaudación":
        resultados = resultados.sort_values(
            "Recaudación",
            ascending=False,
            na_position="last",
        )

    elif orden == "Menor recaudación":
        resultados = resultados.sort_values(
            "Recaudación",
            ascending=True,
            na_position="last",
        )

    elif orden == "Título de A a Z":
        resultados = resultados.sort_values("Título")

    else:
        resultados = resultados.sort_values("Puesto")

    resultados = resultados.reset_index(drop=True)

    st.markdown(
        f'<h2 class="seccion-titulo">Películas de {anio}</h2>',
        unsafe_allow_html=True,
    )

    datos_monetarios = resultados[
        resultados["Recaudación"].notna()
        & (resultados["Recaudación"] > 0)
    ]

    recaudacion_total = datos_monetarios["Recaudación"].sum()
    recaudacion_promedio = datos_monetarios["Recaudación"].mean()

    if not datos_monetarios.empty:
        pelicula_lider = datos_monetarios.loc[
            datos_monetarios["Recaudación"].idxmax()
        ]
    else:
        pelicula_lider = None

    metrica1, metrica2, metrica3, metrica4 = st.columns(4)

    metrica1.metric(
        "Películas encontradas",
        len(resultados),
    )

    metrica2.metric(
        "Recaudación acumulada",
        mostrar_moneda(recaudacion_total),
    )

    metrica3.metric(
        "Promedio",
        mostrar_moneda(recaudacion_promedio),
    )

    metrica4.metric(
        "Película líder",
        pelicula_lider["Título"]
        if pelicula_lider is not None
        else "Sin información",
    )

    if resultados["Recaudación"].isna().any():
        st.markdown(
            """
            <div class="aviso">
                Algunas películas no tienen una cifra de recaudación
                disponible. Esas filas se mantienen en la tabla, pero no
                participan en las sumas ni en los gráficos monetarios.
            </div>
            """,
            unsafe_allow_html=True,
        )

    tabla, graficos, detalle = st.tabs(
        [
            "Tabla",
            "Gráficos",
            "Detalle de una película",
        ]
    )

    with tabla:

        columnas_visibles = [
            "Puesto",
            "Título",
            "Recaudación",
            "Dirección",
            "Productora",
            "País",
        ]

        st.dataframe(
            resultados[columnas_visibles],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Puesto": st.column_config.NumberColumn(
                    "Puesto",
                    format="%d",
                    width="small",
                ),
                "Título": st.column_config.TextColumn(
                    "Título",
                    width="large",
                ),
                "Recaudación": st.column_config.NumberColumn(
                    "Recaudación",
                    format="US$ %.0f",
                    width="medium",
                ),
                "Dirección": st.column_config.TextColumn(
                    "Dirección",
                    width="medium",
                ),
                "Productora": st.column_config.TextColumn(
                    "Productora",
                    width="large",
                ),
                "País": st.column_config.TextColumn(
                    "País",
                    width="medium",
                ),
            },
        )

        datos_para_descargar = resultados[
            [
                "Puesto",
                "Año",
                "Título",
                "Recaudación",
                "Dirección",
                "Productora",
                "País",
            ]
        ]

        descarga1, descarga2 = st.columns(2)

        descarga1.download_button(
            "Descargar resultados en Excel",
            data=convertir_a_excel(datos_para_descargar),
            file_name=f"taquilla_peru_{anio}.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            use_container_width=True,
        )

        descarga2.download_button(
            "Descargar resultados en CSV",
            data=datos_para_descargar.to_csv(
                index=False
            ).encode("utf-8-sig"),
            file_name=f"taquilla_peru_{anio}.csv",
            mime="text/csv",
            use_container_width=True,
        )

    with graficos:

        datos_grafico = (
            resultados[
                resultados["Recaudación"].notna()
                & (resultados["Recaudación"] > 0)
            ]
            .sort_values("Recaudación", ascending=True)
        )

        if datos_grafico.empty:
            st.info(
                "No hay cifras disponibles para construir el gráfico."
            )

        else:
            figura = px.bar(
                datos_grafico,
                x="Recaudación",
                y="Título",
                orientation="h",
                title=f"Recaudación registrada en {anio}",
                hover_data=[
                    "Dirección",
                    "Productora",
                    "País",
                ],
                labels={
                    "Recaudación": "Recaudación en dólares",
                    "Título": "",
                },
            )

            figura.update_layout(
                height=max(500, len(datos_grafico) * 38),
                margin=dict(l=20, r=20, t=70, b=20),
                xaxis_tickprefix="US$ ",
                xaxis_tickformat=",",
                plot_bgcolor="rgba(0,0,0,0)",
            )

            st.plotly_chart(
                figura,
                use_container_width=True,
                config={"displaylogo": False},
            )

        conteo_paises = (
            resultados["País"]
            .replace("", pd.NA)
            .dropna()
            .value_counts()
            .rename_axis("País")
            .reset_index(name="Películas")
        )

        if not conteo_paises.empty:
            figura_paises = px.bar(
                conteo_paises.sort_values(
                    "Películas",
                    ascending=True,
                ),
                x="Películas",
                y="País",
                orientation="h",
                title="Presencia de países en el ranking",
                text="Películas",
            )

            figura_paises.update_layout(
                height=450,
                margin=dict(l=20, r=20, t=70, b=20),
                plot_bgcolor="rgba(0,0,0,0)",
            )

            st.plotly_chart(
                figura_paises,
                use_container_width=True,
                config={"displaylogo": False},
            )

    with detalle:

        if resultados.empty:
            st.info(
                "No hay películas que coincidan con los filtros."
            )

        else:
            titulo_elegido = st.selectbox(
                "Selecciona una película",
                resultados["Título"].tolist(),
            )

            pelicula = resultados[
                resultados["Título"] == titulo_elegido
            ].iloc[0]

            st.markdown(
                f"""
                <div class="ficha">
                    <h2>{pelicula["Título"]}</h2>
                    <p><strong>Puesto:</strong> {pelicula["Puesto"]}</p>
                    <p><strong>Año:</strong> {pelicula["Año"]}</p>
                    <p>
                        <strong>Recaudación:</strong>
                        {mostrar_moneda(pelicula["Recaudación"])}
                    </p>
                    <p>
                        <strong>Dirección:</strong>
                        {pelicula["Dirección"] or "Sin información"}
                    </p>
                    <p>
                        <strong>Productora:</strong>
                        {pelicula["Productora"] or "Sin información"}
                    </p>
                    <p>
                        <strong>País:</strong>
                        {pelicula["País"] or "Sin información"}
                    </p>
                </div>
                """,
                unsafe_allow_html=True,
            )


# SECCIÓN 2: COMPARACIÓN HISTÓRICA


elif seccion == "Comparación histórica":

    st.markdown(
        '<h2 class="seccion-titulo">Comparación entre años</h2>',
        unsafe_allow_html=True,
    )

    anios_seleccionados = st.sidebar.multiselect(
        "Años incluidos",
        anios_disponibles,
        default=list(anios_disponibles),
    )

    if not anios_seleccionados:
        st.info("Selecciona al menos un año.")
        st.stop()

    historico = base_completa[
        base_completa["Año"].isin(anios_seleccionados)
    ].copy()

    historico_valido = historico[
        historico["Recaudación"].notna()
        & (historico["Recaudación"] > 0)
    ]

    resumen = (
        historico_valido
        .groupby("Año", as_index=False)
        .agg(
            Recaudación_total=("Recaudación", "sum"),
            Recaudación_promedio=("Recaudación", "mean"),
            Películas_con_dato=("Recaudación", "count"),
        )
    )

    columna1, columna2, columna3 = st.columns(3)

    columna1.metric(
        "Registros analizados",
        len(historico),
    )

    columna2.metric(
        "Recaudación acumulada",
        mostrar_moneda(historico_valido["Recaudación"].sum()),
    )

    columna3.metric(
        "Películas diferentes",
        historico["Título"].nunique(),
    )

    if not resumen.empty:
        figura_historica = px.line(
            resumen,
            x="Año",
            y="Recaudación_total",
            markers=True,
            title="Recaudación registrada por año",
            labels={
                "Recaudación_total": "Recaudación total",
                "Año": "",
            },
        )

        figura_historica.update_layout(
            height=500,
            xaxis=dict(dtick=1),
            yaxis_tickprefix="US$ ",
            yaxis_tickformat=",",
            plot_bgcolor="rgba(0,0,0,0)",
        )

        st.plotly_chart(
            figura_historica,
            use_container_width=True,
            config={"displaylogo": False},
        )

    st.subheader("Películas con mayor recaudación de toda la base")

    cantidad = st.slider(
        "Cantidad de películas",
        min_value=5,
        max_value=40,
        value=20,
    )

    top = (
        historico_valido
        .sort_values("Recaudación", ascending=False)
        .head(cantidad)
        .copy()
    )

    st.dataframe(
        top[
            [
                "Año",
                "Título",
                "Recaudación",
                "Dirección",
                "Productora",
                "País",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Recaudación": st.column_config.NumberColumn(
                "Recaudación",
                format="US$ %.0f",
            )
        },
    )


# SECCIÓN 3: BASE COMPLETA


elif seccion == "Base completa":

    st.markdown(
        '<h2 class="seccion-titulo">Base consolidada</h2>',
        unsafe_allow_html=True,
    )

    busqueda_general = st.sidebar.text_input(
        "Buscar en todos los años",
        placeholder="Película, director, país...",
    )

    anios_base = st.sidebar.multiselect(
        "Años",
        anios_disponibles,
        default=list(anios_disponibles),
    )

    base_filtrada = base_completa[
        base_completa["Año"].isin(anios_base)
    ].copy()

    base_filtrada = filtrar_texto(
        base_filtrada,
        busqueda_general,
    )

    base_filtrada = base_filtrada.sort_values(
        ["Año", "Puesto"],
        ascending=[False, True],
    )

    st.write(
        f"La búsqueda contiene **{len(base_filtrada)} registros**."
    )

    st.dataframe(
        base_filtrada[
            [
                "Año",
                "Puesto",
                "Título",
                "Recaudación",
                "Dirección",
                "Productora",
                "País",
            ]
        ],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Recaudación": st.column_config.NumberColumn(
                "Recaudación",
                format="US$ %.0f",
            )
        },
    )

    st.download_button(
        "Descargar base consolidada",
        data=convertir_a_excel(base_filtrada),
        file_name="taquilla_peru_2015_2025.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


# SECCIÓN 4: INFORMACIÓN DEL PROYECTO


else:

    st.markdown(
        '<h2 class="seccion-titulo">Acerca de los datos</h2>',
        unsafe_allow_html=True,
    )

    st.markdown(
        """
        Esta página reúne los archivos anuales de películas más
        taquilleras en Perú entre 2015 y 2025.

        Cada película incluye:

        - título;
        - recaudación;
        - dirección;
        - productora;
        - país de producción.

        La aplicación lee directamente los Excel guardados en la carpeta
        `data`. Al cargar los archivos, limpia los textos, convierte las
        recaudaciones a números y añade las columnas de año y puesto.

        Los archivos originales no son modificados. Los filtros y las
        descargas se generan únicamente dentro de la aplicación.
        """
    )

    st.warning(
        "Antes de presentar esta base como una fuente académica o "
        "periodística, conviene verificar los rankings y las cifras con "
        "las fuentes originales de taquilla."
    )


# ---------------------------------------------------------
# Pie de página
# ---------------------------------------------------------

st.markdown(
    """
    <div class="pie">
        Proyecto de visualización de datos cinematográficos ·
        Perú, 2015–2025
    </div>
    """,
    unsafe_allow_html=True,
)
