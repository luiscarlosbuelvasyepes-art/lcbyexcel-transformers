import os
from io import BytesIO

import pandas as pd
from flask import Flask, render_template, request, send_file, jsonify

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024

df_original = None
df_procesado = None
archivo_nombre = None


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok"}, 200


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")


@app.route("/cargar", methods=["POST"])
def cargar_archivo():
    global df_original, df_procesado, archivo_nombre
    
    try:
        if "archivo" not in request.files:
            return jsonify({"error": "No se seleccionó archivo"}), 400
        
        archivo = request.files["archivo"]
        if archivo.filename == "":
            return jsonify({"error": "Nombre de archivo vacío"}), 400
        
        if not archivo.filename.lower().endswith(".xlsx"):
            return jsonify({"error": "Solo se permiten archivos .xlsx"}), 400
        
        df = pd.read_excel(archivo)
        
        if df.shape[1] < 2:
            return jsonify({"error": "El archivo debe tener al menos dos columnas"}), 400
        
        df_original = df
        df_procesado = None
        archivo_nombre = archivo.filename
        
        # Convertir DataFrame a tabla HTML
        tabla_html = df.fillna("").to_html(classes="tabla-datos", index=False)
        
        return jsonify({
            "success": True,
            "nombre": archivo_nombre,
            "filas": len(df),
            "columnas": len(df.columns),
            "tabla": tabla_html
        })
    
    except Exception as error:
        return jsonify({"error": f"Error al cargar archivo: {str(error)}"}), 500


@app.route("/procesar", methods=["POST"])
def procesar():
    global df_original, df_procesado
    
    try:
        if df_original is None:
            return jsonify({"error": "Primero debes cargar un archivo"}), 400
        
        df = df_original.copy()
        
        # Eliminar filas completamente vacías
        df = df.dropna(how="all")
        
        if df.empty:
            return jsonify({"error": "No hay filas válidas para procesar"}), 400
        
        # Segunda columna por posición, reemplazar vacíos por 0
        nombre_segunda_columna = df.columns[1]
        segunda_columna_numerica = pd.to_numeric(df[nombre_segunda_columna], errors="coerce")
        segunda_columna_numerica = segunda_columna_numerica.fillna(0)
        df[nombre_segunda_columna] = segunda_columna_numerica
        
        # Calcular columna ACUM
        df["ACUM"] = segunda_columna_numerica.cumsum()
        
        df_procesado = df
        
        # Convertir a tabla HTML
        tabla_html = df.fillna("").to_html(classes="tabla-datos", index=False)
        
        return jsonify({
            "success": True,
            "mensaje": "Datos procesados correctamente",
            "filas": len(df),
            "columnas": len(df.columns),
            "tabla": tabla_html
        })
    
    except Exception as error:
        return jsonify({"error": f"Error al procesar: {str(error)}"}), 500


@app.route("/descargar", methods=["POST"])
def descargar():
    global df_procesado, archivo_nombre
    
    try:
        if df_procesado is None:
            return jsonify({"error": "Primero debes procesar los datos"}), 400
        
        # Generar nombre del archivo descargado
        if archivo_nombre:
            nombre_salida = archivo_nombre.replace(".xlsx", "_procesado.xlsx")
        else:
            nombre_salida = "resultado_procesado.xlsx"
        
        # Guardar a BytesIO
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_procesado.to_excel(writer, index=False, sheet_name="Datos")
        
        output.seek(0)
        
        return send_file(
            output,
            as_attachment=True,
            download_name=nombre_salida,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    except Exception as error:
        return jsonify({"error": f"Error al descargar: {str(error)}"}), 500


@app.errorhandler(413)
def archivo_muy_grande(_error):
    return jsonify({"error": "El archivo supera el tamaño permitido de 16 MB"}), 413


if __name__ == "__main__":
    puerto = int(os.environ.get("PORT", "5000"))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug, host="0.0.0.0", port=puerto)


    mejor_indice = 0
    mejor_puntaje = -1
    limite = min(len(df), 20)
    for indice in range(limite):
        puntaje = 0
        for valor in df.iloc[indice].tolist():
            clave = clave_texto(valor)
            if clave in ALIAS_METRICAS:
                puntaje += 4
            elif clave in ALIAS_PERIODO:
                puntaje += 3
            elif clave in ALIAS_ESTUDIANTE:
                puntaje += 3
            elif normalizar_texto(valor).upper() in PERIODOS:
                puntaje += 1
        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_indice = indice
    return mejor_indice


def preparar_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    df_limpio = df_raw.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    if df_limpio.empty:
        raise ValueError("El archivo no contiene datos utilizables.")

    fila_encabezado = detectar_fila_encabezado(df_limpio)
    fila = df_limpio.iloc[fila_encabezado].tolist()
    usados: set[str] = set()
    encabezados: list[str] = []

    for indice, valor in enumerate(fila, start=1):
        texto = normalizar_texto(valor) or f"COL_{indice}"
        encabezados.append(nombre_columna_unico(texto, usados))

    df_datos = df_limpio.iloc[fila_encabezado + 1 :].copy()
    df_datos.columns = encabezados
    df_datos = df_datos.dropna(how="all").reset_index(drop=True)

    if df_datos.empty:
        raise ValueError("No se encontraron filas de datos debajo del encabezado detectado.")

    return df_datos


def detectar_columna_periodo(df: pd.DataFrame) -> str:
    mejor_columna = ""
    mejor_puntaje = -1.0

    for columna in df.columns:
        nombre = clave_texto(columna)
        valores = df[columna].apply(normalizar_texto)
        valores = valores[valores != ""]
        if valores.empty:
            continue

        coincidencias = valores.apply(lambda valor: valor.upper() in PERIODOS).sum()
        puntaje = coincidencias / len(valores)
        if nombre in ALIAS_PERIODO:
            puntaje += 1

        if puntaje > mejor_puntaje:
            mejor_puntaje = puntaje
            mejor_columna = columna

    if not mejor_columna or mejor_puntaje <= 0:
        raise ValueError("No fue posible identificar la columna del periodo.")
    return mejor_columna


def detectar_columnas_estudiante(df: pd.DataFrame, columna_periodo: str) -> tuple[str, str | None]:
    candidatos: list[tuple[float, str]] = []

    for columna in df.columns:
        if columna == columna_periodo:
            continue

        nombre = clave_texto(columna)
        serie = df[columna].apply(normalizar_texto)
        valores = serie[serie != ""]
        if valores.empty:
            continue

        longitud_media = valores.map(len).mean()
        variedad = valores.nunique() / max(len(valores), 1)
        puntaje = variedad + min(longitud_media / 20, 1)

        proporcion_letras = valores.apply(contiene_letras).mean()
        proporcion_numerica = valores.apply(parece_numero).mean()
        puntaje += proporcion_letras * 2
        puntaje -= proporcion_numerica * 2

        if any(alias in nombre for alias in ALIAS_ESTUDIANTE):
            puntaje += 3
        if valores.apply(lambda valor: valor.upper() in PERIODOS).mean() > 0.5:
            puntaje = -1

        candidatos.append((puntaje, columna))

    candidatos.sort(reverse=True)
    if not candidatos or candidatos[0][0] <= 0:
        raise ValueError("No fue posible identificar la columna del estudiante.")

    principal = candidatos[0][1]
    secundario = None
    for puntaje, columna in candidatos[1:]:
        nombre = clave_texto(columna)
        if puntaje > 0.5 and any(alias in nombre for alias in ["id", "codigo", "documento"]):
            secundario = columna
            break
    return principal, secundario


def detectar_columnas_metricas(df: pd.DataFrame, ignoradas: set[str]) -> dict[str, str]:
    mapeo: dict[str, str] = {}
    for columna in df.columns:
        if columna in ignoradas:
            continue

        clave = clave_texto(columna)
        if clave in ALIAS_METRICAS:
            mapeo[ALIAS_METRICAS[clave]] = columna
            continue

        serie = df[columna].apply(normalizar_texto)
        no_vacios = serie[serie != ""]
        if no_vacios.empty:
            continue

        proporcion_numerica = no_vacios.apply(parece_numero).mean()
        if proporcion_numerica >= 0.35:
            nombre_visible = normalizar_texto(columna).upper()
            if nombre_visible and not clave.startswith("col_"):
                mapeo[nombre_visible] = columna
    return mapeo


def fusionar_columnas_auxiliares(df: pd.DataFrame, columna_periodo: str) -> pd.DataFrame:
    trabajo = df.copy()
    columnas = list(trabajo.columns)
    a_eliminar: list[str] = []

    periodos = trabajo[columna_periodo].apply(lambda valor: normalizar_texto(valor).upper())
    es_final = periodos == "FINAL"

    for indice, columna in enumerate(columnas):
        if indice == 0 or not es_columna_auxiliar(columna):
            continue

        destino = columnas[indice - 1]
        if es_columna_auxiliar(destino):
            continue

        destino_vacio = trabajo[destino].apply(lambda v: pd.isna(v) or normalizar_texto(v) == "")
        auxiliar_con_dato = ~trabajo[columna].apply(lambda v: pd.isna(v) or normalizar_texto(v) == "")

        mascara_final = es_final & destino_vacio & auxiliar_con_dato
        if mascara_final.any():
            trabajo.loc[mascara_final, destino] = trabajo.loc[mascara_final, columna]

        destino_vacio = trabajo[destino].apply(lambda v: pd.isna(v) or normalizar_texto(v) == "")
        mascara_general = destino_vacio & auxiliar_con_dato
        if mascara_general.any():
            trabajo.loc[mascara_general, destino] = trabajo.loc[mascara_general, columna]

        a_eliminar.append(columna)

    if a_eliminar:
        trabajo = trabajo.drop(columns=a_eliminar, errors="ignore")
    return trabajo


def construir_etiqueta_estudiante(
    fila: pd.Series,
    columna_principal: str,
    columna_secundaria: str | None,
) -> str:
    principal = normalizar_texto(fila.get(columna_principal, ""))
    secundario = normalizar_texto(fila.get(columna_secundaria, "")) if columna_secundaria else ""
    if principal and secundario and principal != secundario:
        return f"{principal} ({secundario})"
    return principal or secundario


def normalizar_tabla_fuente(df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    columna_periodo = detectar_columna_periodo(df)
    df = fusionar_columnas_auxiliares(df, columna_periodo)

    columna_periodo = detectar_columna_periodo(df)
    columna_estudiante, columna_id = detectar_columnas_estudiante(df, columna_periodo)

    ignoradas = {columna_periodo, columna_estudiante}
    if columna_id:
        ignoradas.add(columna_id)

    columnas_metricas = detectar_columnas_metricas(df, ignoradas)
    if not columnas_metricas:
        raise ValueError("No se detectaron columnas de metricas en el archivo.")

    metricas_ordenadas: list[str] = []
    for metrica in METRICAS_BASE:
        if metrica in columnas_metricas:
            metricas_ordenadas.append(metrica)
    for metrica in columnas_metricas.keys():
        if metrica not in metricas_ordenadas:
            metricas_ordenadas.append(metrica)

    trabajo = df.copy()
    trabajo[columna_estudiante] = trabajo[columna_estudiante].ffill()
    if columna_id:
        trabajo[columna_id] = trabajo[columna_id].ffill()

    trabajo["_periodo"] = trabajo[columna_periodo].apply(lambda v: normalizar_texto(v).upper())
    trabajo = trabajo[trabajo["_periodo"].isin(PERIODOS)].copy()
    if trabajo.empty:
        raise ValueError("No se encontraron filas con periodos 001, 002, 003 o FINAL.")

    trabajo["_estudiante"] = trabajo.apply(
        lambda fila: construir_etiqueta_estudiante(fila, columna_estudiante, columna_id),
        axis=1,
    )
    trabajo = trabajo[trabajo["_estudiante"].apply(es_etiqueta_estudiante_valida)].copy()
    if trabajo.empty:
        raise ValueError("No se detectaron estudiantes validos en el archivo.")

    columnas_finales = ["_estudiante", "_periodo"]
    renombrar: dict[str, str] = {}

    for metrica, columna_original in columnas_metricas.items():
        columnas_finales.append(columna_original)
        renombrar[columna_original] = metrica

    normalizado = trabajo[columnas_finales].rename(columns=renombrar).copy()
    for metrica in metricas_ordenadas:
        if metrica in normalizado.columns:
            normalizado[metrica] = pd.to_numeric(normalizado[metrica], errors="coerce")

    return normalizado, metricas_ordenadas


def formatear_valor(valor: object) -> object:
    if pd.isna(valor):
        return ""
    if isinstance(valor, float):
        if math.isclose(valor, round(valor)):
            return int(round(valor))
        return round(valor, 2)
    return valor


def crear_matriz_estudiante(df_estudiante: pd.DataFrame, metricas: list[str], periodos: list[str]) -> pd.DataFrame:
    matriz = pd.DataFrame(index=metricas, columns=periodos, dtype=object).fillna("")

    for _, fila in df_estudiante.iterrows():
        periodo = fila["_periodo"]
        if periodo not in periodos:
            continue
        for metrica in metricas:
            if metrica in df_estudiante.columns:
                valor = fila.get(metrica)
                if pd.notna(valor):
                    matriz.loc[metrica, periodo] = formatear_valor(valor)
    return matriz


def etiqueta_nivelacion(promedio: float | None) -> str:
    if promedio is None or pd.isna(promedio):
        return "SIN NOTA"
    if promedio >= 4.6:
        return "SUPERIOR"
    if promedio >= 4.0:
        return "ALTO"
    if promedio >= 3.0:
        return "BASICO"
    return "BAJO"


def aplicar_logica_nivelacion(matriz: pd.DataFrame, periodos: list[str]) -> pd.DataFrame:
    resultado = matriz.copy()
    if "PROM" not in resultado.index:
        return resultado

    niveles: dict[str, str] = {}
    requiere: dict[str, str] = {}
    for periodo in periodos:
        valor = resultado.loc["PROM", periodo] if periodo in resultado.columns else ""
        try:
            promedio = float(valor)
        except (TypeError, ValueError):
            promedio = None

        nivel = etiqueta_nivelacion(promedio)
        niveles[periodo] = nivel
        requiere[periodo] = "SI" if nivel == "BAJO" else "NO"

    resultado.loc["NIVELACION"] = pd.Series(niveles)
    resultado.loc["REQUIERE_NIVELACION"] = pd.Series(requiere)
    return resultado


def nombre_hoja_seguro(nombre: str, usadas: set[str]) -> str:
    limpio = re.sub(r"[\\/*?:\[\]]", "_", nombre).strip()
    if not limpio:
        limpio = "Estudiante"
    base = limpio[:31]
    candidato = base
    i = 2
    while candidato in usadas:
        sufijo = f"_{i}"
        candidato = f"{base[:31 - len(sufijo)]}{sufijo}"
        i += 1
    usadas.add(candidato)
    return candidato


def construir_excel_reporte(reportes: dict[str, pd.DataFrame], periodos: list[str]) -> BytesIO:
    libro = Workbook()
    hoja_resumen = libro.active
    hoja_resumen.title = "Resumen"

    hoja_resumen.append(["Estudiante", "Periodo", "PROM", "Nivelacion", "Requiere Nivelacion"])
    usadas: set[str] = {"Resumen"}

    for estudiante, matriz in reportes.items():
        hoja = libro.create_sheet(title=nombre_hoja_seguro(estudiante, usadas))
        hoja.append(["METRICA", *periodos])
        for metrica in matriz.index.tolist():
            hoja.append([metrica, *[matriz.loc[metrica, p] if p in matriz.columns else "" for p in periodos]])

        for periodo in periodos:
            prom = matriz.loc["PROM", periodo] if "PROM" in matriz.index and periodo in matriz.columns else ""
            niv = matriz.loc["NIVELACION", periodo] if "NIVELACION" in matriz.index and periodo in matriz.columns else ""
            req = (
                matriz.loc["REQUIERE_NIVELACION", periodo]
                if "REQUIERE_NIVELACION" in matriz.index and periodo in matriz.columns
                else ""
            )
            hoja_resumen.append([estudiante, periodo, prom, niv, req])

    salida = BytesIO()
    libro.save(salida)
    salida.seek(0)
    return salida


def generar_reportes_desde_excel(file_storage, periodos: list[str]) -> BytesIO:
    if not periodos:
        raise ValueError("Debes seleccionar al menos un periodo.")

    df_raw = pd.read_excel(file_storage, header=None, dtype=object)
    df_datos = preparar_dataframe(df_raw)
    df_normalizado, metricas = normalizar_tabla_fuente(df_datos)

    reportes: dict[str, pd.DataFrame] = {}
    for estudiante, grupo in df_normalizado.groupby("_estudiante", sort=True):
        matriz = crear_matriz_estudiante(grupo, metricas, periodos)
        matriz = aplicar_logica_nivelacion(matriz, periodos)
        reportes[estudiante] = matriz

    if not reportes:
        raise ValueError("No se generaron reportes individuales.")

    return construir_excel_reporte(reportes, periodos)


@app.route("/", methods=["GET"])
def index():
    return render_template("index.html", periodos=PERIODOS, periodos_default=["001"])


@app.route("/generar", methods=["POST"])
def generar():
    archivo = request.files.get("archivo_excel")
    periodos = request.form.getlist("periodos")

    if archivo is None or archivo.filename == "":
        return render_template(
            "index.html",
            periodos=PERIODOS,
            periodos_default=periodos or ["001"],
            error="Debes cargar un archivo Excel (.xlsx).",
        ), 400

    if not archivo.filename.lower().endswith(".xlsx"):
        return render_template(
            "index.html",
            periodos=PERIODOS,
            periodos_default=periodos or ["001"],
            error="Formato invalido. Solo se permite archivo .xlsx",
        ), 400

    try:
        stream = generar_reportes_desde_excel(archivo, periodos)
    except Exception as error:
        return render_template(
            "index.html",
            periodos=PERIODOS,
            periodos_default=periodos or ["001"],
            error=f"No fue posible generar el reporte: {error}",
        ), 400

    marca_tiempo = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre = f"reporte_boletines_{marca_tiempo}.xlsx"
    return send_file(
        stream,
        as_attachment=True,
        download_name=nombre,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    puerto = int(os.environ.get("PORT", "5000"))
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(debug=debug, host="0.0.0.0", port=puerto)