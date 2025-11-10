import json
import copy
from io import BytesIO
from typing import Dict, Any, List, Tuple, Optional

import pandas as pd
import streamlit as st
import xml.etree.ElementTree as ET
from xml.dom import minidom


# ==========================
# Utilidades de negocio
# ==========================

def tiene_lista_con_items(servicios: Any) -> bool:
    """Retorna True si el diccionario 'servicios' tiene al menos una lista con 1 item."""
    if not isinstance(servicios, dict):
        return False
    for v in servicios.values():
        if isinstance(v, list) and len(v) > 0:
            return True
    return False


def ajustar_signo_servicios(servicios: Dict[str, Any], signo: int) -> None:
    """
    Multiplica por 'signo' algunos campos numéricos típicos de RIPS en todas las listas de servicios.
    Esto permite, por ejemplo, convertir una factura en nota crédito usando valores negativos.
    """
    for lista in servicios.values():
        if not isinstance(lista, list):
            continue
        for item in lista:
            if not isinstance(item, dict):
                continue
            for campo in ("vrServicio", "valorPagoModerador"):
                if campo in item and isinstance(item[campo], (int, float)):
                    item[campo] = item[campo] * signo


def copiar_servicios_factura_a_nota(
    factura: Dict[str, Any],
    nota: Dict[str, Any],
    forzar_signo: Optional[int] = None,
) -> Tuple[Dict[str, Any], Dict[str, Any]]:
    """
    Copia el bloque 'servicios' de la factura a la nota crédito.
    Busca cada usuario por tipo/numero de documento y, si falla, solo por número de documento.

    forzar_signo:
        None  -> deja los valores tal cual están en la factura.
        +1    -> fuerza a que queden positivos.
        -1    -> fuerza a que queden negativos.
    """
    inv_users = factura.get("usuarios", [])
    note_users = nota.get("usuarios", [])

    inv_map_full: Dict[Tuple[str, str], Dict[str, Any]] = {}
    inv_map_by_num: Dict[str, Dict[str, Any]] = {}

    for u in inv_users:
        tipo = u.get("tipoDocumentoIdentificacion")
        num = u.get("numDocumentoIdentificacion")
        servicios = u.get("servicios", {})
        if tipo is None or num is None:
            continue
        inv_map_full[(tipo, num)] = servicios
        inv_map_by_num[num] = servicios

    modificados = 0
    ya_tenian_servicios = 0
    sin_encontrar: List[Tuple[str, str]] = []

    for u in note_users:
        tipo = u.get("tipoDocumentoIdentificacion")
        num = u.get("numDocumentoIdentificacion")
        key_full = (tipo, num)

        servicios_actuales = u.get("servicios")

        # ¿Ya tiene listas con ítems?
        if tiene_lista_con_items(servicios_actuales):
            ya_tenian_servicios += 1
            continue

        servicios_origen = inv_map_full.get(key_full) or inv_map_by_num.get(num)

        if servicios_origen is None:
            sin_encontrar.append(key_full)
            continue

        nuevo_servicios = copy.deepcopy(servicios_origen)
        if forzar_signo in (1, -1):
            ajustar_signo_servicios(nuevo_servicios, forzar_signo)

        u["servicios"] = nuevo_servicios
        modificados += 1

    resumen = {
        "total_usuarios_factura": len(inv_users),
        "total_usuarios_nota": len(note_users),
        "usuarios_modificados": modificados,
        "usuarios_ya_tenian_servicios": ya_tenian_servicios,
        "usuarios_sin_encontrar": sin_encontrar,
    }
    return nota, resumen


def validar_estructura_servicios(nota: Dict[str, Any]) -> List[int]:
    """
    Devuelve la lista de índices de usuarios cuya estructura 'servicios'
    NO cumple con 'al menos una lista con 1 item'.
    """
    malos: List[int] = []
    for i, u in enumerate(nota.get("usuarios", [])):
        if not tiene_lista_con_items(u.get("servicios")):
            malos.append(i)
    return malos


def generar_resumen_usuarios(nota: Dict[str, Any]) -> pd.DataFrame:
    filas: List[Dict[str, Any]] = []
    for idx, u in enumerate(nota.get("usuarios", [])):
        servicios = u.get("servicios", {})
        tiene_serv = tiene_lista_con_items(servicios)
        num_listas = 0
        total_items = 0
        if isinstance(servicios, dict):
            for v in servicios.values():
                if isinstance(v, list):
                    num_listas += 1
                    total_items += len(v)
        filas.append(
            {
                "idx": idx,
                "tipoDocumentoIdentificacion": u.get("tipoDocumentoIdentificacion"),
                "numDocumentoIdentificacion": u.get("numDocumentoIdentificacion"),
                "estadoServicios": "OK" if tiene_serv else "INCOMPLETO",
                "numListasServicios": num_listas,
                "totalItemsServicios": total_items,
            }
        )
    return pd.DataFrame(filas)


# ==========================
# Excel masivo
# ==========================

def generar_excel_servicios(nota: Dict[str, Any]) -> BytesIO:
    """
    Genera un Excel con una fila por usuario y una columna 'servicios_json'
    que contiene el JSON de los servicios para edición masiva.
    """
    filas: List[Dict[str, Any]] = []
    for idx, u in enumerate(nota.get("usuarios", [])):
        filas.append(
            {
                "index": idx,
                "tipoDocumentoIdentificacion": u.get("tipoDocumentoIdentificacion"),
                "numDocumentoIdentificacion": u.get("numDocumentoIdentificacion"),
                "servicios_json": json.dumps(u.get("servicios", {}), ensure_ascii=False),
            }
        )
    df = pd.DataFrame(filas)
    buffer = BytesIO()
    # 👇 Aquí el cambio: usamos openpyxl para escribir el .xlsx
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="servicios")
    buffer.seek(0)
    return buffer


def aplicar_excel_servicios(nota: Dict[str, Any], archivo_excel) -> Tuple[Dict[str, Any], List[str]]:
    """
    Aplica los cambios de servicios contenidos en un Excel.
    Espera las columnas: index, servicios_json.
    """
    errores: List[str] = []
    try:
        df = pd.read_excel(archivo_excel)
    except Exception as exc:
        errores.append(f"No se pudo leer el Excel: {exc}")
        return nota, errores

    if "index" not in df.columns or "servicios_json" not in df.columns:
        errores.append("El Excel debe contener las columnas 'index' y 'servicios_json'.")
        return nota, errores

    usuarios = nota.get("usuarios", [])

    for _, fila in df.iterrows():
        try:
            idx = int(fila["index"])
        except Exception:
            errores.append(f"Índice inválido en fila: {fila.get('index')}")
            continue

        if not (0 <= idx < len(usuarios)):
            errores.append(f"Índice fuera de rango en fila: {idx}")
            continue

        servicios_str = fila.get("servicios_json")
        if pd.isna(servicios_str):
            # Nada que actualizar para este usuario
            continue

        try:
            servicios_nuevos = json.loads(servicios_str)
        except Exception as exc:
            errores.append(f"JSON inválido en fila índice {idx}: {exc}")
            continue

        usuarios[idx]["servicios"] = servicios_nuevos

    nota["usuarios"] = usuarios
    return nota, errores


# ==========================
# JSON -> XML (estructura genérica)
# ==========================

def nota_json_a_xml_element(nota: Dict[str, Any]) -> ET.Element:
    """
    Convierte la nota/factura JSON en un XML genérico de trabajo.
    NO es un XML oficial DIAN, es una representación estructurada
    para efectos de consulta y descarga.
    """
    root = ET.Element("RipsDocumento")
    for key, val in nota.items():
        if key == "usuarios":
            continue
        child = ET.SubElement(root, key)
        child.text = "" if val is None else str(val)

    usuarios_el = ET.SubElement(root, "usuarios")
    for u in nota.get("usuarios", []):
        u_el = ET.SubElement(usuarios_el, "usuario")
        for key, val in u.items():
            if key == "servicios":
                serv_el = ET.SubElement(u_el, "servicios")
                if isinstance(val, dict):
                    for tipo_serv, lista in val.items():
                        tipo_el = ET.SubElement(serv_el, str(tipo_serv))
                        if isinstance(lista, list):
                            for item in lista:
                                item_el = ET.SubElement(tipo_el, "item")
                                if isinstance(item, dict):
                                    for kk, vv in item.items():
                                        campo_el = ET.SubElement(item_el, str(kk))
                                        campo_el.text = "" if vv is None else str(vv)
                continue
            campo_el = ET.SubElement(u_el, str(key))
            campo_el.text = "" if val is None else str(val)
    return root


def nota_json_a_xml_bytes(nota: Dict[str, Any]) -> bytes:
    elem = nota_json_a_xml_element(nota)
    rough_xml = ET.tostring(elem, encoding="utf-8")
    dom = minidom.parseString(rough_xml)
    pretty = dom.toprettyxml(indent="  ", encoding="utf-8")
    return pretty


# ==========================
# Helpers de sesión
# ==========================

def cargar_json_en_estado(uploaded_file, state_key: str, name_key: str) -> None:
    if uploaded_file is None:
        return
    nombre_subido = uploaded_file.name
    nombre_actual = st.session_state.get(name_key)
    if nombre_actual == nombre_subido and state_key in st.session_state:
        # Ya lo tenemos cargado
        return
    try:
        data = json.load(uploaded_file)
    except Exception as exc:
        st.error(f"No se pudo leer el JSON '{nombre_subido}': {exc}")
        return
    st.session_state[state_key] = data
    st.session_state[name_key] = nombre_subido


def obtener_nota() -> Optional[Dict[str, Any]]:
    return st.session_state.get("nota_data")


def obtener_factura() -> Optional[Dict[str, Any]]:
    return st.session_state.get("factura_data")


# ==========================
# Interfaz Streamlit
# ==========================

def main():
    st.set_page_config(page_title="Asistente RIPS JSON / Notas Crédito", layout="wide")
    st.title("🧾 Asistente RIPS JSON / Notas Crédito")
    st.write(
        "Cargue la **factura (JSON completo)** y la **nota/crédito o archivo RIPS incompleto** en JSON. "
        "La aplicación le permitirá copiar los servicios faltantes, editar manualmente, "
        "hacer ajustes masivos por Excel y descargar el resultado en JSON y XML."
    )

    # ---- Sidebar: carga de archivos ----
    st.sidebar.header("1️⃣ Cargar archivos")

    factura_file = st.sidebar.file_uploader(
        "JSON de referencia (Factura completa)", type=["json"], key="factura_uploader"
    )
    nota_file = st.sidebar.file_uploader(
        "JSON a corregir (Nota crédito / RIPS)", type=["json"], key="nota_uploader"
    )
    excel_file = st.sidebar.file_uploader(
        "Plantilla Excel con servicios actualizados (opcional)", type=["xlsx"], key="excel_uploader"
    )

    if "factura_data" not in st.session_state:
        st.session_state["factura_data"] = None
        st.session_state["factura_name"] = None
    if "nota_data" not in st.session_state:
        st.session_state["nota_data"] = None
        st.session_state["nota_name"] = None

    # Cargar JSONs en sesión
    cargar_json_en_estado(factura_file, "factura_data", "factura_name")
    cargar_json_en_estado(nota_file, "nota_data", "nota_name")

    factura_data = obtener_factura()
    nota_data = obtener_nota()

    col_meta1, col_meta2 = st.columns(2)

    with col_meta1:
        st.subheader("📄 Factura / JSON de referencia")
        if factura_data:
            st.markdown(f"**Archivo:** `{st.session_state.get('factura_name')}`")
            st.json(
                {
                    "numDocumentoIdObligado": factura_data.get("numDocumentoIdObligado"),
                    "numFactura": factura_data.get("numFactura"),
                    "tipoNota": factura_data.get("tipoNota"),
                    "numNota": factura_data.get("numNota"),
                    "totalUsuarios": len(factura_data.get("usuarios", [])),
                }
            )
        else:
            st.info("Suba un JSON de factura completa en la barra lateral.")

    with col_meta2:
        st.subheader("🧾 Nota / JSON a corregir")
        if nota_data:
            st.markdown(f"**Archivo:** `{st.session_state.get('nota_name')}`")
            st.json(
                {
                    "numDocumentoIdObligado": nota_data.get("numDocumentoIdObligado"),
                    "numFactura": nota_data.get("numFactura"),
                    "tipoNota": nota_data.get("tipoNota"),
                    "numNota": nota_data.get("numNota"),
                    "totalUsuarios": len(nota_data.get("usuarios", [])),
                }
            )
        else:
            st.info("Suba el JSON de la nota/crédito o archivo RIPS incompleto.")

    if not nota_data:
        st.stop()

    # ---- Resumen de usuarios / validación ----
    st.markdown("---")
    st.subheader("2️⃣ Resumen y validación de usuarios")

    df_resumen = generar_resumen_usuarios(nota_data)
    if df_resumen.empty:
        st.warning("El JSON de la nota no contiene usuarios.")
    else:
        col_tabla, col_info = st.columns([3, 1])
        with col_tabla:
            st.dataframe(df_resumen, use_container_width=True, height=400)
        with col_info:
            total = len(df_resumen)
            incompletos = (df_resumen["estadoServicios"] == "INCOMPLETO").sum()
            st.metric("Usuarios totales", total)
            st.metric("Usuarios con servicios incompletos", incompletos)
            if incompletos == 0:
                st.success("Todos los usuarios tienen al menos una lista de servicios con 1 ítem.")
            else:
                st.error(
                    "Hay usuarios con 'servicios' vacío o sin listas con ítems. "
                    "Puede rellenarlos automáticamente desde la factura o editar manualmente."
                )

    # ---- Copiar servicios desde la factura ----
    st.markdown("---")
    st.subheader("3️⃣ Rellenar servicios desde JSON de referencia")

    if not factura_data:
        st.info("Para poder copiar servicios automáticamente, cargue primero el JSON de la factura en la barra lateral.")
    else:
        col_signo, col_boton = st.columns([2, 1])
        with col_signo:
            opcion_signo = st.selectbox(
                "Manejo del signo en `vrServicio` y `valorPagoModerador`:",
                (
                    "Dejar igual que la factura",
                    "Forzar valores POSITIVOS",
                    "Forzar valores NEGATIVOS",
                ),
            )
            signo = None
            if opcion_signo == "Forzar valores POSITIVOS":
                signo = 1
            elif opcion_signo == "Forzar valores NEGATIVOS":
                signo = -1
        with col_boton:
            if st.button("Rellenar servicios vacíos desde factura"):
                nota_trabajo = copy.deepcopy(nota_data)
                nota_actualizada, resumen = copiar_servicios_factura_a_nota(
                    factura_data, nota_trabajo, signo
                )
                st.session_state["nota_data"] = nota_actualizada
                nota_data = nota_actualizada  # actualizar referencia local
                st.success(
                    f"Servicios copiados. Usuarios modificados: {resumen['usuarios_modificados']}, "
                    f"ya tenían servicios: {resumen['usuarios_ya_tenian_servicios']}, "
                    f"sin coincidencia en factura: {len(resumen['usuarios_sin_encontrar'])}."
                )

                malos = validar_estructura_servicios(nota_data)
                if malos:
                    st.warning(
                        f"Aún hay {len(malos)} usuario(s) con servicios incompletos. "
                        f"Índices de ejemplo: {malos[:10]}"
                    )
                else:
                    st.success("Todos los usuarios cumplen la estructura mínima de servicios.")

                # Recalcular resumen
                df_resumen = generar_resumen_usuarios(nota_data)

    # ---- Edición individual ----
    st.markdown("---")
    st.subheader("4️⃣ Edición individual de servicios")

    usuarios = nota_data.get("usuarios", [])
    if not usuarios:
        st.warning("No hay usuarios para editar.")
    else:
        max_idx = len(usuarios) - 1
        idx_sel = st.number_input(
            "Seleccione el índice de usuario a editar",
            min_value=0,
            max_value=max_idx,
            value=0,
            step=1,
        )
        usuario = usuarios[idx_sel]
        st.write(
            f"Usuario índice **{idx_sel}** – "
            f"{usuario.get('tipoDocumentoIdentificacion')} {usuario.get('numDocumentoIdentificacion')}"
        )

        servicios_actuales_str = json.dumps(usuario.get("servicios", {}), ensure_ascii=False, indent=2)
        servicios_editados = st.text_area(
            "Edite el JSON de `servicios` para este usuario (estructura dict con listas).",
            value=servicios_actuales_str,
            height=300,
            key=f"servicios_usuario_{idx_sel}",
        )

        if st.button("Guardar cambios en este usuario"):
            try:
                servicios_nuevos = json.loads(servicios_editados)
            except json.JSONDecodeError as exc:
                st.error(f"El JSON de servicios no es válido: {exc}")
            else:
                usuario["servicios"] = servicios_nuevos
                usuarios[idx_sel] = usuario
                nota_data["usuarios"] = usuarios
                st.session_state["nota_data"] = nota_data
                st.success("Servicios actualizados correctamente para este usuario.")

    # ---- Edición masiva por Excel ----
    st.markdown("---")
    st.subheader("5️⃣ Edición masiva por Excel")

    col_descarga, col_subida = st.columns(2)

    with col_descarga:
        buffer_excel = generar_excel_servicios(nota_data)
        st.download_button(
            "⬇️ Descargar plantilla Excel desde la nota actual",
            data=buffer_excel,
            file_name="plantilla_servicios_rips.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    with col_subida:
        if excel_file is not None:
            if st.button("Aplicar cambios desde Excel cargado"):
                nota_actualizada, errores = aplicar_excel_servicios(nota_data, excel_file)
                st.session_state["nota_data"] = nota_actualizada
                nota_data = nota_actualizada
                if errores:
                    st.warning("Se aplicaron los cambios, pero hubo advertencias:")
                    for e in errores:
                        st.write(f"- {e}")
                else:
                    st.success("Cambios masivos aplicados correctamente desde el Excel.")

    # ---- Descarga de resultados ----
    st.markdown("---")
    st.subheader("6️⃣ Descargar JSON y XML resultantes")

    nota_json_bytes = json.dumps(nota_data, ensure_ascii=False, indent=2).encode("utf-8")
    nombre_nota_base = st.session_state.get("nota_name") or "nota_corregida"

    col_json, col_xml = st.columns(2)
    with col_json:
        st.download_button(
            "⬇️ Descargar JSON corregido",
            data=nota_json_bytes,
            file_name=f"{nombre_nota_base.rsplit('.', 1)[0]}_corregida.json",
            mime="application/json",
        )

    with col_xml:
        xml_bytes = nota_json_a_xml_bytes(nota_data)
        st.download_button(
            "⬇️ Descargar XML generado desde JSON",
            data=xml_bytes,
            file_name=f"{nombre_nota_base.rsplit('.', 1)[0]}.xml",
            mime="application/xml",
        )


if __name__ == "__main__":
    main()
