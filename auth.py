"""
auth.py — Módulo de autenticación para web_zaal_ia
- Login por clave única (sin nombre de usuario)
- Roles: admin / usuario
- Registro de actividad en SQLite
"""

import hashlib
import os
import sqlite3
import datetime
from pathlib import Path
import streamlit as st

DB_PATH = Path(__file__).resolve().parent / "usuarios.db"
CLAVE_ADMIN_DEFAULT = "admin3510"


# ==========================================================
# UTILIDADES DE BASE DE DATOS
# ==========================================================

def _get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def init_db():
    """Crea las tablas y el usuario admin por defecto si no existen."""
    with _get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS usuarios (
                id        INTEGER PRIMARY KEY AUTOINCREMENT,
                nombre    TEXT    NOT NULL,
                clave_hash TEXT   NOT NULL,
                clave_salt TEXT   NOT NULL,
                agencia   TEXT    NOT NULL,
                rol       TEXT    NOT NULL DEFAULT 'usuario'
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS actividad (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                usuario_id INTEGER,
                nombre     TEXT,
                agencia    TEXT,
                fase       TEXT,
                fecha_hora TEXT
            )
        """)
        conn.commit()

        # Crear admin por defecto si no existe ningún admin
        cur = conn.execute("SELECT id FROM usuarios WHERE rol = 'admin' LIMIT 1")
        if cur.fetchone() is None:
            salt, hsh = _hash_clave(CLAVE_ADMIN_DEFAULT)
            conn.execute(
                "INSERT INTO usuarios (nombre, clave_hash, clave_salt, agencia, rol) VALUES (?, ?, ?, ?, ?)",
                ("Administrador", hsh, salt, "Valencia", "admin")
            )
            conn.commit()


# ==========================================================
# HASH DE CLAVES
# ==========================================================

def _hash_clave(clave: str) -> tuple[str, str]:
    """Devuelve (salt_hex, hash_hex) para la clave dada."""
    salt = os.urandom(16)
    hsh = hashlib.pbkdf2_hmac("sha256", clave.encode(), salt, 260_000)
    return salt.hex(), hsh.hex()


def _verificar_clave(clave: str, salt_hex: str, hash_hex: str) -> bool:
    salt = bytes.fromhex(salt_hex)
    hsh = hashlib.pbkdf2_hmac("sha256", clave.encode(), salt, 260_000)
    return hsh.hex() == hash_hex


# ==========================================================
# OPERACIONES DE USUARIOS
# ==========================================================

def login_por_clave(clave: str) -> dict | None:
    """Busca un usuario por clave. Devuelve dict con datos o None."""
    with _get_conn() as conn:
        cur = conn.execute("SELECT id, nombre, clave_hash, clave_salt, agencia, rol FROM usuarios")
        for row in cur.fetchall():
            uid, nombre, clave_hash, clave_salt, agencia, rol = row
            if _verificar_clave(clave, clave_salt, clave_hash):
                return {"id": uid, "nombre": nombre, "agencia": agencia, "rol": rol}
    return None


def listar_usuarios() -> list[dict]:
    with _get_conn() as conn:
        cur = conn.execute("SELECT id, nombre, agencia, rol FROM usuarios ORDER BY rol DESC, nombre")
        return [{"id": r[0], "nombre": r[1], "agencia": r[2], "rol": r[3]} for r in cur.fetchall()]


def crear_usuario(nombre: str, clave: str, agencia: str, rol: str) -> str | None:
    """Crea usuario. Devuelve mensaje de error o None si todo fue bien."""
    nombre = nombre.strip()
    clave = clave.strip()
    if not nombre or not clave:
        return "El nombre y la clave no pueden estar vacíos."
    if len(clave) < 4:
        return "La clave debe tener al menos 4 caracteres."
    # Verificar clave duplicada
    if login_por_clave(clave) is not None:
        return "Esa clave ya está en uso por otro usuario."
    salt, hsh = _hash_clave(clave)
    with _get_conn() as conn:
        conn.execute(
            "INSERT INTO usuarios (nombre, clave_hash, clave_salt, agencia, rol) VALUES (?, ?, ?, ?, ?)",
            (nombre, hsh, salt, agencia, rol)
        )
        conn.commit()
    return None


def editar_usuario(uid: int, nombre: str, agencia: str, rol: str, nueva_clave: str = "") -> str | None:
    """Edita un usuario. Si nueva_clave está vacía, no cambia la clave."""
    nombre = nombre.strip()
    if not nombre:
        return "El nombre no puede estar vacío."
    with _get_conn() as conn:
        if nueva_clave.strip():
            clave = nueva_clave.strip()
            if len(clave) < 4:
                return "La clave debe tener al menos 4 caracteres."
            # Verificar que la nueva clave no la usa otro usuario
            existente = login_por_clave(clave)
            if existente is not None and existente["id"] != uid:
                return "Esa clave ya está en uso por otro usuario."
            salt, hsh = _hash_clave(clave)
            conn.execute(
                "UPDATE usuarios SET nombre=?, agencia=?, rol=?, clave_hash=?, clave_salt=? WHERE id=?",
                (nombre, agencia, rol, hsh, salt, uid)
            )
        else:
            conn.execute(
                "UPDATE usuarios SET nombre=?, agencia=?, rol=? WHERE id=?",
                (nombre, agencia, rol, uid)
            )
        conn.commit()
    return None


def eliminar_usuario(uid: int) -> str | None:
    """Elimina usuario. Protege contra borrar el último admin."""
    with _get_conn() as conn:
        cur = conn.execute("SELECT rol FROM usuarios WHERE id=?", (uid,))
        row = cur.fetchone()
        if row is None:
            return "Usuario no encontrado."
        if row[0] == "admin":
            cur2 = conn.execute("SELECT COUNT(*) FROM usuarios WHERE rol='admin'")
            if cur2.fetchone()[0] <= 1:
                return "No se puede eliminar el único administrador."
        conn.execute("DELETE FROM usuarios WHERE id=?", (uid,))
        conn.commit()
    return None


# ==========================================================
# REGISTRO DE ACTIVIDAD
# ==========================================================

def registrar_actividad(usuario_id: int, nombre: str, agencia: str, fase: str):
    fecha_hora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with _get_conn() as conn:
        conn.execute(
            "INSERT INTO actividad (usuario_id, nombre, agencia, fase, fecha_hora) VALUES (?, ?, ?, ?, ?)",
            (usuario_id, nombre, agencia, fase, fecha_hora)
        )
        conn.commit()


def listar_actividad(limit: int = 200) -> list[dict]:
    with _get_conn() as conn:
        cur = conn.execute(
            "SELECT nombre, agencia, fase, fecha_hora FROM actividad ORDER BY id DESC LIMIT ?",
            (limit,)
        )
        return [{"nombre": r[0], "agencia": r[1], "fase": r[2], "fecha_hora": r[3]} for r in cur.fetchall()]


# ==========================================================
# PANTALLA DE LOGIN
# ==========================================================

def render_login():
    """Muestra la pantalla de login. Actualiza st.session_state['usuario'] si correcto."""
    st.markdown(
        "<h1 style='text-align:center; margin-bottom:0.2em;'>ZAAL IA</h1>"
        "<p style='text-align:center; color:gray; margin-top:0;'>Reparto determinista</p>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    col_l, col_c, col_r = st.columns([1, 1.2, 1])
    with col_c:
        st.markdown("### Acceso")
        clave = st.text_input("Clave de acceso", type="password", key="_login_clave")
        if st.button("Entrar", use_container_width=True, type="primary"):
            if not clave:
                st.error("Introduce tu clave de acceso.")
            else:
                usuario = login_por_clave(clave)
                if usuario is None:
                    st.error("Clave incorrecta.")
                else:
                    st.session_state["usuario"] = usuario
                    st.rerun()


# ==========================================================
# PANEL DE GESTIÓN DE USUARIOS (solo admin)
# ==========================================================

def render_panel_admin():
    """Panel de administración: crear, editar y eliminar usuarios."""
    st.subheader("Gestión de usuarios")

    usuarios = listar_usuarios()

    # ── Crear nuevo usuario ──────────────────────────────────
    with st.expander("Crear nuevo usuario", expanded=False):
        with st.form("form_crear_usuario"):
            nuevo_nombre = st.text_input("Nombre")
            nueva_clave = st.text_input("Clave", type="password")
            nueva_agencia = st.selectbox("Agencia", ["Valencia", "Castellon"])
            nuevo_rol = st.selectbox("Rol", ["usuario", "admin"])
            submitted = st.form_submit_button("Crear usuario")
            if submitted:
                err = crear_usuario(nuevo_nombre, nueva_clave, nueva_agencia, nuevo_rol)
                if err:
                    st.error(err)
                else:
                    st.success(f"Usuario '{nuevo_nombre}' creado correctamente.")
                    st.rerun()

    st.markdown("---")
    st.markdown("#### Usuarios existentes")

    for u in usuarios:
        icono = "🔑" if u["rol"] == "admin" else "👤"
        label = f"{icono} **{u['nombre']}** — {u['agencia']} ({u['rol']})"
        with st.expander(label, expanded=False):
            with st.form(f"form_editar_{u['id']}"):
                e_nombre  = st.text_input("Nombre",  value=u["nombre"],  key=f"e_nom_{u['id']}")
                e_agencia = st.selectbox("Agencia", ["Valencia", "Castellon"],
                                         index=0 if u["agencia"] == "Valencia" else 1,
                                         key=f"e_age_{u['id']}")
                e_rol     = st.selectbox("Rol", ["usuario", "admin"],
                                          index=0 if u["rol"] == "usuario" else 1,
                                          key=f"e_rol_{u['id']}")
                e_clave   = st.text_input("Nueva clave (dejar vacío para no cambiar)",
                                           type="password", key=f"e_clv_{u['id']}")
                col_g, col_e = st.columns(2)
                with col_g:
                    guardar = st.form_submit_button("Guardar cambios")
                with col_e:
                    eliminar = st.form_submit_button("Eliminar usuario", type="secondary")

                if guardar:
                    err = editar_usuario(u["id"], e_nombre, e_agencia, e_rol, e_clave)
                    if err:
                        st.error(err)
                    else:
                        st.success("Usuario actualizado.")
                        st.rerun()

                if eliminar:
                    err = eliminar_usuario(u["id"])
                    if err:
                        st.error(err)
                    else:
                        st.success("Usuario eliminado.")
                        st.rerun()

    # ── Registro de actividad ────────────────────────────────
    st.markdown("---")
    with st.expander("Registro de actividad (últimas 200 entradas)", expanded=False):
        import pandas as pd
        actividad = listar_actividad()
        if actividad:
            st.dataframe(pd.DataFrame(actividad), use_container_width=True)
        else:
            st.info("Sin actividad registrada aún.")
